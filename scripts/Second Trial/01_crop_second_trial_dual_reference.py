from __future__ import annotations

import argparse
import csv
import json
import re
from pathlib import Path

import cv2
import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
import numpy as np
import tifffile
from PIL import Image, ImageOps


# ============================================================
# 1) CHANGE ONLY THESE PATHS WHEN REUSING ON ANOTHER COMPUTER
# ============================================================

PROJECT_ROOT = Path(
    r"C:\Users\tshib\OneDrive\Desktop\Internship"
)

INPUT_ROOT = PROJECT_ROOT / "data" / "Second Trial"

OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Second Trial"
    / "01_Crop_Dual_Reference"
)


# ============================================================
# 2) PROJECT RULES
# ============================================================

MS_BANDS = ["MS_G", "MS_R", "MS_RE", "MS_NIR"]

DAY_NAME_TO_ORDER = {
    "day 1": 1,
    "day 2": 2,
    "day 3": 3,
    "day 4": 4,
    "day 5": 5,
    "day 9": 9,
    "first day": 1,
    "second day": 2,
    "third day": 3,
    "fourth day": 4,
    "fifth day": 5,
    "ninth day": 9,
}

BAND_OPTIONS = {"ALL", "D", "MS"}


# ============================================================
# 3) HELPERS
# ============================================================

def natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", text)
    ]


def day_sort_key(folder: Path):
    return (
        DAY_NAME_TO_ORDER.get(folder.name.casefold(), 999),
        natural_key(folder.name),
    )


def parse_filter_list(value: str | None):
    if not value:
        return None

    return {
        item.strip().casefold()
        for item in value.split(",")
        if item.strip()
    }


def relative_path(path: Path, root: Path) -> str:
    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


# ============================================================
# 4) FILE CLASSIFICATION
# ============================================================

def parse_image_metadata(path: Path):
    """
    Supports names like:
        DJI_20260627101229_0001_D.JPG
        DJI_20260627101230_0001_MS_NIR.TIF

    Key idea:
    We group by the logical tray image number (e.g. 0001),
    not only by the exact timestamped capture ID.
    """

    stem = path.stem.upper()

    match = re.match(
        r"^(?P<prefix>.+?)_(?P<seq>\d{4})_(?P<band>MS_NIR|MS_RE|MS_R|MS_G|D|F)$",
        stem,
    )

    if match:
        prefix = match.group("prefix")
        seq = match.group("seq")
        band = match.group("band")
        capture_id = f"{prefix}_{seq}"
        logical_key = seq

        return {
            "capture_id": capture_id,
            "logical_key": logical_key,
            "sequence_id": seq,
            "band": band,
        }

    # fallback for unusual names
    for band in ["MS_NIR", "MS_RE", "MS_R", "MS_G", "D", "F"]:
        if re.search(rf"(?:^|_){re.escape(band)}$", stem):
            capture_id = re.sub(rf"_{re.escape(band)}$", "", stem)
            return {
                "capture_id": capture_id,
                "logical_key": capture_id,
                "sequence_id": "",
                "band": band,
            }

    return {
        "capture_id": stem,
        "logical_key": stem,
        "sequence_id": "",
        "band": "UNKNOWN",
    }


def choose_canonical_capture_id(bands_dict: dict):
    """
    Prefer the D capture_id as the output/canonical ID.
    If D is not present, prefer MS_NIR, then any available band.
    """

    preference = ["D", "MS_NIR", "MS_G", "MS_R", "MS_RE"]

    for band in preference:
        if band in bands_dict:
            return bands_dict[band]["capture_id"]

    return "UNKNOWN_CAPTURE"


def find_capture_sets(tray_folder: Path):
    """
    Returns grouped capture sets using logical_key.

    Output structure:
    [
        {
            "logical_key": "0001",
            "canonical_capture_id": "DJI_..._0001",
            "bands": {
                "D": {"path": ..., "capture_id": ...},
                ...
            },
            "capture_ids_seen": [...]
        },
        ...
    ]
    """

    grouped = {}

    files = sorted(
        [
            file
            for file in tray_folder.rglob("*")
            if file.is_file()
            and file.suffix.casefold() in {
                ".jpg", ".jpeg", ".tif", ".tiff"
            }
        ],
        key=lambda p: natural_key(p.name),
    )

    for file in files:
        meta = parse_image_metadata(file)
        band = meta["band"]

        if band in {"UNKNOWN", "F"}:
            continue

        logical_key = meta["logical_key"]

        grouped.setdefault(logical_key, {
            "logical_key": logical_key,
            "bands": {},
            "capture_ids_seen": set(),
        })

        # Keep the first file found for a band if duplicates exist
        if band not in grouped[logical_key]["bands"]:
            grouped[logical_key]["bands"][band] = {
                "path": file,
                "capture_id": meta["capture_id"],
            }

        grouped[logical_key]["capture_ids_seen"].add(meta["capture_id"])

    results = []

    for logical_key, info in grouped.items():
        bands = info["bands"]
        canonical_capture_id = choose_canonical_capture_id(bands)

        results.append(
            {
                "logical_key": logical_key,
                "canonical_capture_id": canonical_capture_id,
                "bands": bands,
                "capture_ids_seen": sorted(info["capture_ids_seen"]),
            }
        )

    results.sort(key=lambda item: natural_key(item["canonical_capture_id"]))
    return results


# ============================================================
# 5) IMAGE READING
# ============================================================

def read_rgb(path: Path) -> np.ndarray:
    with Image.open(path) as image:
        image = ImageOps.exif_transpose(image)
        image = image.convert("RGB")
        return np.asarray(image)


def read_single_band_tif(path: Path) -> np.ndarray:
    array = tifffile.imread(path)

    if array.ndim == 3:
        if array.shape[0] == 1:
            array = array[0]
        elif array.shape[-1] == 1:
            array = array[:, :, 0]
        else:
            raise ValueError(f"{path.name} is not a single-band TIFF.")

    if array.ndim != 2:
        raise ValueError(f"{path.name} is not a 2D single-band TIFF.")

    return array


def ms_preview(array: np.ndarray) -> np.ndarray:
    values = array.astype(np.float32)
    low, high = np.percentile(values, [1, 99])

    if high <= low:
        high = low + 1

    preview = np.clip(
        (values - low) * 255 / (high - low),
        0,
        255,
    ).astype(np.uint8)

    return preview


# ============================================================
# 6) CLICKING THE CORNERS
# ============================================================

def get_four_corners(image, title: str, grayscale: bool = False):
    """
    Left click 4 corners in this order:
    1) top-left
    2) top-right
    3) bottom-right
    4) bottom-left

    Right click removes the last point.
    """

    fig, ax = plt.subplots(figsize=(15, 10))

    if grayscale:
        ax.imshow(image, cmap="gray")
    else:
        ax.imshow(image)

    ax.set_title(
        title
        + "\nClick: Top-left → Top-right → Bottom-right → Bottom-left"
        + "\nRight click removes the last point. Close the window to cancel.",
        fontsize=12,
    )
    ax.axis("off")
    plt.tight_layout()
    plt.show(block=False)

    points = plt.ginput(
        n=4,
        timeout=0,
        show_clicks=True,
        mouse_add=1,
        mouse_pop=3,
        mouse_stop=2,
    )

    plt.close(fig)

    if len(points) != 4:
        return None

    return np.asarray(points, dtype=np.float32)


# ============================================================
# 7) PERSPECTIVE CROP
# ============================================================

def calculate_crop_size(points: np.ndarray):
    top_width = np.linalg.norm(points[1] - points[0])
    bottom_width = np.linalg.norm(points[2] - points[3])

    left_height = np.linalg.norm(points[3] - points[0])
    right_height = np.linalg.norm(points[2] - points[1])

    width = max(20, int(round(max(top_width, bottom_width))))
    height = max(20, int(round(max(left_height, right_height))))

    return width, height


def make_transform(points: np.ndarray):
    width, height = calculate_crop_size(points)

    destination = np.asarray(
        [
            [0, 0],
            [width - 1, 0],
            [width - 1, height - 1],
            [0, height - 1],
        ],
        dtype=np.float32,
    )

    matrix = cv2.getPerspectiveTransform(
        points.astype(np.float32),
        destination,
    )

    return matrix, width, height


def crop_with_transform(
    image: np.ndarray,
    matrix: np.ndarray,
    width: int,
    height: int,
    interpolation: int,
):
    return cv2.warpPerspective(
        image,
        matrix,
        (width, height),
        flags=interpolation,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=0,
    )


# ============================================================
# 8) CONFIG / MANIFEST
# ============================================================

def load_json(path: Path):
    if not path.exists():
        return {"records": {}}

    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {"records": {}}


def save_json(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", encoding="utf-8") as file:
        json.dump(data, file, indent=2)


def expected_outputs(output_folder: Path, capture_id: str, process_bands: str):
    if process_bands == "D":
        return [output_folder / f"{capture_id}_D.JPG"]

    if process_bands == "MS":
        return [
            output_folder / f"{capture_id}_MS_G.TIF",
            output_folder / f"{capture_id}_MS_R.TIF",
            output_folder / f"{capture_id}_MS_RE.TIF",
            output_folder / f"{capture_id}_MS_NIR.TIF",
        ]

    return [
        output_folder / f"{capture_id}_D.JPG",
        output_folder / f"{capture_id}_MS_G.TIF",
        output_folder / f"{capture_id}_MS_R.TIF",
        output_folder / f"{capture_id}_MS_RE.TIF",
        output_folder / f"{capture_id}_MS_NIR.TIF",
    ]


def outputs_exist(output_folder: Path, capture_id: str, process_bands: str):
    return all(path.exists() for path in expected_outputs(output_folder, capture_id, process_bands))


def write_manifest(path: Path, rows: list[dict]):
    fields = [
        "day_order",
        "day",
        "tray",
        "logical_key",
        "capture_id",
        "capture_ids_seen",
        "status",
        "process_bands",
        "source_d",
        "source_ms_g",
        "source_ms_r",
        "source_ms_re",
        "source_ms_nir",
        "output_d",
        "output_ms_g",
        "output_ms_r",
        "output_ms_re",
        "output_ms_nir",
        "d_crop_width",
        "d_crop_height",
        "ms_crop_width",
        "ms_crop_height",
        "notes",
    ]

    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


# ============================================================
# 9) MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "Script 01: Crop Second Trial D/RGB and multispectral images "
            "using separate four-corner references."
        )
    )

    parser.add_argument(
        "--days",
        help='Example: --days "Day 1,Day 9"',
    )

    parser.add_argument(
        "--trays",
        help='Example: --trays "Tray 1,Tray 5"',
    )

    parser.add_argument(
        "--process-bands",
        default="ALL",
        choices=sorted(BAND_OPTIONS),
        help="ALL = crop D + multispectral; D = crop D only; MS = crop multispectral only.",
    )

    parser.add_argument(
        "--redo",
        action="store_true",
        help="Redo selected crop outputs even if they already exist.",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List detected image sets without cropping.",
    )

    args = parser.parse_args()

    process_bands = args.process_bands.upper()

    if not INPUT_ROOT.exists():
        print(f"ERROR: Input folder not found:\n{INPUT_ROOT}")
        return 1

    selected_days = parse_filter_list(args.days)
    selected_trays = parse_filter_list(args.trays)

    day_folders = sorted(
        [
            folder
            for folder in INPUT_ROOT.iterdir()
            if folder.is_dir()
            and folder.name.casefold() in DAY_NAME_TO_ORDER
        ],
        key=day_sort_key,
    )

    jobs = []

    for day_folder in day_folders:
        if selected_days and day_folder.name.casefold() not in selected_days:
            continue

        tray_folders = sorted(
            [folder for folder in day_folder.iterdir() if folder.is_dir()],
            key=lambda folder: natural_key(folder.name),
        )

        for tray_folder in tray_folders:
            if selected_trays and tray_folder.name.casefold() not in selected_trays:
                continue

            capture_sets = find_capture_sets(tray_folder)

            for capture_set in capture_sets:
                bands = capture_set["bands"]

                if process_bands == "ALL":
                    missing = [band for band in ["D", *MS_BANDS] if band not in bands]
                elif process_bands == "D":
                    missing = [band for band in ["D"] if band not in bands]
                else:  # MS
                    missing = [band for band in MS_BANDS if band not in bands]

                jobs.append(
                    {
                        "day_folder": day_folder,
                        "tray_folder": tray_folder,
                        "logical_key": capture_set["logical_key"],
                        "capture_id": capture_set["canonical_capture_id"],
                        "bands": bands,
                        "capture_ids_seen": capture_set["capture_ids_seen"],
                        "missing": missing,
                    }
                )

    if not jobs:
        print("No valid image sets were found.")
        return 1

    print("\nSCRIPT 01 — DUAL-REFERENCE CROPPING")
    print("=" * 70)
    print(f"Input folder:\n{INPUT_ROOT}")
    print(f"\nOutput folder:\n{OUTPUT_ROOT}")
    print(f"\nProcessing mode: {process_bands}")
    print("\nF preview images are ignored.\n")

    for job in jobs:
        if job["missing"]:
            state = "INCOMPLETE: " + ", ".join(job["missing"])
        else:
            state = "READY"

        id_note = ""
        if len(job["capture_ids_seen"]) > 1:
            id_note = " | merged timestamp set"

        print(
            f"{job['day_folder'].name} > {job['tray_folder'].name} > "
            f"{job['capture_id']} : {state}{id_note}"
        )

    if args.dry_run:
        print("\nDry run complete. No crops created.")
        return 0

    config_path = OUTPUT_ROOT / "_config" / "crop_points.json"
    manifest_path = OUTPUT_ROOT / "_reports" / "crop_manifest.csv"

    config = load_json(config_path)
    manifest_rows = []

    for job in jobs:
        day_name = job["day_folder"].name
        tray_name = job["tray_folder"].name
        logical_key = job["logical_key"]
        capture_id = job["capture_id"]
        bands = job["bands"]
        capture_ids_seen = job["capture_ids_seen"]
        missing = job["missing"]

        output_folder = OUTPUT_ROOT / day_name / tray_name
        output_folder.mkdir(parents=True, exist_ok=True)

        base_row = {
            "day_order": DAY_NAME_TO_ORDER.get(day_name.casefold(), 999),
            "day": day_name,
            "tray": tray_name,
            "logical_key": logical_key,
            "capture_id": capture_id,
            "capture_ids_seen": "; ".join(capture_ids_seen),
            "process_bands": process_bands,
            "source_d": relative_path(bands["D"]["path"], INPUT_ROOT) if "D" in bands else "",
            "source_ms_g": relative_path(bands["MS_G"]["path"], INPUT_ROOT) if "MS_G" in bands else "",
            "source_ms_r": relative_path(bands["MS_R"]["path"], INPUT_ROOT) if "MS_R" in bands else "",
            "source_ms_re": relative_path(bands["MS_RE"]["path"], INPUT_ROOT) if "MS_RE" in bands else "",
            "source_ms_nir": relative_path(bands["MS_NIR"]["path"], INPUT_ROOT) if "MS_NIR" in bands else "",
            "output_d": str(output_folder / f"{capture_id}_D.JPG"),
            "output_ms_g": str(output_folder / f"{capture_id}_MS_G.TIF"),
            "output_ms_r": str(output_folder / f"{capture_id}_MS_R.TIF"),
            "output_ms_re": str(output_folder / f"{capture_id}_MS_RE.TIF"),
            "output_ms_nir": str(output_folder / f"{capture_id}_MS_NIR.TIF"),
            "d_crop_width": "",
            "d_crop_height": "",
            "ms_crop_width": "",
            "ms_crop_height": "",
        }

        if missing:
            manifest_rows.append(
                {
                    **base_row,
                    "status": "SKIPPED_INCOMPLETE",
                    "notes": "Missing required band(s): " + ", ".join(missing),
                }
            )

            print(f"SKIPPED: {day_name} > {tray_name} > {capture_id}")
            continue

        if outputs_exist(output_folder, capture_id, process_bands) and not args.redo:
            manifest_rows.append(
                {
                    **base_row,
                    "status": "SKIPPED_EXISTING",
                    "notes": "Requested output(s) already exist. Use --redo to overwrite them.",
                }
            )

            print(f"SKIPPED EXISTING: {day_name} > {tray_name} > {capture_id}")
            continue

        try:
            d_width = d_height = ""
            ms_width = ms_height = ""

            # ------------------------------------------------
            # D CROP
            # ------------------------------------------------
            if process_bands in {"ALL", "D"}:
                d_rgb = read_rgb(bands["D"]["path"])

                d_points = get_four_corners(
                    d_rgb,
                    f"D/RGB | {day_name} | {tray_name} | {capture_id}",
                    grayscale=False,
                )

                if d_points is None:
                    manifest_rows.append(
                        {
                            **base_row,
                            "status": "CANCELLED",
                            "notes": "D crop cancelled by user.",
                        }
                    )
                    print(f"CANCELLED: {day_name} > {tray_name} > {capture_id}")
                    continue

                d_matrix, d_width, d_height = make_transform(d_points)

                d_crop = crop_with_transform(
                    d_rgb,
                    d_matrix,
                    d_width,
                    d_height,
                    cv2.INTER_LINEAR,
                )

                Image.fromarray(d_crop).save(
                    output_folder / f"{capture_id}_D.JPG",
                    quality=100,
                    subsampling=0,
                )
            else:
                d_points = None

            # ------------------------------------------------
            # MS CROP
            # ------------------------------------------------
            if process_bands in {"ALL", "MS"}:
                ms_nir = read_single_band_tif(bands["MS_NIR"]["path"])

                ms_points = get_four_corners(
                    ms_preview(ms_nir),
                    f"MS_NIR | {day_name} | {tray_name} | {capture_id}",
                    grayscale=True,
                )

                if ms_points is None:
                    status = "CANCELLED" if process_bands == "MS" else "PARTIAL_CANCELLED"
                    note = "MS crop cancelled by user." if process_bands == "MS" else "D crop saved, but MS crop cancelled."

                    manifest_rows.append(
                        {
                            **base_row,
                            "status": status,
                            "d_crop_width": d_width,
                            "d_crop_height": d_height,
                            "notes": note,
                        }
                    )
                    print(f"{status}: {day_name} > {tray_name} > {capture_id}")
                    continue

                ms_matrix, ms_width, ms_height = make_transform(ms_points)

                for band in MS_BANDS:
                    array = read_single_band_tif(bands[band]["path"])

                    cropped = crop_with_transform(
                        array,
                        ms_matrix,
                        ms_width,
                        ms_height,
                        cv2.INTER_NEAREST,
                    )

                    tifffile.imwrite(
                        output_folder / f"{capture_id}_{band}.TIF",
                        cropped,
                    )
            else:
                ms_points = None

            config_key = f"{day_name}|{tray_name}|{logical_key}"

            existing_record = config["records"].get(config_key, {})

            if d_points is not None:
                existing_record["d_points"] = d_points.tolist()
                existing_record["d_crop_size"] = [d_width, d_height]

            if ms_points is not None:
                existing_record["ms_points"] = ms_points.tolist()
                existing_record["ms_crop_size"] = [ms_width, ms_height]

            existing_record["day"] = day_name
            existing_record["tray"] = tray_name
            existing_record["logical_key"] = logical_key
            existing_record["capture_id"] = capture_id
            existing_record["capture_ids_seen"] = capture_ids_seen

            config["records"][config_key] = existing_record
            save_json(config_path, config)

            manifest_rows.append(
                {
                    **base_row,
                    "status": "PASS",
                    "d_crop_width": d_width,
                    "d_crop_height": d_height,
                    "ms_crop_width": ms_width,
                    "ms_crop_height": ms_height,
                    "notes": (
                        "Cropped successfully. "
                        "Timestamp mismatch across bands was grouped logically by image sequence."
                        if len(capture_ids_seen) > 1
                        else "Cropped successfully."
                    ),
                }
            )

            print(f"PASS: {day_name} > {tray_name} > {capture_id}")

        except Exception as error:
            manifest_rows.append(
                {
                    **base_row,
                    "status": "FAIL",
                    "notes": str(error),
                }
            )

            print(f"FAIL: {day_name} > {tray_name} > {capture_id} | {error}")

    manifest_rows.sort(
        key=lambda row: (
            row["day_order"],
            natural_key(row["tray"]),
            natural_key(row["capture_id"]),
        )
    )

    write_manifest(manifest_path, manifest_rows)

    pass_count = sum(row["status"] == "PASS" for row in manifest_rows)
    fail_count = sum(row["status"] == "FAIL" for row in manifest_rows)

    print("\n" + "=" * 70)
    print("SCRIPT 01 FINISHED")
    print(f"PASS: {pass_count}")
    print(f"FAIL: {fail_count}")
    print(f"\nManifest:\n{manifest_path}")
    print(f"\nSaved crop points:\n{config_path}")

    return 0 if fail_count == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())