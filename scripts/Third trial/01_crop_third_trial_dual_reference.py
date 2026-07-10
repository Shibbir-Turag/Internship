from __future__ import annotations

"""
SCRIPT 01 — THIRD TRIAL DUAL-REFERENCE CROPPING

Purpose
-------
Crop Third Trial D/RGB and multispectral tray images using separate crop
references.

This script is adapted from the Second Trial Script 01 workflow but updated for
Third Trial.

It supports:
- Day 1 to Day 7
- Tray 1 to Tray 12
- D/RGB images
- MS_G, MS_R, MS_RE, MS_NIR images
- F preview images ignored
- Raw images stored either:
    A) directly inside each Day folder
    B) inside Day/Tray subfolders

Important
---------
D/RGB and MS images are cropped separately because their geometry can differ.

For each tray image set:
1. D crop uses the D/RGB image as its reference.
2. MS crop uses MS_NIR as the MS reference.
3. The MS_NIR crop transform is applied to MS_G, MS_R, MS_RE and MS_NIR.

Output
------
outputs/Third trial/01_Crop_Dual_Reference/
    Day X/
        Tray Y/
            <capture_id>_D.JPG
            <capture_id>_MS_G.TIF
            <capture_id>_MS_R.TIF
            <capture_id>_MS_RE.TIF
            <capture_id>_MS_NIR.TIF

    _config/
        crop_points.json

    _reports/
        crop_manifest.csv
"""

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
# 1) PATHS — CHANGE ONLY THESE WHEN REUSING
# ============================================================

PROJECT_ROOT = Path(
    r"C:\Users\tshib\OneDrive\Desktop\Internship"
)

INPUT_ROOT = (
    PROJECT_ROOT
    / "data"
    / "Third Trial"
)

OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "01_Crop_Dual_Reference"
)


# ============================================================
# 2) TRIAL SETTINGS
# ============================================================

MS_BANDS = [
    "MS_G",
    "MS_R",
    "MS_RE",
    "MS_NIR",
]

ALL_REQUIRED_BANDS = [
    "D",
    *MS_BANDS,
]

DAY_NAME_TO_ORDER = {
    "day 1": 1,
    "day 2": 2,
    "day 3": 3,
    "day 4": 4,
    "day 5": 5,
    "day 6": 6,
    "day 7": 7,
    "first day": 1,
    "second day": 2,
    "third day": 3,
    "fourth day": 4,
    "fifth day": 5,
    "sixth day": 6,
    "seventh day": 7,
}

EXPECTED_TRAYS = list(range(1, 13))

BAND_OPTIONS = {
    "ALL",
    "D",
    "MS",
}

SUPPORTED_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
}


# ============================================================
# 3) GENERAL HELPERS
# ============================================================

def natural_key(text: object):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", str(text))
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


def tray_number_from_name(value: object):
    match = re.search(r"(\d+)", str(value))

    if match:
        return int(match.group(1))

    return None


def tray_name_from_number(tray_no: int):
    return f"Tray {tray_no}"


def relative_path(path: Path | None, root: Path):
    if path is None:
        return ""

    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def safe_json_key(day: str, tray: str, capture_id: str):
    return f"{day}|{tray}|{capture_id}"


# ============================================================
# 4) FILE CLASSIFICATION
# ============================================================

def parse_image_metadata(path: Path):
    """
    Supports DJI names like:
        DJI_20260702125244_0001_D.JPG
        DJI_20260702125244_0001_MS_NIR.TIF

    Key point:
    The logical tray number is usually the four-digit sequence ID:
        0001 -> Tray 1
        0012 -> Tray 12
    """

    stem = path.stem.upper()

    match = re.match(
        r"^(?P<prefix>.+?)_"
        r"(?P<seq>\d{4})_"
        r"(?P<band>MS_NIR|MS_RE|MS_R|MS_G|D|F)$",
        stem,
    )

    if match:
        prefix = match.group("prefix")
        sequence_id = match.group("seq")
        band = match.group("band")

        capture_id = f"{prefix}_{sequence_id}"
        logical_key = sequence_id

        return {
            "capture_id": capture_id,
            "logical_key": logical_key,
            "sequence_id": sequence_id,
            "sequence_number": int(sequence_id),
            "band": band,
        }

    for band in [
        "MS_NIR",
        "MS_RE",
        "MS_R",
        "MS_G",
        "D",
        "F",
    ]:
        if re.search(rf"(?:^|_){re.escape(band)}$", stem):
            capture_id = re.sub(
                rf"_{re.escape(band)}$",
                "",
                stem,
            )

            return {
                "capture_id": capture_id,
                "logical_key": capture_id,
                "sequence_id": "",
                "sequence_number": None,
                "band": band,
            }

    return {
        "capture_id": stem,
        "logical_key": stem,
        "sequence_id": "",
        "sequence_number": None,
        "band": "UNKNOWN",
    }


def choose_canonical_capture_id(bands_dict: dict):
    preference = [
        "D",
        "MS_NIR",
        "MS_G",
        "MS_R",
        "MS_RE",
    ]

    for band in preference:
        if band in bands_dict:
            return bands_dict[band]["capture_id"]

    return "UNKNOWN_CAPTURE"


def group_image_files(files: list[Path]):
    grouped = {}

    for file in sorted(files, key=lambda path: natural_key(path.name)):
        meta = parse_image_metadata(file)
        band = meta["band"]

        if band in {
            "UNKNOWN",
            "F",
        }:
            continue

        logical_key = meta["logical_key"]

        grouped.setdefault(
            logical_key,
            {
                "logical_key": logical_key,
                "sequence_id": meta["sequence_id"],
                "sequence_number": meta["sequence_number"],
                "bands": {},
                "capture_ids_seen": set(),
            },
        )

        if band not in grouped[logical_key]["bands"]:
            grouped[logical_key]["bands"][band] = {
                "path": file,
                "capture_id": meta["capture_id"],
            }

        grouped[logical_key]["capture_ids_seen"].add(
            meta["capture_id"]
        )

    results = []

    for logical_key, info in grouped.items():
        bands = info["bands"]

        results.append(
            {
                "logical_key": logical_key,
                "sequence_id": info["sequence_id"],
                "sequence_number": info["sequence_number"],
                "canonical_capture_id": choose_canonical_capture_id(bands),
                "bands": bands,
                "capture_ids_seen": sorted(info["capture_ids_seen"]),
            }
        )

    results.sort(
        key=lambda item: (
            item["sequence_number"]
            if item["sequence_number"] is not None
            else 9999,
            natural_key(item["canonical_capture_id"]),
        )
    )

    return results


def image_files_directly_in(folder: Path):
    return [
        file
        for file in folder.iterdir()
        if file.is_file()
        and file.suffix.casefold() in SUPPORTED_EXTENSIONS
    ]


def image_files_recursively_in(folder: Path):
    return [
        file
        for file in folder.rglob("*")
        if file.is_file()
        and file.suffix.casefold() in SUPPORTED_EXTENSIONS
    ]


def find_jobs_for_day(day_folder: Path):
    """
    Supports two layouts.

    Layout A:
        Day X contains raw images directly.
        Sequence 0001 maps to Tray 1.

    Layout B:
        Day X contains Tray folders.
        Images inside Tray N folder map to Tray N.
    """

    jobs = []

    direct_files = image_files_directly_in(day_folder)
    direct_sets = group_image_files(direct_files)

    if direct_sets:
        for capture_set in direct_sets:
            sequence_number = capture_set["sequence_number"]

            if sequence_number is None:
                continue

            if sequence_number not in EXPECTED_TRAYS:
                continue

            jobs.append(
                {
                    "day_folder": day_folder,
                    "tray_name": tray_name_from_number(sequence_number),
                    "tray_no": sequence_number,
                    "logical_key": capture_set["logical_key"],
                    "capture_id": capture_set["canonical_capture_id"],
                    "bands": capture_set["bands"],
                    "capture_ids_seen": capture_set["capture_ids_seen"],
                    "layout_source": "day_folder_sequence_mapping",
                }
            )

    tray_folders = sorted(
        [
            folder
            for folder in day_folder.iterdir()
            if folder.is_dir()
            and folder.name.casefold().startswith("tray")
        ],
        key=lambda folder: natural_key(folder.name),
    )

    for tray_folder in tray_folders:
        tray_no = tray_number_from_name(tray_folder.name)

        if tray_no is None:
            continue

        files = image_files_recursively_in(tray_folder)
        capture_sets = group_image_files(files)

        for capture_set in capture_sets:
            jobs.append(
                {
                    "day_folder": day_folder,
                    "tray_name": tray_name_from_number(tray_no),
                    "tray_no": tray_no,
                    "logical_key": capture_set["logical_key"],
                    "capture_id": capture_set["canonical_capture_id"],
                    "bands": capture_set["bands"],
                    "capture_ids_seen": capture_set["capture_ids_seen"],
                    "layout_source": "tray_folder_mapping",
                }
            )

    # Remove duplicates if both layouts accidentally contain same Day/Tray.
    deduped = {}

    for job in jobs:
        key = (
            job["day_folder"].name,
            job["tray_no"],
            job["logical_key"],
        )

        if key not in deduped:
            deduped[key] = job
            continue

        # Prefer direct day-folder files if both exist because that matches
        # the current Third Trial screenshot structure.
        if job["layout_source"] == "day_folder_sequence_mapping":
            deduped[key] = job

    return sorted(
        deduped.values(),
        key=lambda item: (
            item["tray_no"],
            natural_key(item["capture_id"]),
        ),
    )


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
            raise ValueError(
                f"{path.name} is not a single-band TIFF. Shape: {array.shape}"
            )

    if array.ndim != 2:
        raise ValueError(
            f"{path.name} is not a 2D single-band TIFF. Shape: {array.shape}"
        )

    return array


def ms_preview(array: np.ndarray) -> np.ndarray:
    values = array.astype(np.float32)

    low, high = np.percentile(
        values,
        [
            1,
            99,
        ],
    )

    if high <= low:
        high = low + 1.0

    preview = np.clip(
        (values - low) * 255.0 / (high - low),
        0,
        255,
    ).astype(np.uint8)

    return preview


# ============================================================
# 6) MANUAL FOUR-CORNER SELECTION
# ============================================================

def get_four_corners(
    image: np.ndarray,
    title: str,
    grayscale: bool = False,
):
    """
    Click four tray corners in this order:

    1. Top-left
    2. Top-right
    3. Bottom-right
    4. Bottom-left

    Right-click removes the latest point.
    """

    fig, ax = plt.subplots(figsize=(15, 10))

    if grayscale:
        ax.imshow(image, cmap="gray")
    else:
        ax.imshow(image)

    ax.set_title(
        title
        + "\nClick: Top-left → Top-right → Bottom-right → Bottom-left"
        + "\nRight-click removes the latest point. Close the window to cancel.",
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

    width = max(
        20,
        int(round(max(top_width, bottom_width))),
    )

    height = max(
        20,
        int(round(max(left_height, right_height))),
    )

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
        return {
            "records": {},
        }

    try:
        with path.open("r", encoding="utf-8") as file:
            return json.load(file)

    except Exception:
        return {
            "records": {},
        }


def save_json(path: Path, data: dict):
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open("w", encoding="utf-8") as file:
        json.dump(
            data,
            file,
            indent=2,
        )


def expected_outputs(
    output_folder: Path,
    capture_id: str,
    process_bands: str,
):
    if process_bands == "D":
        return [
            output_folder / f"{capture_id}_D.JPG",
        ]

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


def outputs_exist(
    output_folder: Path,
    capture_id: str,
    process_bands: str,
):
    return all(
        path.exists()
        for path in expected_outputs(
            output_folder,
            capture_id,
            process_bands,
        )
    )


def write_manifest(
    path: Path,
    rows: list[dict],
):
    fields = [
        "day_order",
        "day",
        "tray",
        "tray_no",
        "logical_key",
        "capture_id",
        "capture_ids_seen",
        "layout_source",
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

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with path.open(
        "w",
        newline="",
        encoding="utf-8",
    ) as file:
        writer = csv.DictWriter(
            file,
            fieldnames=fields,
        )

        writer.writeheader()
        writer.writerows(rows)


# ============================================================
# 9) MAIN
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Script 01: Crop Third Trial D/RGB and multispectral images "
            "using separate four-corner references."
        )
    )

    parser.add_argument(
        "--days",
        help='Example: --days "Day 1,Day 4"',
    )

    parser.add_argument(
        "--trays",
        help='Example: --trays "Tray 1,Tray 12"',
    )

    parser.add_argument(
        "--process-bands",
        default="ALL",
        choices=sorted(BAND_OPTIONS),
        help=(
            "ALL = crop D + multispectral; "
            "D = crop D only; "
            "MS = crop multispectral only."
        ),
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

    if not day_folders:
        print(f"ERROR: No Day folders found inside:\n{INPUT_ROOT}")
        return 1

    jobs = []

    for day_folder in day_folders:
        if selected_days and day_folder.name.casefold() not in selected_days:
            continue

        day_jobs = find_jobs_for_day(day_folder)

        for job in day_jobs:
            if selected_trays and job["tray_name"].casefold() not in selected_trays:
                continue

            bands = job["bands"]

            if process_bands == "ALL":
                missing = [
                    band
                    for band in ALL_REQUIRED_BANDS
                    if band not in bands
                ]

            elif process_bands == "D":
                missing = [
                    band
                    for band in ["D"]
                    if band not in bands
                ]

            else:
                missing = [
                    band
                    for band in MS_BANDS
                    if band not in bands
                ]

            job["missing"] = missing
            jobs.append(job)

    if not jobs:
        print("No valid Third Trial image sets were found.")
        return 1

    print("\nSCRIPT 01 — THIRD TRIAL DUAL-REFERENCE CROPPING")
    print("=" * 70)
    print(f"Input folder:\n{INPUT_ROOT}")
    print(f"\nOutput folder:\n{OUTPUT_ROOT}")
    print(f"\nProcessing mode: {process_bands}")
    print("\nF preview images are ignored.")
    print("\nDetected image sets:\n")

    for job in jobs:
        if job["missing"]:
            state = "INCOMPLETE: " + ", ".join(job["missing"])
        else:
            state = "READY"

        id_note = ""

        if len(job["capture_ids_seen"]) > 1:
            id_note = " | merged timestamp set"

        print(
            f"{job['day_folder'].name} > {job['tray_name']} > "
            f"{job['capture_id']} : {state} | {job['layout_source']}{id_note}"
        )

    if args.dry_run:
        print("\nDry run complete. No crops created.")
        return 0

    config_path = (
        OUTPUT_ROOT
        / "_config"
        / "crop_points.json"
    )

    manifest_path = (
        OUTPUT_ROOT
        / "_reports"
        / "crop_manifest.csv"
    )

    config = load_json(config_path)

    if "records" not in config:
        config["records"] = {}

    manifest_rows = []

    for job in jobs:
        day_folder = job["day_folder"]
        day_name = day_folder.name
        tray_name = job["tray_name"]
        tray_no = job["tray_no"]
        capture_id = job["capture_id"]
        logical_key = job["logical_key"]
        bands = job["bands"]
        missing = job["missing"]

        output_folder = (
            OUTPUT_ROOT
            / day_name
            / tray_name
        )

        output_folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        row = {
            "day_order": DAY_NAME_TO_ORDER.get(
                day_name.casefold(),
                "",
            ),
            "day": day_name,
            "tray": tray_name,
            "tray_no": tray_no,
            "logical_key": logical_key,
            "capture_id": capture_id,
            "capture_ids_seen": "; ".join(job["capture_ids_seen"]),
            "layout_source": job["layout_source"],
            "status": "",
            "process_bands": process_bands,
            "source_d": relative_path(
                bands.get("D", {}).get("path"),
                INPUT_ROOT,
            ),
            "source_ms_g": relative_path(
                bands.get("MS_G", {}).get("path"),
                INPUT_ROOT,
            ),
            "source_ms_r": relative_path(
                bands.get("MS_R", {}).get("path"),
                INPUT_ROOT,
            ),
            "source_ms_re": relative_path(
                bands.get("MS_RE", {}).get("path"),
                INPUT_ROOT,
            ),
            "source_ms_nir": relative_path(
                bands.get("MS_NIR", {}).get("path"),
                INPUT_ROOT,
            ),
            "output_d": "",
            "output_ms_g": "",
            "output_ms_r": "",
            "output_ms_re": "",
            "output_ms_nir": "",
            "d_crop_width": "",
            "d_crop_height": "",
            "ms_crop_width": "",
            "ms_crop_height": "",
            "notes": "",
        }

        if missing:
            row["status"] = "SKIPPED_INCOMPLETE"
            row["notes"] = (
                "Missing required band(s): "
                + ", ".join(missing)
            )

            manifest_rows.append(row)
            print(
                f"\nSKIPPED incomplete set: {day_name} > {tray_name} > {capture_id}"
            )
            continue

        if (
            outputs_exist(output_folder, capture_id, process_bands)
            and not args.redo
        ):
            row["status"] = "SKIPPED_EXISTS"
            row["notes"] = (
                "Expected outputs already exist. Use --redo to overwrite."
            )

            if (output_folder / f"{capture_id}_D.JPG").exists():
                row["output_d"] = relative_path(
                    output_folder / f"{capture_id}_D.JPG",
                    OUTPUT_ROOT,
                )

            for band in MS_BANDS:
                path = output_folder / f"{capture_id}_{band}.TIF"

                if path.exists():
                    row[f"output_{band.casefold()}"] = relative_path(
                        path,
                        OUTPUT_ROOT,
                    )

            manifest_rows.append(row)

            print(
                f"\nSKIPPED existing: {day_name} > {tray_name} > {capture_id}"
            )
            continue

        config_key = safe_json_key(
            day_name,
            tray_name,
            capture_id,
        )

        crop_record = config["records"].setdefault(
            config_key,
            {
                "day": day_name,
                "tray": tray_name,
                "tray_no": tray_no,
                "capture_id": capture_id,
                "d_points": None,
                "ms_points": None,
            },
        )

        try:
            d_matrix = None
            d_width = None
            d_height = None

            if process_bands in {
                "ALL",
                "D",
            }:
                d_path = bands["D"]["path"]

                d_image = read_rgb(d_path)

                if args.redo or crop_record.get("d_points") is None:
                    print(
                        f"\nD/RGB crop selection: {day_name} > {tray_name} > {capture_id}"
                    )

                    d_points = get_four_corners(
                        d_image,
                        title=(
                            f"D/RGB crop | {day_name} | {tray_name} | {capture_id}"
                        ),
                        grayscale=False,
                    )

                    if d_points is None:
                        row["status"] = "CANCELLED"
                        row["notes"] = (
                            "D/RGB crop was cancelled before four points were selected."
                        )

                        manifest_rows.append(row)
                        save_json(config_path, config)
                        continue

                    crop_record["d_points"] = d_points.tolist()
                    save_json(config_path, config)

                d_points = np.asarray(
                    crop_record["d_points"],
                    dtype=np.float32,
                )

                d_matrix, d_width, d_height = make_transform(
                    d_points
                )

                d_crop = crop_with_transform(
                    d_image,
                    d_matrix,
                    d_width,
                    d_height,
                    interpolation=cv2.INTER_CUBIC,
                )

                d_output = output_folder / f"{capture_id}_D.JPG"

                cv2.imwrite(
                    str(d_output),
                    cv2.cvtColor(
                        d_crop,
                        cv2.COLOR_RGB2BGR,
                    ),
                    [
                        cv2.IMWRITE_JPEG_QUALITY,
                        95,
                    ],
                )

                row["output_d"] = relative_path(
                    d_output,
                    OUTPUT_ROOT,
                )

                row["d_crop_width"] = d_width
                row["d_crop_height"] = d_height

            ms_matrix = None
            ms_width = None
            ms_height = None

            if process_bands in {
                "ALL",
                "MS",
            }:
                ms_nir_path = bands["MS_NIR"]["path"]

                ms_nir_array = read_single_band_tif(
                    ms_nir_path
                )

                ms_nir_display = ms_preview(
                    ms_nir_array
                )

                if args.redo or crop_record.get("ms_points") is None:
                    print(
                        f"\nMS crop selection: {day_name} > {tray_name} > {capture_id}"
                    )

                    ms_points = get_four_corners(
                        ms_nir_display,
                        title=(
                            f"MS_NIR crop | {day_name} | {tray_name} | {capture_id}"
                        ),
                        grayscale=True,
                    )

                    if ms_points is None:
                        row["status"] = "CANCELLED"
                        row["notes"] = (
                            "MS crop was cancelled before four points were selected."
                        )

                        manifest_rows.append(row)
                        save_json(config_path, config)
                        continue

                    crop_record["ms_points"] = ms_points.tolist()
                    save_json(config_path, config)

                ms_points = np.asarray(
                    crop_record["ms_points"],
                    dtype=np.float32,
                )

                ms_matrix, ms_width, ms_height = make_transform(
                    ms_points
                )

                row["ms_crop_width"] = ms_width
                row["ms_crop_height"] = ms_height

                for band in MS_BANDS:
                    source_path = bands[band]["path"]

                    source_array = read_single_band_tif(
                        source_path
                    )

                    ms_crop = crop_with_transform(
                        source_array,
                        ms_matrix,
                        ms_width,
                        ms_height,
                        interpolation=cv2.INTER_NEAREST,
                    )

                    output_path = output_folder / f"{capture_id}_{band}.TIF"

                    tifffile.imwrite(
                        str(output_path),
                        ms_crop,
                    )

                    row[f"output_{band.casefold()}"] = relative_path(
                        output_path,
                        OUTPUT_ROOT,
                    )

            row["status"] = "CROPPED"
            row["notes"] = (
                "D/RGB and/or MS crop completed successfully."
            )

            print(
                f"\nCROPPED: {day_name} > {tray_name} > {capture_id}"
            )

        except Exception as error:
            row["status"] = "FAILED"
            row["notes"] = str(error)

            print(
                f"\nFAILED: {day_name} > {tray_name} > {capture_id}"
            )
            print(f"Reason: {error}")

        manifest_rows.append(row)

    write_manifest(
        manifest_path,
        manifest_rows,
    )

    save_json(
        config_path,
        config,
    )

    status_counts = {}

    for row in manifest_rows:
        status_counts[row["status"]] = (
            status_counts.get(row["status"], 0)
            + 1
        )

    print("\n" + "=" * 70)
    print("SCRIPT 01 FINISHED")
    print("=" * 70)

    for status, count in sorted(status_counts.items()):
        print(f"{status}: {count}")

    print(f"\nCrop point config:\n{config_path}")
    print(f"\nManifest:\n{manifest_path}")
    print(f"\nOutput folder:\n{OUTPUT_ROOT}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())