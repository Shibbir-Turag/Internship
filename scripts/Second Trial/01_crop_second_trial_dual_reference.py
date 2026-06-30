from __future__ import annotations

import argparse
import csv
import json
import os
import re
from pathlib import Path

import cv2
import numpy as np
import tifffile


# ============================================================
# PROJECT SETTINGS
# ============================================================

INTERNSHIP_ROOT = None

DAY_ORDER = {
    "first day": 1,
    "second day": 2,
    "third day": 3,
    "fourth day": 4,
    "fifth day": 5,
    "ninth day": 9,
}

MS_BANDS = ["MS_G", "MS_R", "MS_RE", "MS_NIR"]
REQUIRED_BANDS = ["D", *MS_BANDS]

OUTPUT_FOLDER_NAME = "01_Crop_Dual_Reference"

DISPLAY_MAX_WIDTH = 1400
DISPLAY_MAX_HEIGHT = 900


# ============================================================
# PATH FUNCTIONS
# ============================================================

def get_internship_root() -> Path:
    if INTERNSHIP_ROOT is not None:
        return Path(INTERNSHIP_ROOT)

    candidates = []

    if os.environ.get("OneDrive"):
        candidates.append(
            Path(os.environ["OneDrive"]) / "Desktop" / "Internship"
        )

    candidates.extend(
        [
            Path.home() / "OneDrive" / "Desktop" / "Internship",
            Path.home() / "Desktop" / "Internship",
        ]
    )

    for candidate in candidates:
        if (candidate / "data" / "Second Trial").exists():
            return candidate

    return candidates[0]


def natural_key(text: str):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", text)
    ]


def day_sort_key(folder: Path):
    return (
        DAY_ORDER.get(folder.name.casefold(), 999),
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


def safe_filename(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", text)


# ============================================================
# IMAGE FILE CLASSIFICATION
# ============================================================

def classify_image(path: Path):
    """
    Examples:
        DJI_20260623124403_0008_D.JPG      -> capture ID, D
        DJI_20260623124403_0008_MS_NIR.TIF -> capture ID, MS_NIR

    F preview images are intentionally ignored.
    """

    stem = path.stem.upper()

    ordered_tokens = [
        ("MS_NIR", "MS_NIR"),
        ("MS_RE", "MS_RE"),
        ("MS_R", "MS_R"),
        ("MS_G", "MS_G"),
        ("D", "D"),
        ("F", "F_IGNORED"),
    ]

    for token, label in ordered_tokens:
        if re.search(rf"(?:^|_){re.escape(token)}$", stem):
            capture_id = re.sub(
                rf"_{re.escape(token)}$",
                "",
                stem,
            )

            return capture_id or stem, label

    return stem, "UNKNOWN"


def find_capture_sets(tray_folder: Path):
    """
    Build:
    {
        capture_id: {
            "D": Path(...),
            "MS_G": Path(...),
            ...
        }
    }
    """

    capture_sets = {}

    image_files = sorted(
        [
            file
            for file in tray_folder.rglob("*")
            if file.is_file()
            and file.suffix.casefold() in {
                ".jpg",
                ".jpeg",
                ".tif",
                ".tiff",
            }
        ],
        key=lambda file: natural_key(file.name),
    )

    for image_path in image_files:
        capture_id, band = classify_image(image_path)

        if band in {"UNKNOWN", "F_IGNORED"}:
            continue

        capture_sets.setdefault(capture_id, {})
        capture_sets[capture_id].setdefault(band, [])

        capture_sets[capture_id][band].append(image_path)

    cleaned_sets = {}

    for capture_id, band_lists in capture_sets.items():
        cleaned_sets[capture_id] = {}

        for band, paths in band_lists.items():
            cleaned_sets[capture_id][band] = paths[0]

    return cleaned_sets


# ============================================================
# IMAGE READING / DISPLAY
# ============================================================

def read_d_image(path: Path) -> np.ndarray:
    image = cv2.imread(str(path), cv2.IMREAD_COLOR)

    if image is None:
        raise ValueError(f"Could not read D/RGB image: {path.name}")

    return image


def read_ms_band(path: Path) -> np.ndarray:
    array = tifffile.imread(path)

    if array.ndim == 3:
        if array.shape[0] == 1:
            array = array[0]
        elif array.shape[-1] == 1:
            array = array[:, :, 0]
        else:
            raise ValueError(
                f"{path.name} is not a single-band multispectral TIFF."
            )

    if array.ndim != 2:
        raise ValueError(
            f"{path.name} is not a 2D multispectral band."
        )

    return array


def ms_preview(ms_array: np.ndarray) -> np.ndarray:
    """
    Display-only preview of MS_NIR.
    Source pixels are never changed.
    """

    values = ms_array.astype(np.float32)

    low, high = np.percentile(values, [1, 99])

    if high <= low:
        high = low + 1

    scaled = np.clip(
        (values - low) * 255 / (high - low),
        0,
        255,
    ).astype(np.uint8)

    return cv2.cvtColor(
        scaled,
        cv2.COLOR_GRAY2BGR,
    )


def resize_for_display(image: np.ndarray):
    height, width = image.shape[:2]

    scale = min(
        DISPLAY_MAX_WIDTH / width,
        DISPLAY_MAX_HEIGHT / height,
        1.0,
    )

    if scale == 1.0:
        return image.copy(), scale

    resized = cv2.resize(
        image,
        (
            int(round(width * scale)),
            int(round(height * scale)),
        ),
        interpolation=cv2.INTER_AREA,
    )

    return resized, scale


# ============================================================
# FOUR-POINT CLICK INTERFACE
# ============================================================

def collect_four_corners(
    source_preview: np.ndarray,
    title: str,
):
    """
    Click order:
        1. top-left
        2. top-right
        3. bottom-right
        4. bottom-left

    Controls:
        R = clear selections
        ESC = cancel current crop
    """

    display, scale = resize_for_display(source_preview)
    points = []

    window_name = title

    def redraw():
        canvas = display.copy()

        instructions = [
            "Click 4 corners in this order:",
            "1 Top-left | 2 Top-right | 3 Bottom-right | 4 Bottom-left",
            "R = reset | ESC = cancel",
        ]

        for index, text in enumerate(instructions):
            cv2.putText(
                canvas,
                text,
                (15, 30 + index * 28),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.70,
                (0, 0, 0),
                4,
                cv2.LINE_AA,
            )

            cv2.putText(
                canvas,
                text,
                (15, 30 + index * 28),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.70,
                (255, 255, 255),
                2,
                cv2.LINE_AA,
            )

        for index, point in enumerate(points):
            x, y = point

            cv2.circle(
                canvas,
                (x, y),
                7,
                (0, 0, 255),
                -1,
            )

            cv2.putText(
                canvas,
                str(index + 1),
                (x + 10, y - 10),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.85,
                (0, 0, 255),
                2,
                cv2.LINE_AA,
            )

        if len(points) > 1:
            cv2.polylines(
                canvas,
                [np.asarray(points, dtype=np.int32)],
                False,
                (0, 255, 255),
                2,
            )

        cv2.imshow(window_name, canvas)

    def mouse_callback(event, x, y, flags, parameter):
        if event == cv2.EVENT_LBUTTONDOWN and len(points) < 4:
            points.append((x, y))
            redraw()

    cv2.namedWindow(window_name, cv2.WINDOW_NORMAL)
    cv2.resizeWindow(
        window_name,
        display.shape[1],
        display.shape[0],
    )

    cv2.setMouseCallback(
        window_name,
        mouse_callback,
    )

    redraw()

    while len(points) < 4:
        key = cv2.waitKey(20) & 0xFF

        if key in {27, ord("q")}:
            cv2.destroyWindow(window_name)
            return None

        if key in {ord("r"), ord("R")}:
            points.clear()
            redraw()

    cv2.destroyWindow(window_name)

    source_points = np.asarray(
        [
            [
                point[0] / scale,
                point[1] / scale,
            ]
            for point in points
        ],
        dtype=np.float32,
    )

    return source_points


# ============================================================
# PERSPECTIVE CROP FUNCTIONS
# ============================================================

def crop_output_size(points: np.ndarray):
    """
    Calculate output dimensions from a four-corner quadrilateral.
    """

    top_width = np.linalg.norm(points[1] - points[0])
    bottom_width = np.linalg.norm(points[2] - points[3])

    left_height = np.linalg.norm(points[3] - points[0])
    right_height = np.linalg.norm(points[2] - points[1])

    output_width = max(20, int(round(max(top_width, bottom_width))))
    output_height = max(20, int(round(max(left_height, right_height))))

    return output_width, output_height


def perspective_crop(
    image: np.ndarray,
    points: np.ndarray,
    interpolation: int,
):
    width, height = crop_output_size(points)

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

    cropped = cv2.warpPerspective(
        image,
        matrix,
        (width, height),
        flags=interpolation,
        borderMode=cv2.BORDER_CONSTANT,
        borderValue=0,
    )

    return cropped, matrix, width, height


# ============================================================
# CONFIG / MANIFEST
# ============================================================

def load_config(config_path: Path):
    if not config_path.exists():
        return {"records": {}}

    try:
        with config_path.open("r", encoding="utf-8") as file:
            return json.load(file)
    except Exception:
        return {"records": {}}


def save_config(config_path: Path, config: dict):
    config_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with config_path.open("w", encoding="utf-8") as file:
        json.dump(
            config,
            file,
            indent=2,
        )


def output_files_complete(
    output_folder: Path,
    capture_id: str,
):
    expected = [
        output_folder / f"{capture_id}_D.JPG",
        output_folder / f"{capture_id}_MS_G.TIF",
        output_folder / f"{capture_id}_MS_R.TIF",
        output_folder / f"{capture_id}_MS_RE.TIF",
        output_folder / f"{capture_id}_MS_NIR.TIF",
    ]

    return all(path.exists() for path in expected)


def write_manifest(
    manifest_path: Path,
    rows: list[dict],
):
    fields = [
        "day_order",
        "day",
        "tray",
        "capture_id",
        "status",
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

    manifest_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with manifest_path.open(
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
# MAIN WORKFLOW
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "01 — Crop Second Trial D/RGB and multispectral images "
            "with independent four-corner references."
        )
    )

    parser.add_argument(
        "--days",
        help='Optional example: --days "First Day,Ninth Day"',
    )

    parser.add_argument(
        "--trays",
        help='Optional example: --trays "Tray 1,Tray 2"',
    )

    parser.add_argument(
        "--redo",
        action="store_true",
        help=(
            "Redo crops even when the five expected output files "
            "already exist."
        ),
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List all source capture sets without cropping.",
    )

    args = parser.parse_args()

    internship_root = get_internship_root()

    data_root = (
        internship_root
        / "data"
        / "Second Trial"
    )

    output_root = (
        internship_root
        / "outputs"
        / "Second Trial"
        / OUTPUT_FOLDER_NAME
    )

    config_path = (
        output_root
        / "_config"
        / "crop_config.json"
    )

    manifest_path = (
        output_root
        / "_reports"
        / "crop_manifest.csv"
    )

    if not data_root.exists():
        print(
            f"ERROR: Source folder not found:\n{data_root}"
        )
        return 1

    selected_days = parse_filter_list(args.days)
    selected_trays = parse_filter_list(args.trays)

    day_folders = sorted(
        [
            folder
            for folder in data_root.iterdir()
            if folder.is_dir()
            and folder.name.casefold() in DAY_ORDER
        ],
        key=day_sort_key,
    )

    jobs = []

    for day_folder in day_folders:
        if (
            selected_days
            and day_folder.name.casefold() not in selected_days
        ):
            continue

        tray_folders = sorted(
            [
                folder
                for folder in day_folder.iterdir()
                if folder.is_dir()
            ],
            key=lambda folder: natural_key(folder.name),
        )

        for tray_folder in tray_folders:
            if (
                selected_trays
                and tray_folder.name.casefold() not in selected_trays
            ):
                continue

            capture_sets = find_capture_sets(tray_folder)

            for capture_id, files in capture_sets.items():
                missing = [
                    band
                    for band in REQUIRED_BANDS
                    if band not in files
                ]

                jobs.append(
                    {
                        "day_folder": day_folder,
                        "tray_folder": tray_folder,
                        "capture_id": capture_id,
                        "files": files,
                        "missing": missing,
                    }
                )

    if not jobs:
        print("No capture sets were found.")
        return 1

    print("\n01 — DUAL-REFERENCE CROP WORKFLOW")
    print("=" * 70)
    print(f"Source: {data_root}")
    print(f"Output: {output_root}")
    print("\nF preview files are ignored.")

    for job in jobs:
        status = (
            "READY"
            if not job["missing"]
            else "INCOMPLETE: missing " + ", ".join(job["missing"])
        )

        print(
            f"{job['day_folder'].name} > "
            f"{job['tray_folder'].name} > "
            f"{job['capture_id']} : {status}"
        )

    if args.dry_run:
        print("\nDry run complete. No crops were created.")
        return 0

    config = load_config(config_path)
    manifest_rows = []

    for job in jobs:
        day_folder = job["day_folder"]
        tray_folder = job["tray_folder"]
        capture_id = job["capture_id"]
        files = job["files"]
        missing = job["missing"]

        day_name = day_folder.name
        tray_name = tray_folder.name
        day_order = DAY_ORDER.get(day_name.casefold(), 999)

        output_folder = (
            output_root
            / day_name
            / tray_name
        )

        output_folder.mkdir(
            parents=True,
            exist_ok=True,
        )

        base_row = {
            "day_order": day_order,
            "day": day_name,
            "tray": tray_name,
            "capture_id": capture_id,
            "source_d": str(files.get("D", "")),
            "source_ms_g": str(files.get("MS_G", "")),
            "source_ms_r": str(files.get("MS_R", "")),
            "source_ms_re": str(files.get("MS_RE", "")),
            "source_ms_nir": str(files.get("MS_NIR", "")),
            "output_d": str(
                output_folder / f"{capture_id}_D.JPG"
            ),
            "output_ms_g": str(
                output_folder / f"{capture_id}_MS_G.TIF"
            ),
            "output_ms_r": str(
                output_folder / f"{capture_id}_MS_R.TIF"
            ),
            "output_ms_re": str(
                output_folder / f"{capture_id}_MS_RE.TIF"
            ),
            "output_ms_nir": str(
                output_folder / f"{capture_id}_MS_NIR.TIF"
            ),
            "d_crop_width": "",
            "d_crop_height": "",
            "ms_crop_width": "",
            "ms_crop_height": "",
            "notes": "",
        }

        if missing:
            manifest_rows.append(
                {
                    **base_row,
                    "status": "SKIPPED_INCOMPLETE",
                    "notes": "Missing required files: "
                    + ", ".join(missing),
                }
            )

            print(
                f"SKIPPED: {day_name} > {tray_name} > "
                f"{capture_id} | Missing {', '.join(missing)}"
            )

            continue

        if (
            output_files_complete(output_folder, capture_id)
            and not args.redo
        ):
            manifest_rows.append(
                {
                    **base_row,
                    "status": "SKIPPED_EXISTING",
                    "notes": (
                        "Expected D and multispectral output files "
                        "already exist."
                    ),
                }
            )

            print(
                f"SKIPPED: {day_name} > {tray_name} > "
                f"{capture_id} | Existing crop found"
            )

            continue

        try:
            # ------------------------------------------------
            # D / RGB CROP
            # ------------------------------------------------
            d_image = read_d_image(files["D"])

            d_title = (
                f"D/RGB crop | {day_name} | "
                f"{tray_name} | {capture_id}"
            )

            d_points = collect_four_corners(
                d_image,
                d_title,
            )

            if d_points is None:
                manifest_rows.append(
                    {
                        **base_row,
                        "status": "CANCELLED",
                        "notes": "D/RGB crop cancelled by user.",
                    }
                )

                print(
                    f"CANCELLED: {day_name} > {tray_name} > "
                    f"{capture_id} | D/RGB"
                )

                continue

            d_crop, _, d_width, d_height = perspective_crop(
                d_image,
                d_points,
                cv2.INTER_LINEAR,
            )

            d_output = output_folder / f"{capture_id}_D.JPG"

            cv2.imwrite(
                str(d_output),
                d_crop,
                [
                    cv2.IMWRITE_JPEG_QUALITY,
                    100,
                ],
            )

            # ------------------------------------------------
            # MULTISPECTRAL CROP
            # ------------------------------------------------
            ms_nir = read_ms_band(files["MS_NIR"])

            ms_title = (
                f"MS_NIR crop | {day_name} | "
                f"{tray_name} | {capture_id}"
            )

            ms_points = collect_four_corners(
                ms_preview(ms_nir),
                ms_title,
            )

            if ms_points is None:
                manifest_rows.append(
                    {
                        **base_row,
                        "status": "PARTIAL_CANCELLED",
                        "d_crop_width": d_width,
                        "d_crop_height": d_height,
                        "notes": (
                            "D crop saved, but multispectral crop "
                            "was cancelled by user."
                        ),
                    }
                )

                print(
                    f"PARTIAL: {day_name} > {tray_name} > "
                    f"{capture_id} | MS crop cancelled"
                )

                continue

            ms_output_width, ms_output_height = crop_output_size(
                ms_points
            )

            for band in MS_BANDS:
                source_band = read_ms_band(files[band])

                destination = np.asarray(
                    [
                        [0, 0],
                        [ms_output_width - 1, 0],
                        [ms_output_width - 1, ms_output_height - 1],
                        [0, ms_output_height - 1],
                    ],
                    dtype=np.float32,
                )

                matrix = cv2.getPerspectiveTransform(
                    ms_points.astype(np.float32),
                    destination,
                )

                cropped_band = cv2.warpPerspective(
                    source_band,
                    matrix,
                    (
                        ms_output_width,
                        ms_output_height,
                    ),
                    flags=cv2.INTER_NEAREST,
                    borderMode=cv2.BORDER_CONSTANT,
                    borderValue=0,
                )

                output_path = (
                    output_folder
                    / f"{capture_id}_{band}.TIF"
                )

                tifffile.imwrite(
                    output_path,
                    cropped_band,
                )

            config_key = (
                f"{day_name}|{tray_name}|{capture_id}"
            )

            config["records"][config_key] = {
                "day": day_name,
                "tray": tray_name,
                "capture_id": capture_id,
                "d_points": d_points.tolist(),
                "ms_points": ms_points.tolist(),
                "d_crop_size": [d_width, d_height],
                "ms_crop_size": [
                    ms_output_width,
                    ms_output_height,
                ],
            }

            save_config(config_path, config)

            manifest_rows.append(
                {
                    **base_row,
                    "status": "PASS",
                    "d_crop_width": d_width,
                    "d_crop_height": d_height,
                    "ms_crop_width": ms_output_width,
                    "ms_crop_height": ms_output_height,
                    "notes": (
                        "D/RGB and multispectral crops saved using "
                        "independent four-corner references. "
                        "F preview ignored."
                    ),
                }
            )

            print(
                f"PASS: {day_name} > {tray_name} > "
                f"{capture_id}"
            )

        except Exception as error:
            manifest_rows.append(
                {
                    **base_row,
                    "status": "FAIL",
                    "notes": str(error),
                }
            )

            print(
                f"FAIL: {day_name} > {tray_name} > "
                f"{capture_id} | {error}"
            )

    manifest_rows.sort(
        key=lambda row: (
            row["day_order"],
            natural_key(row["tray"]),
            natural_key(row["capture_id"]),
        )
    )

    write_manifest(
        manifest_path,
        manifest_rows,
    )

    passed = sum(
        row["status"] == "PASS"
        for row in manifest_rows
    )

    failed = sum(
        row["status"] == "FAIL"
        for row in manifest_rows
    )

    print("\n" + "=" * 70)
    print("SCRIPT 01 FINISHED")
    print("=" * 70)
    print(f"PASS: {passed}")
    print(f"FAIL: {failed}")
    print(f"\nOutput folder:\n{output_root}")
    print(f"\nManifest:\n{manifest_path}")
    print(f"\nCrop settings:\n{config_path}")

    return 0 if failed == 0 else 1


if __name__ == "__main__":
    raise SystemExit(main())