from __future__ import annotations

"""
SCRIPT 00 — FIX THIRD TRIAL DAY 4 IMAGE ORIENTATION

Purpose
-------
Rotate all Third Trial Day 4 raw images 90 degrees counter-clockwise in place.

Reason
------
Day 4 images were captured sideways, with the tray arrow on the right side.
The correct orientation should match the other days, where the tray arrow faces
forward/upward.

Important
---------
This script replaces the Day 4 files in the original folder.

To protect the raw data:
1. A backup is stored in:
   outputs/Third trial/00_Day4_Orientation_Fix/backup_original_day4_images

2. If a backup already exists from a previous incomplete run, this script uses
   that backup as the clean source. This prevents double-rotating files that may
   have been partially processed before.

Default mode is dry-run.
Use --apply to actually rotate the files.
"""

import argparse
import csv
import shutil
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np


# ============================================================
# 1) PATHS
# ============================================================

PROJECT_ROOT = Path(r"C:\Users\tshib\OneDrive\Desktop\Internship")

DAY4_ROOT = (
    PROJECT_ROOT
    / "data"
    / "Third Trial"
    / "Day 4"
)

OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "00_Day4_Orientation_Fix"
)

BACKUP_ROOT = OUTPUT_ROOT / "backup_original_day4_images"
REPORTS_ROOT = OUTPUT_ROOT / "_reports"


# ============================================================
# 2) SETTINGS
# ============================================================

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".tif",
    ".tiff",
    ".bmp",
}

ROTATION_DIRECTION = "CCW_90"


# ============================================================
# 3) FILE HELPERS
# ============================================================

def is_image_file(path: Path) -> bool:
    return (
        path.is_file()
        and path.suffix.casefold() in IMAGE_EXTENSIONS
    )


def collect_day4_images() -> list[Path]:
    if not DAY4_ROOT.exists():
        raise FileNotFoundError(
            f"Day 4 folder not found:\n{DAY4_ROOT}"
        )

    images = [
        path
        for path in DAY4_ROOT.rglob("*")
        if is_image_file(path)
    ]

    # Do not process anything inside possible backup/report folders
    images = [
        path
        for path in images
        if "00_Day4_Orientation_Fix" not in str(path)
    ]

    images.sort(key=lambda path: str(path).casefold())

    return images


def relative_to_day4(path: Path) -> Path:
    return path.relative_to(DAY4_ROOT)


def backup_path_for(original_path: Path) -> Path:
    return BACKUP_ROOT / relative_to_day4(original_path)


def ensure_backup(original_path: Path) -> Path:
    backup_path = backup_path_for(original_path)

    backup_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if not backup_path.exists():
        shutil.copy2(original_path, backup_path)

    return backup_path


# ============================================================
# 4) IMAGE READING / WRITING
# ============================================================

def read_image_any_depth(path: Path) -> np.ndarray:
    """
    Reads JPG/PNG/TIF images using OpenCV while preserving bit depth.

    This is more reliable for DJI MS .TIF files than the previous PIL version.
    """

    raw_bytes = np.fromfile(
        str(path),
        dtype=np.uint8,
    )

    image = cv2.imdecode(
        raw_bytes,
        cv2.IMREAD_UNCHANGED,
    )

    if image is None:
        raise ValueError(
            f"OpenCV could not read this image: {path}"
        )

    return image


def write_image_any_depth(
    path: Path,
    image: np.ndarray,
) -> None:
    suffix = path.suffix.casefold()

    if suffix in {".jpg", ".jpeg"}:
        encode_params = [
            int(cv2.IMWRITE_JPEG_QUALITY),
            95,
        ]

    elif suffix == ".png":
        encode_params = [
            int(cv2.IMWRITE_PNG_COMPRESSION),
            3,
        ]

    elif suffix in {".tif", ".tiff"}:
        encode_params = []

    else:
        encode_params = []

    success, encoded = cv2.imencode(
        path.suffix,
        image,
        encode_params,
    )

    if not success:
        raise ValueError(
            f"OpenCV could not encode output image: {path}"
        )

    encoded.tofile(str(path))


def rotate_counter_clockwise(image: np.ndarray) -> np.ndarray:
    return cv2.rotate(
        image,
        cv2.ROTATE_90_COUNTERCLOCKWISE,
    )


def rotate_file_from_source(
    source_path: Path,
    target_path: Path,
) -> tuple[tuple[int, int], tuple[int, int]]:
    image = read_image_any_depth(source_path)

    original_height, original_width = image.shape[:2]

    rotated = rotate_counter_clockwise(image)

    new_height, new_width = rotated.shape[:2]

    temp_path = target_path.with_name(
        target_path.stem + "_rotation_temp" + target_path.suffix
    )

    write_image_any_depth(
        temp_path,
        rotated,
    )

    temp_path.replace(target_path)

    return (
        (original_width, original_height),
        (new_width, new_height),
    )


# ============================================================
# 5) REPORT HELPERS
# ============================================================

def write_manifest(rows: list[dict]) -> Path:
    REPORTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    manifest_path = (
        REPORTS_ROOT
        / "day4_orientation_fix_manifest.csv"
    )

    fieldnames = [
        "file_no",
        "file_name",
        "relative_path",
        "original_path",
        "backup_path",
        "source_used_for_rotation",
        "rotation",
        "original_width",
        "original_height",
        "new_width",
        "new_height",
        "status",
        "notes",
    ]

    with manifest_path.open(
        "w",
        newline="",
        encoding="utf-8",
    ) as file:
        writer = csv.DictWriter(
            file,
            fieldnames=fieldnames,
        )

        writer.writeheader()
        writer.writerows(rows)

    return manifest_path


def write_readme(
    total_images: int,
    rotated_count: int,
    failed_count: int,
    dry_run: bool,
) -> Path:
    OUTPUT_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    readme_path = OUTPUT_ROOT / "README_Day4_Orientation_Fix.txt"

    text = f"""THIRD TRIAL DAY 4 ORIENTATION FIX
=================================

Created:
{datetime.now().strftime("%Y-%m-%d %H:%M:%S")}

Purpose:
Day 4 images were captured sideways compared with the other Third Trial days.
This script rotates Day 4 image files 90 degrees counter-clockwise so the tray
arrow faces forward/upward.

Input folder:
{DAY4_ROOT}

Rotation:
90 degrees counter-clockwise

Mode:
{"DRY RUN - no files were changed" if dry_run else "APPLY - files were rotated and replaced in place"}

Backup folder:
{BACKUP_ROOT}

Important:
If a backup already existed from a previous incomplete run, this script used the
backup image as the clean source before writing the rotated image back into the
Day 4 folder. This prevents double rotation.

Summary:
Total image files found: {total_images}
Successfully rotated: {rotated_count}
Failed: {failed_count}

Manifest:
{REPORTS_ROOT / "day4_orientation_fix_manifest.csv"}

Next step:
Open the Day 4 folder and check several D, F, MS_G, MS_NIR, MS_R, and MS_RE
images. The tray arrow should face forward/upward like the other days.
"""

    readme_path.write_text(
        text,
        encoding="utf-8",
    )

    return readme_path


# ============================================================
# 6) MAIN
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Rotate Third Trial Day 4 images 90 degrees counter-clockwise "
            "in place, with backup protection."
        )
    )

    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually rotate the files. Without this, only dry-run is performed.",
    )

    parser.add_argument(
        "--ignore-existing-backup",
        action="store_true",
        help=(
            "Use the current Day 4 file as the source even if a backup exists. "
            "Do not use this unless you are sure the current files are original."
        ),
    )

    args = parser.parse_args()

    print("\nSCRIPT 00 — FIX THIRD TRIAL DAY 4 ORIENTATION")
    print("=" * 70)
    print(f"Day 4 folder:\n{DAY4_ROOT}")
    print(f"\nOutput report folder:\n{OUTPUT_ROOT}")
    print("\nRotation: 90 degrees counter-clockwise")
    print(
        "\nMode: "
        + ("APPLY — files will be modified" if args.apply else "DRY RUN — no files modified")
    )

    images = collect_day4_images()

    if not images:
        print("\nNo image files were found in Day 4.")
        return 1

    print(f"\nImage files found: {len(images)}\n")

    rows = []
    rotated_count = 0
    failed_count = 0

    for index, image_path in enumerate(images, start=1):
        relative_path = relative_to_day4(image_path)

        row = {
            "file_no": index,
            "file_name": image_path.name,
            "relative_path": str(relative_path),
            "original_path": str(image_path),
            "backup_path": "",
            "source_used_for_rotation": "",
            "rotation": ROTATION_DIRECTION,
            "original_width": "",
            "original_height": "",
            "new_width": "",
            "new_height": "",
            "status": "",
            "notes": "",
        }

        if not args.apply:
            backup_path = backup_path_for(image_path)
            source_note = (
                "existing backup would be used"
                if backup_path.exists() and not args.ignore_existing_backup
                else "current Day 4 file would be used"
            )

            row["backup_path"] = str(backup_path)
            row["source_used_for_rotation"] = source_note
            row["status"] = "DRY_RUN"
            row["notes"] = "File would be rotated 90 degrees counter-clockwise."
            rows.append(row)

            print(
                f"{index:03d}: DRY_RUN | {relative_path} | {source_note}"
            )
            continue

        try:
            backup_path = ensure_backup(image_path)
            row["backup_path"] = str(backup_path)

            if backup_path.exists() and not args.ignore_existing_backup:
                source_path = backup_path
                row["source_used_for_rotation"] = "backup_original"
            else:
                source_path = image_path
                row["source_used_for_rotation"] = "current_day4_file"

            original_size, new_size = rotate_file_from_source(
                source_path=source_path,
                target_path=image_path,
            )

            row["original_width"] = original_size[0]
            row["original_height"] = original_size[1]
            row["new_width"] = new_size[0]
            row["new_height"] = new_size[1]
            row["status"] = "ROTATED"
            row["notes"] = (
                "Image was rotated 90 degrees counter-clockwise and replaced in place."
            )

            rotated_count += 1

            print(
                f"{index:03d}: ROTATED | {relative_path} | "
                f"{original_size[0]}x{original_size[1]} -> {new_size[0]}x{new_size[1]}"
            )

        except Exception as error:
            row["status"] = "FAILED"
            row["notes"] = str(error)
            failed_count += 1

            print(
                f"{index:03d}: FAILED | {relative_path} | {error}"
            )

        rows.append(row)

    manifest_path = write_manifest(rows)

    readme_path = write_readme(
        total_images=len(images),
        rotated_count=rotated_count,
        failed_count=failed_count,
        dry_run=not args.apply,
    )

    print("\n" + "=" * 70)

    if args.apply:
        print("DAY 4 ORIENTATION FIX FINISHED")
        print("=" * 70)
        print(f"ROTATED: {rotated_count}")
        print(f"FAILED: {failed_count}")
        print(f"\nBackup originals:\n{BACKUP_ROOT}")
    else:
        print("DRY RUN FINISHED")
        print("=" * 70)
        print("No files were modified.")

    print(f"\nManifest:\n{manifest_path}")
    print(f"\nREADME:\n{readme_path}")

    if failed_count > 0:
        print(
            "\nWARNING: Some files failed. Check the manifest before continuing."
        )
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())