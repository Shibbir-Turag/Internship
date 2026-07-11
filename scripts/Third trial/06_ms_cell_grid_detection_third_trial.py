from __future__ import annotations

"""
SCRIPT 06 — THIRD TRIAL INDEPENDENT MULTISPECTRAL 70-CELL GRID DETECTION

Purpose
-------
Detect a separate 7 × 10 = 70-cell grid for Trial 3 multispectral images.

This script is used after:
    Script 01 — Crop Dual Reference
    Script 03 — D/RGB Cell Grid Detection
    Script 04 — Visible Emergence
    Script 05 — RGB Growth/Treatment Comparison

This script DOES:
- read cropped multispectral images from Script 01
- find complete MS band sets: MS_G, MS_R, MS_RE, MS_NIR
- use MS_NIR as the reference band for grid detection
- detect the 70-cell grid independently from RGB coordinates
- save MS-specific cell centres and square ownership zones
- save MS grid overlays for visual checking
- save CSV, Excel, and JSON reports

This script DOES NOT:
- reuse D/RGB cell coordinates
- calculate NDVI or NDRE
- analyse emergence
- compare treatments
- generate a Word report

Why independent MS grid detection is needed
-------------------------------------------
The multispectral images can have a different resolution, alignment, lens
geometry, and crop geometry compared with the D/RGB image. Therefore, the
MS grid must be detected independently before calculating NDVI/NDRE in Script 07.

Main output folder
------------------
outputs/Third trial/06_MS_Cell_Grid_Detection
"""

import argparse
import csv
import json
import math
import re
from collections import defaultdict
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image, ImageOps


# ============================================================
# 1) PATHS — CHANGE PROJECT_ROOT ONLY IF NEEDED
# ============================================================

PROJECT_ROOT = Path(
    r"C:\Users\tshib\OneDrive\Desktop\Internship"
)

CROP_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "01_Crop_Dual_Reference"
)

OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "06_MS_Cell_Grid_Detection"
)

REPORTS_ROOT = OUTPUT_ROOT / "_reports"
CONFIG_ROOT = OUTPUT_ROOT / "_config"
MANUAL_POINTS_JSON = CONFIG_ROOT / "manual_ms_grid_points.json"


# ============================================================
# 2) GRID SETTINGS
# ============================================================

ROWS = 7
COLS = 10
EXPECTED_CELLS = ROWS * COLS

DAY_NAME_TO_ORDER = {
    "day 1": 1,
    "day 2": 2,
    "day 3": 3,
    "day 4": 4,
    "day 5": 5,
    "day 6": 6,
    "day 7": 7,
}

REQUIRED_MS_BANDS = [
    "G",
    "R",
    "RE",
    "NIR",
]

REFERENCE_BAND = "NIR"

# Square ownership zones.
# 0.90 means each square uses 90% of the median row/column spacing.
SQUARE_ZONE_RATIO = 0.90

# Automatic pass thresholds.
PASS_MIN_SUPPORTED_CELLS = 65
CHECK_MIN_SUPPORTED_CELLS = 55
MAX_MEAN_DISTANCE_RATIO_PASS = 0.22
MAX_MEAN_DISTANCE_RATIO_CHECK = 0.32
MAX_SPACING_CV_PASS = 0.18
MAX_SPACING_CV_CHECK = 0.28


# ============================================================
# 3) OPTIONAL TIFF READER
# ============================================================

try:
    import tifffile

    TIFFFILE_AVAILABLE = True
except Exception:
    TIFFFILE_AVAILABLE = False


# ============================================================
# 4) GENERAL HELPERS
# ============================================================

def natural_key(text: object):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", str(text))
    ]


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"


def parse_filter_list(value: str | None):
    if not value:
        return None

    return {
        item.strip().casefold()
        for item in value.split(",")
        if item.strip()
    }


def tray_number_from_name(name: str):
    match = re.search(r"(\d+)", name)
    return int(match.group(1)) if match else ""


def day_sort_key(folder: Path):
    return (
        DAY_NAME_TO_ORDER.get(folder.name.casefold(), 999),
        natural_key(folder.name),
    )


def record_key(day: str, tray: str, capture_id: str):
    return f"{day}|{tray}|{capture_id}"


def relative_path(path: Path | None, root: Path):
    if path is None:
        return ""

    try:
        return str(path.relative_to(root))
    except ValueError:
        return str(path)


def safe_float(value, default=math.nan):
    try:
        return float(value)
    except Exception:
        return default


def safe_int(value, default=""):
    try:
        return int(float(value))
    except Exception:
        return default


# ============================================================
# 5) FIND MULTISPECTRAL IMAGE SETS
# ============================================================

def parse_ms_file(path: Path):
    """
    Expected examples:
        DJI_20260629153749_0001_MS_G.TIF
        DJI_20260629153749_0001_MS_R.TIF
        DJI_20260629153749_0001_MS_RE.TIF
        DJI_20260629153749_0001_MS_NIR.TIF
    """

    if path.suffix.casefold() not in {
        ".tif",
        ".tiff",
    }:
        return None

    stem = path.stem.upper()

    match = re.match(
        r"^(?P<capture>.+)_MS_(?P<band>G|R|RE|NIR)$",
        stem,
    )

    if not match:
        return None

    return {
        "capture_id": match.group("capture"),
        "band": match.group("band"),
        "path": path,
    }


def find_ms_sets(tray_folder: Path):
    grouped = defaultdict(dict)

    for file in tray_folder.rglob("*"):
        if not file.is_file():
            continue

        parsed = parse_ms_file(file)

        if parsed is None:
            continue

        grouped[parsed["capture_id"]][parsed["band"]] = parsed["path"]

    rows = []

    for capture_id, bands in grouped.items():
        missing_bands = [
            band
            for band in REQUIRED_MS_BANDS
            if band not in bands
        ]

        rows.append(
            {
                "capture_id": capture_id,
                "bands": bands,
                "missing_bands": missing_bands,
                "complete": len(missing_bands) == 0,
            }
        )

    return sorted(
        rows,
        key=lambda item: natural_key(item["capture_id"]),
    )


def collect_jobs(days_filter=None, trays_filter=None):
    if not CROP_ROOT.exists():
        raise FileNotFoundError(
            f"Input crop folder not found:\n{CROP_ROOT}"
        )

    jobs = []

    day_folders = sorted(
        [
            folder
            for folder in CROP_ROOT.iterdir()
            if folder.is_dir()
            and folder.name.casefold() in DAY_NAME_TO_ORDER
        ],
        key=day_sort_key,
    )

    for day_folder in day_folders:
        if days_filter and day_folder.name.casefold() not in days_filter:
            continue

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
            if trays_filter and tray_folder.name.casefold() not in trays_filter:
                continue

            ms_sets = find_ms_sets(tray_folder)

            for ms_set in ms_sets:
                jobs.append(
                    {
                        "day": day_folder.name,
                        "day_order": DAY_NAME_TO_ORDER[day_folder.name.casefold()],
                        "tray": tray_folder.name,
                        "tray_no": tray_number_from_name(tray_folder.name),
                        "capture_id": ms_set["capture_id"],
                        "bands": ms_set["bands"],
                        "missing_bands": ms_set["missing_bands"],
                        "complete": ms_set["complete"],
                    }
                )

    return sorted(
        jobs,
        key=lambda item: (
            item["day_order"],
            int(item["tray_no"]),
            natural_key(item["capture_id"]),
        ),
    )


# ============================================================
# 6) IMAGE READING AND NORMALISATION
# ============================================================

def read_ms_image(path: Path) -> np.ndarray:
    if TIFFFILE_AVAILABLE:
        array = tifffile.imread(str(path))
    else:
        with Image.open(path) as image:
            image = ImageOps.exif_transpose(image)
            array = np.asarray(image)

    if array.ndim == 3:
        array = array[:, :, 0]

    return array.astype(np.float32)


def normalise_to_uint8(array: np.ndarray) -> np.ndarray:
    valid = array[np.isfinite(array)]

    if valid.size == 0:
        return np.zeros_like(array, dtype=np.uint8)

    low, high = np.percentile(
        valid,
        [1, 99],
    )

    if high <= low:
        low, high = float(valid.min()), float(valid.max())

    if high <= low:
        return np.zeros_like(array, dtype=np.uint8)

    scaled = (
        (array - low)
        / (high - low)
        * 255.0
    )

    scaled = np.clip(
        scaled,
        0,
        255,
    )

    return scaled.astype(np.uint8)


def make_preprocess_variants(gray: np.ndarray):
    variants = {}

    clahe = cv2.createCLAHE(
        clipLimit=2.0,
        tileGridSize=(8, 8),
    )

    enhanced = clahe.apply(gray)

    variants["gray"] = gray
    variants["enhanced"] = enhanced
    variants["inverted"] = cv2.bitwise_not(enhanced)
    variants["blurred"] = cv2.GaussianBlur(
        enhanced,
        (5, 5),
        0,
    )
    variants["inverted_blurred"] = cv2.GaussianBlur(
        cv2.bitwise_not(enhanced),
        (5, 5),
        0,
    )

    return variants


# ============================================================
# 7) CANDIDATE POINT DETECTION
# ============================================================

def non_max_suppression_points(
    points: list[tuple[float, float, float]],
    min_distance: float,
):
    if not points:
        return []

    points = sorted(
        points,
        key=lambda item: item[2],
        reverse=True,
    )

    kept = []

    for point in points:
        x, y, radius = point

        too_close = False

        for existing in kept:
            ex, ey, _er = existing

            distance = math.hypot(
                x - ex,
                y - ey,
            )

            if distance < min_distance:
                too_close = True
                break

        if not too_close:
            kept.append(point)

    return kept


def hough_circle_points(gray: np.ndarray):
    height, width = gray.shape[:2]

    estimated_spacing = min(
        width / COLS,
        height / ROWS,
    )

    min_dist = max(
        8,
        int(estimated_spacing * 0.55),
    )

    min_radius = max(
        3,
        int(estimated_spacing * 0.13),
    )

    max_radius = max(
        min_radius + 2,
        int(estimated_spacing * 0.48),
    )

    variants = make_preprocess_variants(gray)

    configs = [
        {
            "dp": 1.2,
            "param1": 45,
            "param2": 12,
        },
        {
            "dp": 1.2,
            "param1": 55,
            "param2": 16,
        },
        {
            "dp": 1.2,
            "param1": 65,
            "param2": 20,
        },
        {
            "dp": 1.5,
            "param1": 55,
            "param2": 15,
        },
        {
            "dp": 1.5,
            "param1": 70,
            "param2": 22,
        },
        {
            "dp": 1.8,
            "param1": 75,
            "param2": 20,
        },
    ]

    candidates = []

    for variant_name, image in variants.items():
        for config_index, config in enumerate(configs, start=1):
            circles = cv2.HoughCircles(
                image,
                cv2.HOUGH_GRADIENT,
                dp=config["dp"],
                minDist=min_dist,
                param1=config["param1"],
                param2=config["param2"],
                minRadius=min_radius,
                maxRadius=max_radius,
            )

            if circles is None:
                continue

            circles = np.round(
                circles[0, :]
            ).astype(float)

            points = [
                (
                    float(circle[0]),
                    float(circle[1]),
                    float(circle[2]),
                )
                for circle in circles
            ]

            points = non_max_suppression_points(
                points,
                min_distance=estimated_spacing * 0.40,
            )

            if len(points) < 20:
                continue

            candidates.append(
                {
                    "method": f"hough_{variant_name}_config_{config_index}",
                    "points": points,
                    "source_count": len(points),
                }
            )

    return candidates


def contour_blob_points(gray: np.ndarray):
    height, width = gray.shape[:2]

    estimated_spacing = min(
        width / COLS,
        height / ROWS,
    )

    min_area = math.pi * (estimated_spacing * 0.10) ** 2
    max_area = math.pi * (estimated_spacing * 0.55) ** 2

    variants = make_preprocess_variants(gray)

    candidates = []

    for variant_name, image in variants.items():
        blur = cv2.GaussianBlur(
            image,
            (5, 5),
            0,
        )

        threshold_methods = []

        _otsu_value, otsu = cv2.threshold(
            blur,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU,
        )

        threshold_methods.append(
            ("otsu", otsu)
        )

        adaptive = cv2.adaptiveThreshold(
            blur,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            41,
            3,
        )

        threshold_methods.append(
            ("adaptive", adaptive)
        )

        for threshold_name, binary in threshold_methods:
            contours, _hierarchy = cv2.findContours(
                binary,
                cv2.RETR_EXTERNAL,
                cv2.CHAIN_APPROX_SIMPLE,
            )

            points = []

            for contour in contours:
                area = cv2.contourArea(contour)

                if area < min_area or area > max_area:
                    continue

                perimeter = cv2.arcLength(
                    contour,
                    True,
                )

                if perimeter <= 0:
                    continue

                circularity = (
                    4.0
                    * math.pi
                    * area
                    / (perimeter * perimeter)
                )

                if circularity < 0.35:
                    continue

                moments = cv2.moments(contour)

                if moments["m00"] == 0:
                    continue

                cx = moments["m10"] / moments["m00"]
                cy = moments["m01"] / moments["m00"]

                radius = math.sqrt(
                    area / math.pi
                )

                points.append(
                    (
                        float(cx),
                        float(cy),
                        float(radius),
                    )
                )

            points = non_max_suppression_points(
                points,
                min_distance=estimated_spacing * 0.40,
            )

            if len(points) < 20:
                continue

            candidates.append(
                {
                    "method": f"contour_{variant_name}_{threshold_name}",
                    "points": points,
                    "source_count": len(points),
                }
            )

    return candidates


# ============================================================
# 8) GRID FITTING
# ============================================================

def kmeans_1d(
    values: np.ndarray,
    k: int,
    iterations: int = 40,
) -> np.ndarray:
    values = np.asarray(
        values,
        dtype=float,
    )

    if values.size == 0:
        return np.array([])

    if values.size < k:
        return np.linspace(
            values.min(),
            values.max(),
            k,
        )

    centres = np.percentile(
        values,
        np.linspace(
            0,
            100,
            k,
        ),
    )

    for _ in range(iterations):
        distances = np.abs(
            values[:, None]
            - centres[None, :]
        )

        labels = distances.argmin(axis=1)

        new_centres = centres.copy()

        for index in range(k):
            cluster_values = values[labels == index]

            if cluster_values.size > 0:
                new_centres[index] = cluster_values.mean()

        if np.allclose(
            new_centres,
            centres,
            atol=1e-3,
        ):
            break

        centres = new_centres

    return np.sort(centres)


def spacing_cv(centres: np.ndarray) -> float:
    if centres.size < 2:
        return 999.0

    diffs = np.diff(
        np.sort(centres)
    )

    if np.mean(diffs) <= 0:
        return 999.0

    return float(
        np.std(diffs)
        / np.mean(diffs)
    )


def build_grid_from_centres(
    x_centres: np.ndarray,
    y_centres: np.ndarray,
    image_width: int,
    image_height: int,
    square_ratio: float = SQUARE_ZONE_RATIO,
):
    x_centres = np.sort(
        np.asarray(
            x_centres,
            dtype=float,
        )
    )

    y_centres = np.sort(
        np.asarray(
            y_centres,
            dtype=float,
        )
    )

    if x_centres.size != COLS or y_centres.size != ROWS:
        return []

    x_spacing = float(
        np.median(
            np.diff(x_centres)
        )
    )

    y_spacing = float(
        np.median(
            np.diff(y_centres)
        )
    )

    square_side = min(
        x_spacing,
        y_spacing,
    ) * square_ratio

    half_side = square_side / 2.0

    rows = []

    cell_id = 1

    for row_index, y in enumerate(y_centres, start=1):
        for col_index, x in enumerate(x_centres, start=1):
            x0 = max(
                0.0,
                x - half_side,
            )
            y0 = max(
                0.0,
                y - half_side,
            )
            x1 = min(
                float(image_width),
                x + half_side,
            )
            y1 = min(
                float(image_height),
                y + half_side,
            )

            rows.append(
                {
                    "cell_id": cell_id,
                    "row": row_index,
                    "column": col_index,
                    "x": float(x),
                    "y": float(y),
                    "square_x0": float(x0),
                    "square_y0": float(y0),
                    "square_x1": float(x1),
                    "square_y1": float(y1),
                    "square_side": float(square_side),
                }
            )

            cell_id += 1

    return rows


def evaluate_grid_candidate(
    points: list[tuple[float, float, float]],
    method: str,
    image_width: int,
    image_height: int,
):
    if len(points) < 20:
        return None

    point_array = np.asarray(
        [
            [point[0], point[1]]
            for point in points
        ],
        dtype=float,
    )

    x_centres = kmeans_1d(
        point_array[:, 0],
        COLS,
    )

    y_centres = kmeans_1d(
        point_array[:, 1],
        ROWS,
    )

    grid_rows = build_grid_from_centres(
        x_centres,
        y_centres,
        image_width,
        image_height,
    )

    if len(grid_rows) != EXPECTED_CELLS:
        return None

    grid_points = np.asarray(
        [
            [row["x"], row["y"]]
            for row in grid_rows
        ],
        dtype=float,
    )

    x_spacing = float(
        np.median(
            np.diff(
                np.sort(x_centres)
            )
        )
    )

    y_spacing = float(
        np.median(
            np.diff(
                np.sort(y_centres)
            )
        )
    )

    typical_spacing = max(
        1.0,
        min(
            x_spacing,
            y_spacing,
        ),
    )

    distances = np.sqrt(
        (
            (
                point_array[:, None, :]
                - grid_points[None, :, :]
            )
            ** 2
        ).sum(axis=2)
    )

    nearest_grid_index = distances.argmin(axis=1)
    nearest_distance = distances.min(axis=1)

    support_radius = typical_spacing * 0.36

    supported = nearest_distance <= support_radius

    supported_grid_indices = set(
        nearest_grid_index[supported].tolist()
    )

    supported_cell_count = len(
        supported_grid_indices
    )

    if supported.any():
        mean_distance = float(
            nearest_distance[supported].mean()
        )
    else:
        mean_distance = 999.0

    mean_distance_ratio = (
        mean_distance
        / typical_spacing
    )

    row_cv = spacing_cv(
        y_centres
    )

    col_cv = spacing_cv(
        x_centres
    )

    max_cv = max(
        row_cv,
        col_cv,
    )

    missing_cells = EXPECTED_CELLS - supported_cell_count
    source_count_difference = abs(
        len(points)
        - EXPECTED_CELLS
    )

    score = (
        supported_cell_count * 100.0
        - missing_cells * 40.0
        - mean_distance_ratio * 120.0
        - max_cv * 120.0
        - source_count_difference * 1.5
    )

    if (
        supported_cell_count >= PASS_MIN_SUPPORTED_CELLS
        and mean_distance_ratio <= MAX_MEAN_DISTANCE_RATIO_PASS
        and max_cv <= MAX_SPACING_CV_PASS
    ):
        status = "PASS_AUTO"
    elif (
        supported_cell_count >= CHECK_MIN_SUPPORTED_CELLS
        and mean_distance_ratio <= MAX_MEAN_DISTANCE_RATIO_CHECK
        and max_cv <= MAX_SPACING_CV_CHECK
    ):
        status = "CHECK_AUTO"
    else:
        status = "FAIL"

    return {
        "method": method,
        "status": status,
        "grid_rows": grid_rows,
        "source_count": len(points),
        "supported_cell_count": supported_cell_count,
        "missing_cell_count": missing_cells,
        "mean_distance_px": mean_distance,
        "mean_distance_ratio": mean_distance_ratio,
        "row_spacing_cv": row_cv,
        "column_spacing_cv": col_cv,
        "max_spacing_cv": max_cv,
        "score": score,
        "x_spacing": x_spacing,
        "y_spacing": y_spacing,
        "typical_spacing": typical_spacing,
    }


def dimension_fallback_grid(
    image_width: int,
    image_height: int,
):
    x_margin = image_width * 0.075
    y_margin = image_height * 0.085

    x_centres = np.linspace(
        x_margin,
        image_width - x_margin,
        COLS,
    )

    y_centres = np.linspace(
        y_margin,
        image_height - y_margin,
        ROWS,
    )

    grid_rows = build_grid_from_centres(
        x_centres,
        y_centres,
        image_width,
        image_height,
    )

    return {
        "method": "dimension_fallback_even_grid",
        "status": "CHECK_AUTO",
        "grid_rows": grid_rows,
        "source_count": 0,
        "supported_cell_count": 0,
        "missing_cell_count": EXPECTED_CELLS,
        "mean_distance_px": math.nan,
        "mean_distance_ratio": math.nan,
        "row_spacing_cv": 0.0,
        "column_spacing_cv": 0.0,
        "max_spacing_cv": 0.0,
        "score": -9999.0,
        "x_spacing": float(np.median(np.diff(x_centres))),
        "y_spacing": float(np.median(np.diff(y_centres))),
        "typical_spacing": float(
            min(
                np.median(np.diff(x_centres)),
                np.median(np.diff(y_centres)),
            )
        ),
    }


def auto_detect_grid(
    gray: np.ndarray,
    use_dimension_fallback: bool,
):
    image_height, image_width = gray.shape[:2]

    raw_candidates = []

    raw_candidates.extend(
        hough_circle_points(gray)
    )

    raw_candidates.extend(
        contour_blob_points(gray)
    )

    evaluated = []

    for candidate in raw_candidates:
        result = evaluate_grid_candidate(
            points=candidate["points"],
            method=candidate["method"],
            image_width=image_width,
            image_height=image_height,
        )

        if result is not None:
            evaluated.append(result)

    if evaluated:
        evaluated = sorted(
            evaluated,
            key=lambda item: item["score"],
            reverse=True,
        )

        best = evaluated[0]

        best["candidate_count"] = len(evaluated)

        return best

    if use_dimension_fallback:
        fallback = dimension_fallback_grid(
            image_width,
            image_height,
        )

        fallback["candidate_count"] = 0

        return fallback

    return None


# ============================================================
# 9) MANUAL GRID SUPPORT
# ============================================================

def load_manual_points():
    if not MANUAL_POINTS_JSON.exists():
        return {}

    try:
        with MANUAL_POINTS_JSON.open(
            "r",
            encoding="utf-8",
        ) as file:
            return json.load(file)
    except Exception:
        return {}


def save_manual_points(data: dict):
    MANUAL_POINTS_JSON.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    with MANUAL_POINTS_JSON.open(
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(
            data,
            file,
            indent=2,
        )


def manual_grid_from_four_corners(
    points: list[list[float]],
    image_width: int,
    image_height: int,
):
    """
    Points must be clicked in this order:
        1. top-left cell centre
        2. top-right cell centre
        3. bottom-left cell centre
        4. bottom-right cell centre
    """

    if len(points) != 4:
        raise ValueError(
            "Manual grid requires exactly four clicked points."
        )

    top_left = np.asarray(
        points[0],
        dtype=float,
    )
    top_right = np.asarray(
        points[1],
        dtype=float,
    )
    bottom_left = np.asarray(
        points[2],
        dtype=float,
    )
    bottom_right = np.asarray(
        points[3],
        dtype=float,
    )

    centres = []

    for row_index in range(ROWS):
        row_fraction = (
            row_index
            / max(ROWS - 1, 1)
        )

        left_point = (
            top_left * (1.0 - row_fraction)
            + bottom_left * row_fraction
        )

        right_point = (
            top_right * (1.0 - row_fraction)
            + bottom_right * row_fraction
        )

        for col_index in range(COLS):
            col_fraction = (
                col_index
                / max(COLS - 1, 1)
            )

            point = (
                left_point * (1.0 - col_fraction)
                + right_point * col_fraction
            )

            centres.append(point)

    centres = np.asarray(centres)

    x_diffs = []
    y_diffs = []

    for row_index in range(ROWS):
        row_points = centres[
            row_index * COLS:
            (row_index + 1) * COLS
        ]

        for col_index in range(COLS - 1):
            x_diffs.append(
                np.linalg.norm(
                    row_points[col_index + 1]
                    - row_points[col_index]
                )
            )

    for col_index in range(COLS):
        col_points = centres[col_index::COLS]

        for row_index in range(ROWS - 1):
            y_diffs.append(
                np.linalg.norm(
                    col_points[row_index + 1]
                    - col_points[row_index]
                )
            )

    square_side = min(
        float(np.median(x_diffs)),
        float(np.median(y_diffs)),
    ) * SQUARE_ZONE_RATIO

    half_side = square_side / 2.0

    grid_rows = []

    cell_id = 1

    for row_index in range(ROWS):
        for col_index in range(COLS):
            point = centres[
                row_index * COLS
                + col_index
            ]

            x = float(point[0])
            y = float(point[1])

            grid_rows.append(
                {
                    "cell_id": cell_id,
                    "row": row_index + 1,
                    "column": col_index + 1,
                    "x": x,
                    "y": y,
                    "square_x0": max(0.0, x - half_side),
                    "square_y0": max(0.0, y - half_side),
                    "square_x1": min(float(image_width), x + half_side),
                    "square_y1": min(float(image_height), y + half_side),
                    "square_side": square_side,
                }
            )

            cell_id += 1

    return {
        "method": "manual_four_corner_cell_centres",
        "status": "PASS_MANUAL",
        "grid_rows": grid_rows,
        "source_count": 4,
        "supported_cell_count": EXPECTED_CELLS,
        "missing_cell_count": 0,
        "mean_distance_px": 0.0,
        "mean_distance_ratio": 0.0,
        "row_spacing_cv": 0.0,
        "column_spacing_cv": 0.0,
        "max_spacing_cv": 0.0,
        "score": 99999.0,
        "x_spacing": math.nan,
        "y_spacing": math.nan,
        "typical_spacing": square_side / SQUARE_ZONE_RATIO,
        "candidate_count": 1,
    }


def request_manual_points(
    gray: np.ndarray,
    title: str,
):
    import matplotlib.pyplot as plt

    print("\nManual MS grid correction required.")
    print("Click exactly four cell centres in this order:")
    print("1) top-left cell centre")
    print("2) top-right cell centre")
    print("3) bottom-left cell centre")
    print("4) bottom-right cell centre")
    print("Close the plot after clicking the four points.\n")

    figure, axis = plt.subplots(
        figsize=(10, 7),
    )

    axis.imshow(
        gray,
        cmap="gray",
    )

    axis.set_title(
        title
        + "\nClick: TL, TR, BL, BR cell centres",
    )

    axis.axis("off")

    points = plt.ginput(
        4,
        timeout=0,
    )

    plt.close(figure)

    if len(points) != 4:
        raise RuntimeError(
            "Manual correction cancelled or incomplete. Four points are required."
        )

    return [
        [
            float(x),
            float(y),
        ]
        for x, y in points
    ]


# ============================================================
# 10) OVERLAY OUTPUT
# ============================================================

def draw_grid_overlay(
    gray: np.ndarray,
    grid_rows: list[dict],
    output_path: Path,
    title: str,
    status: str,
):
    image_height, image_width = gray.shape[:2]

    bgr = cv2.cvtColor(
        gray,
        cv2.COLOR_GRAY2BGR,
    )

    if status == "PASS_AUTO":
        colour = (0, 200, 0)
    elif status == "PASS_MANUAL":
        colour = (255, 160, 0)
    elif status == "CHECK_AUTO":
        colour = (0, 220, 255)
    else:
        colour = (0, 0, 220)

    median_side = float(
        np.median(
            [
                row["square_side"]
                for row in grid_rows
            ]
        )
    ) if grid_rows else 30.0

    line_width = max(
        1,
        int(round(median_side / 35)),
    )

    font_scale = max(
        0.30,
        min(0.75, median_side / 95),
    )

    font_thickness = max(
        1,
        int(round(font_scale * 2)),
    )

    for row in grid_rows:
        x0 = max(
            0,
            int(round(row["square_x0"])),
        )
        y0 = max(
            0,
            int(round(row["square_y0"])),
        )
        x1 = min(
            image_width - 1,
            int(round(row["square_x1"])),
        )
        y1 = min(
            image_height - 1,
            int(round(row["square_y1"])),
        )

        x = int(round(row["x"]))
        y = int(round(row["y"]))

        cv2.rectangle(
            bgr,
            (x0, y0),
            (x1, y1),
            colour,
            line_width,
        )

        cv2.circle(
            bgr,
            (x, y),
            max(2, line_width),
            colour,
            thickness=-1,
        )

        cv2.putText(
            bgr,
            str(row["cell_id"]),
            (
                max(0, x - 8),
                min(image_height - 2, y + 5),
            ),
            cv2.FONT_HERSHEY_SIMPLEX,
            font_scale,
            colour,
            font_thickness,
            cv2.LINE_AA,
        )

    header_height = max(
        45,
        int(round(median_side * 0.45)),
    )

    cv2.rectangle(
        bgr,
        (0, 0),
        (image_width, header_height),
        (255, 255, 255),
        thickness=-1,
    )

    cv2.putText(
        bgr,
        f"{title} | {status} | Cells: {len(grid_rows)}",
        (12, int(header_height * 0.65)),
        cv2.FONT_HERSHEY_SIMPLEX,
        max(0.42, min(0.80, font_scale)),
        (0, 0, 0),
        2,
        cv2.LINE_AA,
    )

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    cv2.imwrite(
        str(output_path),
        bgr,
        [
            cv2.IMWRITE_JPEG_QUALITY,
            95,
        ],
    )


# ============================================================
# 11) REPORT WRITING
# ============================================================

def write_csv(
    path: Path,
    rows: list[dict],
):
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    if not rows:
        path.write_text(
            "",
            encoding="utf-8",
        )
        return

    fieldnames = list(rows[0].keys())

    with path.open(
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


def style_excel(path: Path):
    workbook = load_workbook(path)

    header_fill = PatternFill(
        "solid",
        fgColor="1F4E78",
    )

    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 34

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = Font(
                color="FFFFFF",
                bold=True,
            )
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )

        for row in worksheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )

        for column_cells in worksheet.columns:
            letter = column_cells[0].column_letter

            longest = max(
                len(str(cell.value))
                if cell.value is not None
                else 0
                for cell in column_cells
            )

            worksheet.column_dimensions[letter].width = min(
                max(12, longest + 2),
                58,
            )

    workbook.save(path)


def write_reports(
    manifest_rows: list[dict],
    coordinate_rows: list[dict],
    settings: dict,
):
    REPORTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    CONFIG_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    manifest_csv = REPORTS_ROOT / "ms_grid_manifest.csv"
    coordinates_csv = REPORTS_ROOT / "ms_cell_coordinates.csv"
    excel_path = REPORTS_ROOT / "ms_cell_grid_report.xlsx"
    settings_path = CONFIG_ROOT / "ms_grid_detection_settings.json"

    write_csv(
        manifest_csv,
        manifest_rows,
    )

    write_csv(
        coordinates_csv,
        coordinate_rows,
    )

    manifest_frame = pd.DataFrame(manifest_rows)
    coordinates_frame = pd.DataFrame(coordinate_rows)

    readme = pd.DataFrame(
        {
            "Notes": [
                "This workbook contains independent multispectral 70-cell grid detection results.",
                "MS_NIR is used as the reference band for grid detection.",
                "D/RGB coordinates are not reused.",
                "PASS_AUTO and PASS_MANUAL are normally acceptable for Script 07.",
                "CHECK_AUTO rows should be visually inspected before use.",
                "FAIL rows need manual correction or recropping before Script 07.",
                "Each accepted MS image set should have exactly 70 coordinate rows.",
            ]
        }
    )

    with pd.ExcelWriter(
        excel_path,
        engine="openpyxl",
    ) as writer:
        manifest_frame.to_excel(
            writer,
            sheet_name="MS Grid Manifest",
            index=False,
        )

        coordinates_frame.to_excel(
            writer,
            sheet_name="MS Cell Coordinates",
            index=False,
        )

        readme.to_excel(
            writer,
            sheet_name="Read Me",
            index=False,
        )

    style_excel(
        excel_path,
    )

    with settings_path.open(
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(
            settings,
            file,
            indent=2,
        )

    return {
        "manifest_csv": manifest_csv,
        "coordinates_csv": coordinates_csv,
        "excel_path": excel_path,
        "settings_path": settings_path,
    }


# ============================================================
# 12) MAIN PROCESSING
# ============================================================

def process_job(
    job: dict,
    args,
    manual_points_store: dict,
):
    day = job["day"]
    day_order = int(job["day_order"])
    tray = job["tray"]
    tray_no = int(job["tray_no"])
    capture_id = job["capture_id"]

    key = record_key(
        day,
        tray,
        capture_id,
    )

    base_output = OUTPUT_ROOT / day / tray

    overlay_path = (
        base_output
        / "overlays"
        / f"{capture_id}_MS_NIR_grid_overlay.jpg"
    )

    if not job["complete"]:
        missing = ", ".join(
            job["missing_bands"]
        )

        manifest = {
            "day_order": day_order,
            "day": day,
            "tray": tray,
            "tray_no": tray_no,
            "capture_id": capture_id,
            "status": "SKIPPED_INCOMPLETE_MS_SET",
            "method": "",
            "reference_band": REFERENCE_BAND,
            "source_count": "",
            "supported_cell_count": "",
            "missing_cell_count": "",
            "mean_distance_px": "",
            "mean_distance_ratio": "",
            "row_spacing_cv": "",
            "column_spacing_cv": "",
            "max_spacing_cv": "",
            "score": "",
            "candidate_count": "",
            "cell_count": 0,
            "overlay_path": "",
            "reference_image_path": "",
            "missing_bands": missing,
            "notes": f"Incomplete MS set. Missing bands: {missing}",
        }

        return manifest, []

    reference_path = job["bands"][REFERENCE_BAND]

    raw = read_ms_image(
        reference_path,
    )

    gray = normalise_to_uint8(
        raw,
    )

    image_height, image_width = gray.shape[:2]

    result = None

    manual_key_exists = key in manual_points_store

    if args.force_manual or manual_key_exists:
        if manual_key_exists and not args.force_manual:
            manual_points = manual_points_store[key]
        else:
            manual_points = request_manual_points(
                gray,
                title=f"{day} | {tray} | {capture_id} | MS_NIR",
            )

            manual_points_store[key] = manual_points

            save_manual_points(
                manual_points_store,
            )

        result = manual_grid_from_four_corners(
            manual_points,
            image_width=image_width,
            image_height=image_height,
        )

    else:
        result = auto_detect_grid(
            gray,
            use_dimension_fallback=args.dimension_fallback,
        )

        if result is None:
            result = {
                "method": "no_valid_auto_candidate",
                "status": "FAIL",
                "grid_rows": [],
                "source_count": 0,
                "supported_cell_count": 0,
                "missing_cell_count": EXPECTED_CELLS,
                "mean_distance_px": math.nan,
                "mean_distance_ratio": math.nan,
                "row_spacing_cv": math.nan,
                "column_spacing_cv": math.nan,
                "max_spacing_cv": math.nan,
                "score": -99999.0,
                "x_spacing": math.nan,
                "y_spacing": math.nan,
                "typical_spacing": math.nan,
                "candidate_count": 0,
            }

        if (
            args.manual_fallback
            and result["status"] in {
                "FAIL",
                "CHECK_AUTO",
            }
        ):
            manual_points = request_manual_points(
                gray,
                title=f"{day} | {tray} | {capture_id} | MS_NIR",
            )

            manual_points_store[key] = manual_points

            save_manual_points(
                manual_points_store,
            )

            result = manual_grid_from_four_corners(
                manual_points,
                image_width=image_width,
                image_height=image_height,
            )

    grid_rows = result["grid_rows"]

    if grid_rows and (args.overwrite or not overlay_path.exists()):
        draw_grid_overlay(
            gray,
            grid_rows,
            overlay_path,
            title=f"{day} | {tray} | {capture_id} | MS_NIR",
            status=result["status"],
        )

    needs_review = result["status"] not in {
        "PASS_AUTO",
        "PASS_MANUAL",
    }

    coordinate_rows = []

    if grid_rows:
        for row in grid_rows:
            coordinate_rows.append(
                {
                    "day_order": day_order,
                    "day": day,
                    "tray": tray,
                    "tray_no": tray_no,
                    "capture_id": capture_id,
                    "reference_band": REFERENCE_BAND,
                    "cell_id": row["cell_id"],
                    "row": row["row"],
                    "column": row["column"],
                    "x": row["x"],
                    "y": row["y"],
                    "square_x0": row["square_x0"],
                    "square_y0": row["square_y0"],
                    "square_x1": row["square_x1"],
                    "square_y1": row["square_y1"],
                    "square_side": row["square_side"],
                    "coordinate_source": result["method"],
                    "grid_status": result["status"],
                    "needs_review": yes_no(needs_review),
                    "image_width": image_width,
                    "image_height": image_height,
                    "ms_g_path": str(job["bands"].get("G", "")),
                    "ms_r_path": str(job["bands"].get("R", "")),
                    "ms_re_path": str(job["bands"].get("RE", "")),
                    "ms_nir_path": str(job["bands"].get("NIR", "")),
                }
            )

    manifest = {
        "day_order": day_order,
        "day": day,
        "tray": tray,
        "tray_no": tray_no,
        "capture_id": capture_id,
        "status": result["status"],
        "method": result["method"],
        "reference_band": REFERENCE_BAND,
        "source_count": safe_int(result.get("source_count")),
        "supported_cell_count": safe_int(result.get("supported_cell_count")),
        "missing_cell_count": safe_int(result.get("missing_cell_count")),
        "mean_distance_px": result.get("mean_distance_px", ""),
        "mean_distance_ratio": result.get("mean_distance_ratio", ""),
        "row_spacing_cv": result.get("row_spacing_cv", ""),
        "column_spacing_cv": result.get("column_spacing_cv", ""),
        "max_spacing_cv": result.get("max_spacing_cv", ""),
        "score": result.get("score", ""),
        "candidate_count": safe_int(result.get("candidate_count")),
        "cell_count": len(coordinate_rows),
        "overlay_path": relative_path(overlay_path, OUTPUT_ROOT) if grid_rows else "",
        "reference_image_path": str(reference_path),
        "missing_bands": "",
        "notes": (
            "Review overlay before Script 07."
            if result["status"] == "CHECK_AUTO"
            else (
                "Manual correction or recropping required."
                if result["status"] == "FAIL"
                else ""
            )
        ),
    }

    return manifest, coordinate_rows


def run_analysis(args):
    days_filter = parse_filter_list(
        args.days,
    )

    trays_filter = parse_filter_list(
        args.trays,
    )

    jobs = collect_jobs(
        days_filter=days_filter,
        trays_filter=trays_filter,
    )

    print("\nSCRIPT 06 — THIRD TRIAL MS CELL GRID DETECTION")
    print("=" * 72)
    print(f"Crop folder:\n{CROP_ROOT}")
    print(f"\nOutput folder:\n{OUTPUT_ROOT}")
    print(f"\nJobs found: {len(jobs)}\n")

    for job in jobs:
        status = "complete" if job["complete"] else "incomplete"
        missing = (
            ""
            if job["complete"]
            else f" | missing: {', '.join(job['missing_bands'])}"
        )

        print(
            f"{job['day']} > {job['tray']} > {job['capture_id']} "
            f"({status}){missing}"
        )

    if args.dry_run:
        print("\nDry run complete. No outputs created.")
        return 0

    if not jobs:
        print("\nNo multispectral image sets were found.")
        return 1

    manual_points_store = load_manual_points()

    manifest_rows = []
    coordinate_rows = []

    for job in jobs:
        print(
            f"\nProcessing: {job['day']} > {job['tray']} > {job['capture_id']}"
        )

        manifest, coords = process_job(
            job,
            args,
            manual_points_store,
        )

        manifest_rows.append(
            manifest,
        )

        coordinate_rows.extend(
            coords,
        )

        print(
            f"{manifest['status']} | cells={manifest['cell_count']} | "
            f"method={manifest['method']} | supported={manifest['supported_cell_count']}"
        )

    settings = {
        "purpose": "Third Trial independent multispectral 70-cell grid detection",
        "crop_root": str(CROP_ROOT),
        "output_root": str(OUTPUT_ROOT),
        "rows": ROWS,
        "columns": COLS,
        "expected_cells": EXPECTED_CELLS,
        "required_ms_bands": REQUIRED_MS_BANDS,
        "reference_band": REFERENCE_BAND,
        "square_zone_ratio": SQUARE_ZONE_RATIO,
        "pass_min_supported_cells": PASS_MIN_SUPPORTED_CELLS,
        "check_min_supported_cells": CHECK_MIN_SUPPORTED_CELLS,
        "max_mean_distance_ratio_pass": MAX_MEAN_DISTANCE_RATIO_PASS,
        "max_mean_distance_ratio_check": MAX_MEAN_DISTANCE_RATIO_CHECK,
        "max_spacing_cv_pass": MAX_SPACING_CV_PASS,
        "max_spacing_cv_check": MAX_SPACING_CV_CHECK,
        "manual_points_json": str(MANUAL_POINTS_JSON),
        "dimension_fallback_enabled": bool(args.dimension_fallback),
        "manual_fallback_enabled": bool(args.manual_fallback),
        "note": (
            "PASS_AUTO and PASS_MANUAL are acceptable for Script 07. "
            "CHECK_AUTO should be visually inspected. FAIL needs correction."
        ),
    }

    output_paths = write_reports(
        manifest_rows,
        coordinate_rows,
        settings,
    )

    status_counts = (
        pd.DataFrame(manifest_rows)["status"]
        .value_counts()
        .to_dict()
        if manifest_rows
        else {}
    )

    print("\n" + "=" * 72)
    print("SCRIPT 06 FINISHED")
    print("=" * 72)

    for status, count in sorted(status_counts.items()):
        print(f"{status}: {count}")

    print(f"\nManifest:\n{output_paths['manifest_csv']}")
    print(f"\nCoordinates:\n{output_paths['coordinates_csv']}")
    print(f"\nExcel report:\n{output_paths['excel_path']}")
    print(f"\nSettings:\n{output_paths['settings_path']}")
    print(f"\nManual points file:\n{MANUAL_POINTS_JSON}")

    pass_count = sum(
        1
        for row in manifest_rows
        if row["status"] in {
            "PASS_AUTO",
            "PASS_MANUAL",
        }
    )

    expected_jobs = 84

    print(
        f"\nAccepted MS grids: {pass_count}/{len(manifest_rows)} "
        f"(expected full Trial 3 target: {expected_jobs})"
    )

    problematic = [
        row
        for row in manifest_rows
        if row["status"] not in {
            "PASS_AUTO",
            "PASS_MANUAL",
        }
    ]

    if problematic:
        print(
            "\nSome MS grids need checking before Script 07. "
            "Open ms_cell_grid_report.xlsx and inspect CHECK_AUTO/FAIL rows."
        )
        return 1

    return 0


# ============================================================
# 13) CLI
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Trial 3 Script 06: independent multispectral 70-cell grid detection."
        )
    )

    parser.add_argument(
        "--days",
        help='Process selected days only. Example: --days "Day 1,Day 7"',
    )

    parser.add_argument(
        "--trays",
        help='Process selected trays only. Example: --trays "Tray 1,Tray 12"',
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite existing overlay images.",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="List multispectral jobs without processing.",
    )

    parser.add_argument(
        "--manual-fallback",
        action="store_true",
        help=(
            "If automatic detection gives CHECK_AUTO or FAIL, open a manual "
            "click window for four corner cell centres."
        ),
    )

    parser.add_argument(
        "--force-manual",
        action="store_true",
        help=(
            "Force manual four-corner correction for each processed image. "
            "Use with --days/--trays for targeted correction."
        ),
    )

    parser.add_argument(
        "--dimension-fallback",
        action="store_true",
        help=(
            "Create an even dimension-based CHECK_AUTO grid if no automatic "
            "candidate is found. This must be visually reviewed."
        ),
    )

    args = parser.parse_args()

    return run_analysis(args)


if __name__ == "__main__":
    raise SystemExit(main())