from pathlib import Path
import re
import cv2
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# First Trial - 01 Standardise Tray Images
# RINA / Orlar First Trial
#
# Purpose:
# - Read original tray photos
# - Manually click four logical corner CELL CENTRES
# - Correct rotated/angled tray images into one standard orientation
# - Save standardised images for all later germination analysis
#
# Tray layout:
# 7 rows x 10 columns = 70 cells
#
# IMPORTANT:
# Later scripts should use:
# outputs/first_trial_germination/01_standardized_images/standardized/
# ============================================================


ROOT_DIR = Path.cwd()

INPUT_DIR = ROOT_DIR / "data" / "First Trial (Two Trays)"

OUTPUT_DIR = ROOT_DIR / "outputs" / "first_trial_germination" / "01_standardized_images"
STANDARDIZED_DIR = OUTPUT_DIR / "standardized"
DIAGNOSTIC_DIR = OUTPUT_DIR / "diagnostics"

STANDARDIZED_DIR.mkdir(parents=True, exist_ok=True)
DIAGNOSTIC_DIR.mkdir(parents=True, exist_ok=True)

CLICK_CSV = OUTPUT_DIR / "clicked_corner_cell_centres.csv"
SUMMARY_CSV = OUTPUT_DIR / "standardized_image_summary.csv"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}

ROWS = 7
COLS = 10
TOTAL_CELLS = ROWS * COLS

# Size of each tray cell in the standardised output image.
# Larger value = higher output image resolution.
CELL_SIZE = 120

OUTPUT_WIDTH = COLS * CELL_SIZE
OUTPUT_HEIGHT = ROWS * CELL_SIZE


def safe_name(text):
    text = text.strip().lower()
    text = text.replace(" ", "_")
    text = re.sub(r"[^a-z0-9_]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def detect_treatment(filename):
    name = filename.lower()

    if "no microbes" in name or "no_microbes" in name or "nomicrobes" in name:
        return "No Microbes"

    if "microbes" in name or "microbe" in name:
        return "Microbes"

    return "Unknown"


def detect_day(filename):
    name = filename.lower()
    match = re.search(r"day\s*([0-9]+)", name)

    if match:
        return int(match.group(1))

    return None


def read_image_bgr(path):
    """
    Reads image safely on Windows paths.
    """
    data = np.fromfile(str(path), dtype=np.uint8)
    image = cv2.imdecode(data, cv2.IMREAD_COLOR)

    if image is None:
        raise ValueError(f"Could not read image: {path}")

    return image


def save_image(path, image_bgr):
    """
    Saves image safely on Windows paths.
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    success, encoded = cv2.imencode(".jpg", image_bgr, [cv2.IMWRITE_JPEG_QUALITY, 95])

    if not success:
        raise ValueError(f"Could not save image: {path}")

    encoded.tofile(str(path))


def bgr_to_rgb(image_bgr):
    return cv2.cvtColor(image_bgr, cv2.COLOR_BGR2RGB)


def get_existing_clicks():
    if CLICK_CSV.exists():
        return pd.read_csv(CLICK_CSV)

    return pd.DataFrame(
        columns=[
            "filename",
            "x_r1c1", "y_r1c1",
            "x_r1c10", "y_r1c10",
            "x_r7c10", "y_r7c10",
            "x_r7c1", "y_r7c1",
        ]
    )


def save_clicks(df):
    df.to_csv(CLICK_CSV, index=False)


def click_corner_cell_centres(image_bgr, filename):
    """
    Click four logical corner cell centres.

    The order is NOT based only on screen position.
    It is based on the tray's logical row-column layout.

    Click:
    1. Row 1, Column 1
    2. Row 1, Column 10
    3. Row 7, Column 10
    4. Row 7, Column 1
    """

    image_rgb = bgr_to_rgb(image_bgr)

    fig, ax = plt.subplots(figsize=(14, 9))
    ax.imshow(image_rgb)

    ax.set_title(
        f"{filename}\n\n"
        "Click the CENTRE of these four LOGICAL tray cells in order:\n"
        "1 = Row 1 Col 1,  2 = Row 1 Col 10,  "
        "3 = Row 7 Col 10,  4 = Row 7 Col 1\n\n"
        "Important: If the photo is rotated, still click the same LOGICAL cells."
    )

    ax.axis("on")

    points = plt.ginput(4, timeout=0)
    plt.close(fig)

    if len(points) != 4:
        raise ValueError("You must click exactly 4 points.")

    return np.array(points, dtype=np.float32)


def get_clicks_for_image(image_bgr, image_path, clicks_df):
    filename = image_path.name

    existing = clicks_df[clicks_df["filename"] == filename]

    if not existing.empty:
        row = existing.iloc[0]

        points = np.array(
            [
                [row["x_r1c1"], row["y_r1c1"]],
                [row["x_r1c10"], row["y_r1c10"]],
                [row["x_r7c10"], row["y_r7c10"]],
                [row["x_r7c1"], row["y_r7c1"]],
            ],
            dtype=np.float32
        )

        return points, clicks_df

    print()
    print(f"Click four logical corner CELL CENTRES for: {filename}")

    points = click_corner_cell_centres(image_bgr, filename)

    new_row = {
        "filename": filename,
        "x_r1c1": points[0][0],
        "y_r1c1": points[0][1],
        "x_r1c10": points[1][0],
        "y_r1c10": points[1][1],
        "x_r7c10": points[2][0],
        "y_r7c10": points[2][1],
        "x_r7c1": points[3][0],
        "y_r7c1": points[3][1],
    }

    clicks_df = pd.concat(
        [clicks_df, pd.DataFrame([new_row])],
        ignore_index=True
    )

    save_clicks(clicks_df)

    return points, clicks_df


def bilinear_point(corner_centres, u, v):
    """
    Bilinear interpolation/extrapolation.

    corner_centres order:
    1. R1C1
    2. R1C10
    3. R7C10
    4. R7C1

    u = left to right
    v = top to bottom
    """

    r1c1, r1c10, r7c10, r7c1 = corner_centres

    top = (1 - u) * r1c1 + u * r1c10
    bottom = (1 - u) * r7c1 + u * r7c10

    point = (1 - v) * top + v * bottom

    return point


def estimate_outer_grid_corners_from_cell_centres(corner_centres):
    """
    The clicked points are the centres of the four corner cells.
    To warp the full tray-cell area, we extrapolate outward by half a cell.

    For 10 columns, centre coordinates go from column 1 centre to column 10 centre.
    The full cell boundary extends half a spacing before col 1 and half after col 10.

    Same logic for 7 rows.
    """

    u_left = -0.5 / (COLS - 1)
    u_right = 1 + 0.5 / (COLS - 1)

    v_top = -0.5 / (ROWS - 1)
    v_bottom = 1 + 0.5 / (ROWS - 1)

    outer_top_left = bilinear_point(corner_centres, u_left, v_top)
    outer_top_right = bilinear_point(corner_centres, u_right, v_top)
    outer_bottom_right = bilinear_point(corner_centres, u_right, v_bottom)
    outer_bottom_left = bilinear_point(corner_centres, u_left, v_bottom)

    return np.array(
        [
            outer_top_left,
            outer_top_right,
            outer_bottom_right,
            outer_bottom_left,
        ],
        dtype=np.float32
    )


def standardise_image(image_bgr, corner_centres):
    """
    Warp the tray-cell area into a standard 10-column x 7-row image.
    """

    src_outer_corners = estimate_outer_grid_corners_from_cell_centres(corner_centres)

    dst_corners = np.array(
        [
            [0, 0],
            [OUTPUT_WIDTH - 1, 0],
            [OUTPUT_WIDTH - 1, OUTPUT_HEIGHT - 1],
            [0, OUTPUT_HEIGHT - 1],
        ],
        dtype=np.float32
    )

    matrix = cv2.getPerspectiveTransform(src_outer_corners, dst_corners)

    standardized = cv2.warpPerspective(
        image_bgr,
        matrix,
        (OUTPUT_WIDTH, OUTPUT_HEIGHT)
    )

    return standardized, src_outer_corners


def draw_diagnostic_original(image_bgr, corner_centres, outer_corners, output_path):
    """
    Save diagnostic image showing clicked centres and estimated tray-cell boundary.
    """

    diagnostic = image_bgr.copy()

    # Draw clicked corner cell centres.
    labels = ["R1C1", "R1C10", "R7C10", "R7C1"]

    for point, label in zip(corner_centres, labels):
        x, y = int(point[0]), int(point[1])
        cv2.circle(diagnostic, (x, y), 12, (0, 255, 255), -1)
        cv2.putText(
            diagnostic,
            label,
            (x + 10, y - 10),
            cv2.FONT_HERSHEY_SIMPLEX,
            0.8,
            (0, 255, 255),
            2,
            cv2.LINE_AA
        )

    # Draw outer boundary.
    pts = outer_corners.reshape((-1, 1, 2)).astype(np.int32)
    cv2.polylines(diagnostic, [pts], True, (0, 255, 0), 4)

    save_image(output_path, diagnostic)


def draw_standardized_grid(standardized_bgr, output_path):
    """
    Save standardised image with grid overlay.
    This lets you check whether every cell is aligned consistently.
    """

    diagnostic = standardized_bgr.copy()

    # Draw vertical grid lines.
    for c in range(COLS + 1):
        x = int(c * CELL_SIZE)
        cv2.line(diagnostic, (x, 0), (x, OUTPUT_HEIGHT), (0, 255, 255), 2)

    # Draw horizontal grid lines.
    for r in range(ROWS + 1):
        y = int(r * CELL_SIZE)
        cv2.line(diagnostic, (0, y), (OUTPUT_WIDTH, y), (0, 255, 255), 2)

    # Add row-column labels.
    for r in range(ROWS):
        for c in range(COLS):
            x = int(c * CELL_SIZE + CELL_SIZE / 2)
            y = int(r * CELL_SIZE + CELL_SIZE / 2)

            cv2.putText(
                diagnostic,
                f"{r + 1}-{c + 1}",
                (x - 22, y + 6),
                cv2.FONT_HERSHEY_SIMPLEX,
                0.45,
                (255, 255, 255),
                1,
                cv2.LINE_AA
            )

    save_image(output_path, diagnostic)


def main():
    if not INPUT_DIR.exists():
        raise FileNotFoundError(
            f"Input folder not found:\n{INPUT_DIR}\n\n"
            "Check that your original images are inside:\n"
            "data/First Trial (Two Trays)"
        )

    image_paths = [
        path for path in sorted(INPUT_DIR.iterdir())
        if (
            path.is_file()
            and path.suffix.lower() in IMAGE_EXTENSIONS
            and "overlay" not in path.name.lower()
            and "mask" not in path.name.lower()
        )
    ]

    if not image_paths:
        raise FileNotFoundError(
            f"No original images found in:\n{INPUT_DIR}"
        )

    print(f"Found {len(image_paths)} original images.")

    clicks_df = get_existing_clicks()
    summary_records = []

    for image_path in image_paths:
        print(f"Processing: {image_path.name}")

        image_bgr = read_image_bgr(image_path)

        treatment = detect_treatment(image_path.stem)
        day = detect_day(image_path.stem)

        corner_centres, clicks_df = get_clicks_for_image(
            image_bgr=image_bgr,
            image_path=image_path,
            clicks_df=clicks_df
        )

        standardized_bgr, outer_corners = standardise_image(
            image_bgr=image_bgr,
            corner_centres=corner_centres
        )

        treatment_safe = safe_name(treatment)
        day_safe = f"day_{day:02d}" if day is not None else "day_unknown"

        standardized_filename = f"{treatment_safe}_{day_safe}_standardized.jpg"
        diagnostic_original_filename = f"{treatment_safe}_{day_safe}_original_click_diagnostic.jpg"
        diagnostic_grid_filename = f"{treatment_safe}_{day_safe}_standardized_grid_diagnostic.jpg"

        standardized_path = STANDARDIZED_DIR / standardized_filename
        diagnostic_original_path = DIAGNOSTIC_DIR / diagnostic_original_filename
        diagnostic_grid_path = DIAGNOSTIC_DIR / diagnostic_grid_filename

        save_image(standardized_path, standardized_bgr)

        draw_diagnostic_original(
            image_bgr=image_bgr,
            corner_centres=corner_centres,
            outer_corners=outer_corners,
            output_path=diagnostic_original_path
        )

        draw_standardized_grid(
            standardized_bgr=standardized_bgr,
            output_path=diagnostic_grid_path
        )

        summary_records.append(
            {
                "original_filename": image_path.name,
                "treatment": treatment,
                "day": day,
                "rows": ROWS,
                "columns": COLS,
                "total_cells": TOTAL_CELLS,
                "standardized_filename": standardized_filename,
                "standardized_path": str(standardized_path),
                "original_click_diagnostic": diagnostic_original_filename,
                "standardized_grid_diagnostic": diagnostic_grid_filename,
                "output_width": OUTPUT_WIDTH,
                "output_height": OUTPUT_HEIGHT,
            }
        )

    summary_df = pd.DataFrame(summary_records)
    summary_df = summary_df.sort_values(["treatment", "day"])
    summary_df.to_csv(SUMMARY_CSV, index=False)

    print()
    print("Done.")
    print(f"Standardized images saved to:\n{STANDARDIZED_DIR}")
    print(f"Diagnostics saved to:\n{DIAGNOSTIC_DIR}")
    print(f"Summary CSV saved to:\n{SUMMARY_CSV}")
    print()
    print(summary_df)


if __name__ == "__main__":
    main()