from pathlib import Path
import re
from collections import deque

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import matplotlib.pyplot as plt
from matplotlib.colors import rgb_to_hsv


# ============================================================
# First Trial - 02 Germination Detection on Standardised Images
# STRICT VERSION
#
# This version reduces false positives by requiring:
# - stronger green colour
# - enough green pixels
# - one connected seedling-like green cluster
#
# Input:
# outputs/first_trial_germination/01_standardized_images/standardized/
#
# Output:
# outputs/first_trial_germination/02_germination_detection_strict/
# ============================================================


ROOT_DIR = Path.cwd()

INPUT_DIR = (
    ROOT_DIR
    / "outputs"
    / "first_trial_germination"
    / "01_standardized_images"
    / "standardized"
)

OUTPUT_DIR = (
    ROOT_DIR
    / "outputs"
    / "first_trial_germination"
    / "02_germination_detection_strict"
)

OVERLAY_DIR = OUTPUT_DIR / "overlays"
MASK_DIR = OUTPUT_DIR / "green_masks"

OVERLAY_DIR.mkdir(parents=True, exist_ok=True)
MASK_DIR.mkdir(parents=True, exist_ok=True)

CELL_CSV = OUTPUT_DIR / "cell_measurements_standardized_strict.csv"
SUMMARY_CSV = OUTPUT_DIR / "image_summary_standardized_strict.csv"
CUMULATIVE_CSV = OUTPUT_DIR / "image_summary_standardized_strict_cumulative.csv"

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}

ROWS = 7
COLS = 10
TOTAL_CELLS = ROWS * COLS

# Ignore outer border area inside each cell.
INNER_MARGIN_RATIO = 0.12

# Strict germination settings.
# If real seedlings are missed, lower these slightly.
MIN_GREEN_PIXELS = 15
MIN_GREEN_RATIO = 0.0015
MIN_MEAN_GREEN_STRENGTH = 5.0
MIN_LARGEST_GREEN_CLUSTER = 7


def safe_name(text):
    text = text.strip().lower()
    text = text.replace(" ", "_")
    text = re.sub(r"[^a-z0-9_]+", "_", text)
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def detect_treatment(filename):
    name = filename.lower()

    if "no_microbes" in name or "no microbes" in name or "nomicrobes" in name:
        return "No Microbes"

    if "microbes" in name or "microbe" in name:
        return "Microbes"

    return "Unknown"


def detect_day(filename):
    name = filename.lower()
    match = re.search(r"day[_\s-]*([0-9]+)", name)

    if match:
        return int(match.group(1))

    return None


def load_image_rgb(path):
    return np.array(Image.open(path).convert("RGB"))


def green_pixel_mask(roi):
    """
    Stricter RGB + HSV green detection.

    This is normal RGB image greenness detection.
    It is not NDVI.
    """

    rgb = roi.astype(float)
    rgb_norm = rgb / 255.0
    hsv = rgb_to_hsv(rgb_norm)

    red = rgb[:, :, 0]
    green = rgb[:, :, 1]
    blue = rgb[:, :, 2]

    hue = hsv[:, :, 0]
    saturation = hsv[:, :, 1]
    value = hsv[:, :, 2]

    excess_green = 2 * green - red - blue

    # Hue range focused on actual green / yellow-green seedling colour.
    hsv_rule = (
        (hue >= 0.16) &
        (hue <= 0.42) &
        (saturation >= 0.18) &
        (value >= 0.18)
    )

    # RGB rule: green must clearly dominate red and blue.
    rgb_rule = (
        (green >= 45) &
        (green > red * 1.04) &
        (green > blue * 1.04) &
        ((green - red) >= 6) &
        (excess_green >= 8)
    )

    mask = hsv_rule & rgb_rule

    return mask, excess_green


def largest_connected_component_size(mask):
    """
    Finds the largest connected group of green pixels.
    This reduces false positives from scattered green noise.
    """

    height, width = mask.shape
    visited = np.zeros_like(mask, dtype=bool)

    largest = 0

    for y in range(height):
        for x in range(width):
            if not mask[y, x] or visited[y, x]:
                continue

            queue = deque()
            queue.append((y, x))
            visited[y, x] = True
            size = 0

            while queue:
                cy, cx = queue.popleft()
                size += 1

                for ny in range(cy - 1, cy + 2):
                    for nx in range(cx - 1, cx + 2):
                        if ny == cy and nx == cx:
                            continue

                        if ny < 0 or ny >= height or nx < 0 or nx >= width:
                            continue

                        if mask[ny, nx] and not visited[ny, nx]:
                            visited[ny, nx] = True
                            queue.append((ny, nx))

            if size > largest:
                largest = size

    return largest


def analyse_cell(image, row_index, col_index):
    height, width = image.shape[:2]

    cell_width = width / COLS
    cell_height = height / ROWS

    x1 = int(col_index * cell_width)
    x2 = int((col_index + 1) * cell_width)
    y1 = int(row_index * cell_height)
    y2 = int((row_index + 1) * cell_height)

    margin_x = int(cell_width * INNER_MARGIN_RATIO)
    margin_y = int(cell_height * INNER_MARGIN_RATIO)

    roi_x1 = x1 + margin_x
    roi_x2 = x2 - margin_x
    roi_y1 = y1 + margin_y
    roi_y2 = y2 - margin_y

    roi = image[roi_y1:roi_y2, roi_x1:roi_x2]

    if roi.size == 0:
        return {
            "x1": x1,
            "y1": y1,
            "x2": x2,
            "y2": y2,
            "roi_x1": roi_x1,
            "roi_y1": roi_y1,
            "roi_x2": roi_x2,
            "roi_y2": roi_y2,
            "green_pixels": 0,
            "total_pixels": 0,
            "green_ratio": 0.0,
            "mean_green_strength": 0.0,
            "max_green_strength": 0.0,
            "largest_green_cluster": 0,
            "germinated_estimate": False,
        }

    mask, green_strength = green_pixel_mask(roi)

    green_pixels = int(mask.sum())
    total_pixels = int(mask.size)
    green_ratio = green_pixels / total_pixels

    if green_pixels > 0:
        mean_green_strength = float(green_strength[mask].mean())
        max_green_strength = float(green_strength[mask].max())
    else:
        mean_green_strength = 0.0
        max_green_strength = 0.0

    largest_cluster = largest_connected_component_size(mask)

    germinated = (
        green_pixels >= MIN_GREEN_PIXELS
        and green_ratio >= MIN_GREEN_RATIO
        and mean_green_strength >= MIN_MEAN_GREEN_STRENGTH
        and largest_cluster >= MIN_LARGEST_GREEN_CLUSTER
    )

    return {
        "x1": x1,
        "y1": y1,
        "x2": x2,
        "y2": y2,
        "roi_x1": roi_x1,
        "roi_y1": roi_y1,
        "roi_x2": roi_x2,
        "roi_y2": roi_y2,
        "green_pixels": green_pixels,
        "total_pixels": total_pixels,
        "green_ratio": green_ratio,
        "mean_green_strength": mean_green_strength,
        "max_green_strength": max_green_strength,
        "largest_green_cluster": largest_cluster,
        "germinated_estimate": germinated,
    }


def create_green_mask_image(image, cell_records):
    mask_image = np.zeros_like(image)

    for record in cell_records:
        roi_x1 = int(record["roi_x1"])
        roi_y1 = int(record["roi_y1"])
        roi_x2 = int(record["roi_x2"])
        roi_y2 = int(record["roi_y2"])

        roi = image[roi_y1:roi_y2, roi_x1:roi_x2]
        mask, _ = green_pixel_mask(roi)

        mask_roi = mask_image[roi_y1:roi_y2, roi_x1:roi_x2]
        mask_roi[mask] = [0, 255, 0]

    return mask_image


def save_overlay(image, cell_records, output_path, title_text):
    pil_image = Image.fromarray(image)
    draw = ImageDraw.Draw(pil_image)

    try:
        font = ImageFont.truetype("arial.ttf", 18)
        small_font = ImageFont.truetype("arial.ttf", 14)
    except:
        font = None
        small_font = None

    for record in cell_records:
        colour = "lime" if record["germinated_estimate"] else "red"

        x1 = int(record["x1"])
        y1 = int(record["y1"])
        x2 = int(record["x2"])
        y2 = int(record["y2"])

        roi_x1 = int(record["roi_x1"])
        roi_y1 = int(record["roi_y1"])
        roi_x2 = int(record["roi_x2"])
        roi_y2 = int(record["roi_y2"])

        draw.rectangle([x1, y1, x2, y2], outline="yellow", width=2)
        draw.rectangle([roi_x1, roi_y1, roi_x2, roi_y2], outline=colour, width=3)

        label = f'{record["row"]}-{record["col"]}'
        draw.text((x1 + 5, y1 + 5), label, fill="white", font=small_font)

    title_height = 70
    canvas = Image.new(
        "RGB",
        (pil_image.width, pil_image.height + title_height),
        "white"
    )

    canvas.paste(pil_image, (0, title_height))
    title_draw = ImageDraw.Draw(canvas)
    title_draw.text((20, 20), title_text, fill="black", font=font)

    canvas.save(output_path, quality=95)


def analyse_image(image_path):
    image = load_image_rgb(image_path)

    treatment = detect_treatment(image_path.stem)
    day = detect_day(image_path.stem)

    cell_records = []

    for row_index in range(ROWS):
        for col_index in range(COLS):
            result = analyse_cell(
                image=image,
                row_index=row_index,
                col_index=col_index
            )

            cell_records.append(
                {
                    "filename": image_path.name,
                    "treatment": treatment,
                    "day": day,
                    "row": row_index + 1,
                    "col": col_index + 1,
                    "cell_id": f"R{row_index + 1:02d}_C{col_index + 1:02d}",
                    **result,
                }
            )

    germinated_count = sum(
        1 for record in cell_records
        if record["germinated_estimate"]
    )

    germination_percent = germinated_count / TOTAL_CELLS * 100

    output_base = safe_name(image_path.stem)

    overlay_path = OVERLAY_DIR / f"{output_base}_strict_overlay.jpg"
    mask_path = MASK_DIR / f"{output_base}_strict_green_mask.jpg"

    title_text = (
        f"{treatment} | Day {day} | "
        f"Estimated germination: {germinated_count}/{TOTAL_CELLS} "
        f"({germination_percent:.2f}%)"
    )

    save_overlay(
        image=image,
        cell_records=cell_records,
        output_path=overlay_path,
        title_text=title_text
    )

    green_mask_image = create_green_mask_image(image, cell_records)
    Image.fromarray(green_mask_image).save(mask_path, quality=95)

    summary_record = {
        "filename": image_path.name,
        "treatment": treatment,
        "day": day,
        "total_cells": TOTAL_CELLS,
        "estimated_germinated_cells": germinated_count,
        "estimated_not_germinated_cells": TOTAL_CELLS - germinated_count,
        "estimated_germination_percent": round(germination_percent, 2),
        "mean_green_ratio": round(
            float(np.mean([r["green_ratio"] for r in cell_records])),
            6
        ),
        "overlay_filename": overlay_path.name,
        "green_mask_filename": mask_path.name,
    }

    return cell_records, summary_record


def make_cumulative_summary(summary_df):
    corrected_rows = []

    for treatment, group in summary_df.groupby("treatment"):
        group = group.sort_values("day").copy()
        previous_best = 0

        for _, row in group.iterrows():
            raw_count = int(row["estimated_germinated_cells"])
            total_cells = int(row["total_cells"])

            cumulative_count = max(previous_best, raw_count)
            cumulative_percent = cumulative_count / total_cells * 100

            row_dict = row.to_dict()
            row_dict["count_decreased_before_correction"] = raw_count < previous_best
            row_dict["cumulative_germinated_cells"] = cumulative_count
            row_dict["cumulative_germination_percent"] = round(
                cumulative_percent,
                2
            )

            corrected_rows.append(row_dict)
            previous_best = cumulative_count

    return pd.DataFrame(corrected_rows)


def make_quick_chart(cumulative_df):
    chart_path = OUTPUT_DIR / "first_trial_standardized_strict_germination_trend.png"

    plt.figure(figsize=(9, 6))

    for treatment, group in cumulative_df.groupby("treatment"):
        group = group.sort_values("day")

        plt.plot(
            group["day"],
            group["cumulative_germination_percent"],
            marker="o",
            linewidth=2,
            label=treatment
        )

    plt.title("First Trial: Germination Trend from Standardised Images")
    plt.xlabel("Day")
    plt.ylabel("Cumulative germination (%)")
    plt.ylim(0, 105)
    plt.xticks(sorted(cumulative_df["day"].dropna().astype(int).unique()))
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()

    plt.savefig(chart_path, dpi=300)
    plt.close()

    return chart_path


def main():
    if not INPUT_DIR.exists():
        raise FileNotFoundError(
            f"Standardised image folder not found:\n{INPUT_DIR}\n\n"
            "Run first_trial_01_standardise_tray_images.py first."
        )

    image_paths = [
        path for path in sorted(INPUT_DIR.iterdir())
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    ]

    if not image_paths:
        raise FileNotFoundError(
            f"No standardised images found in:\n{INPUT_DIR}"
        )

    print(f"Found {len(image_paths)} standardised images.")

    all_cell_records = []
    summary_records = []

    for image_path in image_paths:
        print(f"Processing: {image_path.name}")

        cell_records, summary_record = analyse_image(image_path)

        all_cell_records.extend(cell_records)
        summary_records.append(summary_record)

    cell_df = pd.DataFrame(all_cell_records)
    summary_df = pd.DataFrame(summary_records)

    summary_df = summary_df.sort_values(["treatment", "day"])

    cumulative_df = make_cumulative_summary(summary_df)
    cumulative_df = cumulative_df.sort_values(["treatment", "day"])

    cell_df.to_csv(CELL_CSV, index=False)
    summary_df.to_csv(SUMMARY_CSV, index=False)
    cumulative_df.to_csv(CUMULATIVE_CSV, index=False)

    chart_path = make_quick_chart(cumulative_df)

    print()
    print("Done.")
    print(f"Cell measurements saved to:\n{CELL_CSV}")
    print(f"Image summary saved to:\n{SUMMARY_CSV}")
    print(f"Cumulative summary saved to:\n{CUMULATIVE_CSV}")
    print(f"Overlays saved to:\n{OVERLAY_DIR}")
    print(f"Green masks saved to:\n{MASK_DIR}")
    print(f"Quick chart saved to:\n{chart_path}")
    print()
    print(cumulative_df[
        [
            "filename",
            "treatment",
            "day",
            "total_cells",
            "estimated_germinated_cells",
            "estimated_germination_percent",
            "cumulative_germinated_cells",
            "cumulative_germination_percent",
            "count_decreased_before_correction",
        ]
    ])


if __name__ == "__main__":
    main()