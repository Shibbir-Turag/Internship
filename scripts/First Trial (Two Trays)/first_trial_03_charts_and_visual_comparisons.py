from pathlib import Path
import re
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image


# ============================================================
# First Trial - 03 Charts and Visual Comparisons
# RINA / Orlar First Trial
#
# Uses strict germination detection results from:
# outputs/first_trial_germination/02_germination_detection_strict/
#
# Creates:
# - cumulative germination chart
# - raw automated germination chart
# - germinated cell count chart
# - final Microbes vs No Microbes bar chart
# - side-by-side visual comparisons
# - side-by-side overlay comparisons
# - side-by-side green mask comparisons
# ============================================================


ROOT_DIR = Path.cwd()

BASE_OUTPUT_DIR = ROOT_DIR / "outputs" / "first_trial_germination"

STANDARDIZED_DIR = (
    BASE_OUTPUT_DIR
    / "01_standardized_images"
    / "standardized"
)

STRICT_DIR = (
    BASE_OUTPUT_DIR
    / "02_germination_detection_strict"
)

STRICT_OVERLAY_DIR = STRICT_DIR / "overlays"
STRICT_MASK_DIR = STRICT_DIR / "green_masks"

SUMMARY_CSV = STRICT_DIR / "image_summary_standardized_strict_cumulative.csv"
RAW_SUMMARY_CSV = STRICT_DIR / "image_summary_standardized_strict.csv"

OUTPUT_DIR = BASE_OUTPUT_DIR / "03_charts_and_visual_comparisons"
CHART_DIR = OUTPUT_DIR / "charts"
VISUAL_DIR = OUTPUT_DIR / "visual_comparisons"

CHART_DIR.mkdir(parents=True, exist_ok=True)
VISUAL_DIR.mkdir(parents=True, exist_ok=True)

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp"}


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


def load_summary():
    if not SUMMARY_CSV.exists():
        raise FileNotFoundError(
            f"Could not find:\n{SUMMARY_CSV}\n\n"
            "Run first_trial_02_detect_germination_standardized.py first."
        )

    df = pd.read_csv(SUMMARY_CSV)
    df["day"] = df["day"].astype(int)

    required = [
        "filename",
        "treatment",
        "day",
        "total_cells",
        "estimated_germinated_cells",
        "estimated_germination_percent",
        "cumulative_germinated_cells",
        "cumulative_germination_percent",
    ]

    missing = [col for col in required if col not in df.columns]

    if missing:
        raise ValueError(f"Missing columns in summary CSV: {missing}")

    return df.sort_values(["treatment", "day"])


def load_standardized_images():
    if not STANDARDIZED_DIR.exists():
        raise FileNotFoundError(
            f"Could not find standardised image folder:\n{STANDARDIZED_DIR}"
        )

    records = []

    for path in sorted(STANDARDIZED_DIR.iterdir()):
        if not path.is_file():
            continue

        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue

        treatment = detect_treatment(path.stem)
        day = detect_day(path.stem)

        if treatment == "Unknown" or day is None:
            continue

        records.append(
            {
                "path": path,
                "filename": path.name,
                "treatment": treatment,
                "day": day,
            }
        )

    image_df = pd.DataFrame(records)

    if image_df.empty:
        raise FileNotFoundError(
            f"No standardised Microbes / No Microbes images found in:\n{STANDARDIZED_DIR}"
        )

    return image_df


def find_matching_image(folder, treatment, day, keyword=None):
    if not folder.exists():
        return None

    treatment_key = safe_name(treatment)
    treatment_key_compact = treatment_key.replace("_", "")

    day_patterns = [
        f"day_{day}",
        f"day_{day:02d}",
        f"day{day}",
    ]

    candidates = []

    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue

        if path.suffix.lower() not in IMAGE_EXTENSIONS:
            continue

        name = path.stem.lower()
        compact_name = name.replace("_", "")

        if treatment_key_compact not in compact_name:
            continue

        if not any(pattern in name for pattern in day_patterns):
            continue

        if keyword is not None and keyword not in name:
            continue

        candidates.append(path)

    if not candidates:
        return None

    return candidates[0]


def get_summary_text(summary_df, treatment, day):
    rows = summary_df[
        (summary_df["treatment"] == treatment) &
        (summary_df["day"].astype(int) == int(day))
    ]

    if rows.empty:
        return ""

    row = rows.iloc[0]

    raw_count = int(row["estimated_germinated_cells"])
    cumulative_count = int(row["cumulative_germinated_cells"])
    cumulative_percent = float(row["cumulative_germination_percent"])
    total_cells = int(row["total_cells"])

    return (
        f"Raw: {raw_count}/{total_cells} | "
        f"Cumulative: {cumulative_count}/{total_cells} "
        f"({cumulative_percent:.2f}%)"
    )


def make_raw_germination_chart(df):
    plt.figure(figsize=(9, 6))

    for treatment, group in df.groupby("treatment"):
        group = group.sort_values("day")
        plt.plot(
            group["day"],
            group["estimated_germination_percent"],
            marker="o",
            linewidth=2,
            label=treatment,
        )

    plt.title("First Trial: Raw Automated Germination Percentage")
    plt.xlabel("Day")
    plt.ylabel("Raw automated germination (%)")
    plt.ylim(0, 105)
    plt.xticks(sorted(df["day"].unique()))
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()

    output_path = CHART_DIR / "first_trial_raw_automated_germination_percentage.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_cumulative_germination_chart(df):
    plt.figure(figsize=(9, 6))

    for treatment, group in df.groupby("treatment"):
        group = group.sort_values("day")
        plt.plot(
            group["day"],
            group["cumulative_germination_percent"],
            marker="o",
            linewidth=2,
            label=treatment,
        )

    plt.title("First Trial: Cumulative Germination Percentage")
    plt.xlabel("Day")
    plt.ylabel("Cumulative germination (%)")
    plt.ylim(0, 105)
    plt.xticks(sorted(df["day"].unique()))
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()

    output_path = CHART_DIR / "first_trial_cumulative_germination_percentage.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_germinated_cell_count_chart(df):
    plt.figure(figsize=(9, 6))

    for treatment, group in df.groupby("treatment"):
        group = group.sort_values("day")
        plt.plot(
            group["day"],
            group["cumulative_germinated_cells"],
            marker="o",
            linewidth=2,
            label=treatment,
        )

    max_cells = int(df["total_cells"].max())

    plt.title("First Trial: Cumulative Germinated Cell Count")
    plt.xlabel("Day")
    plt.ylabel("Cumulative germinated cells")
    plt.ylim(0, max_cells + 5)
    plt.xticks(sorted(df["day"].unique()))
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.tight_layout()

    output_path = CHART_DIR / "first_trial_cumulative_germinated_cells.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_final_comparison_chart(df):
    final_df = (
        df.sort_values("day")
        .groupby("treatment", as_index=False)
        .tail(1)
        .sort_values("treatment")
    )

    plt.figure(figsize=(8, 6))

    plt.bar(
        final_df["treatment"],
        final_df["cumulative_germination_percent"],
    )

    plt.title("First Trial: Final Germination Comparison")
    plt.xlabel("Treatment")
    plt.ylabel("Final cumulative germination (%)")
    plt.ylim(0, 105)
    plt.grid(axis="y", alpha=0.3)
    plt.tight_layout()

    output_path = CHART_DIR / "first_trial_final_microbes_vs_no_microbes_comparison.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_day_image_comparison(image_df, summary_df, day):
    microbes = image_df[
        (image_df["treatment"] == "Microbes") &
        (image_df["day"] == day)
    ]

    no_microbes = image_df[
        (image_df["treatment"] == "No Microbes") &
        (image_df["day"] == day)
    ]

    if microbes.empty or no_microbes.empty:
        print(f"Skipping original image comparison for Day {day}: missing image.")
        return None

    microbes_img = Image.open(microbes.iloc[0]["path"]).convert("RGB")
    no_microbes_img = Image.open(no_microbes.iloc[0]["path"]).convert("RGB")

    fig, axes = plt.subplots(1, 2, figsize=(16, 8))

    axes[0].imshow(microbes_img)
    axes[0].set_title(
        f"Microbes - Day {day}\n"
        f"{get_summary_text(summary_df, 'Microbes', day)}"
    )
    axes[0].axis("off")

    axes[1].imshow(no_microbes_img)
    axes[1].set_title(
        f"No Microbes - Day {day}\n"
        f"{get_summary_text(summary_df, 'No Microbes', day)}"
    )
    axes[1].axis("off")

    fig.suptitle(
        f"First Trial Standardised Image Comparison - Day {day}",
        fontsize=16,
    )

    plt.tight_layout()

    output_path = VISUAL_DIR / f"first_trial_day_{day:02d}_standardized_microbes_vs_no_microbes.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_day_output_comparison(summary_df, day, folder, keyword, label, output_suffix):
    microbes_path = find_matching_image(
        folder=folder,
        treatment="Microbes",
        day=day,
        keyword=keyword,
    )

    no_microbes_path = find_matching_image(
        folder=folder,
        treatment="No Microbes",
        day=day,
        keyword=keyword,
    )

    if microbes_path is None or no_microbes_path is None:
        print(f"Skipping {label} comparison for Day {day}: missing image.")
        return None

    microbes_img = Image.open(microbes_path).convert("RGB")
    no_microbes_img = Image.open(no_microbes_path).convert("RGB")

    fig, axes = plt.subplots(1, 2, figsize=(16, 8))

    axes[0].imshow(microbes_img)
    axes[0].set_title(
        f"Microbes - Day {day}\n"
        f"{get_summary_text(summary_df, 'Microbes', day)}"
    )
    axes[0].axis("off")

    axes[1].imshow(no_microbes_img)
    axes[1].set_title(
        f"No Microbes - Day {day}\n"
        f"{get_summary_text(summary_df, 'No Microbes', day)}"
    )
    axes[1].axis("off")

    fig.suptitle(
        f"First Trial {label} Comparison - Day {day}",
        fontsize=16,
    )

    plt.tight_layout()

    output_path = VISUAL_DIR / f"first_trial_day_{day:02d}_{output_suffix}_microbes_vs_no_microbes.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_all_days_panel(image_df, summary_df):
    days = sorted(image_df["day"].dropna().astype(int).unique())

    fig, axes = plt.subplots(len(days), 2, figsize=(14, 5 * len(days)))

    if len(days) == 1:
        axes = [axes]

    for row_index, day in enumerate(days):
        microbes = image_df[
            (image_df["treatment"] == "Microbes") &
            (image_df["day"] == day)
        ]

        no_microbes = image_df[
            (image_df["treatment"] == "No Microbes") &
            (image_df["day"] == day)
        ]

        if microbes.empty or no_microbes.empty:
            continue

        microbes_img = Image.open(microbes.iloc[0]["path"]).convert("RGB")
        no_microbes_img = Image.open(no_microbes.iloc[0]["path"]).convert("RGB")

        axes[row_index][0].imshow(microbes_img)
        axes[row_index][0].set_title(
            f"Microbes - Day {day}\n"
            f"{get_summary_text(summary_df, 'Microbes', day)}"
        )
        axes[row_index][0].axis("off")

        axes[row_index][1].imshow(no_microbes_img)
        axes[row_index][1].set_title(
            f"No Microbes - Day {day}\n"
            f"{get_summary_text(summary_df, 'No Microbes', day)}"
        )
        axes[row_index][1].axis("off")

    fig.suptitle(
        "First Trial: Microbes vs No Microbes Standardised Visual Comparison",
        fontsize=18,
    )

    plt.tight_layout()

    output_path = VISUAL_DIR / "first_trial_all_days_standardized_microbes_vs_no_microbes_panel.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def save_report_table(df):
    report_df = df[
        [
            "treatment",
            "day",
            "total_cells",
            "estimated_germinated_cells",
            "estimated_germination_percent",
            "cumulative_germinated_cells",
            "cumulative_germination_percent",
            "count_decreased_before_correction",
        ]
    ].copy()

    output_path = OUTPUT_DIR / "first_trial_report_table_strict_standardized.csv"
    report_df.to_csv(output_path, index=False)

    return output_path


def main():
    summary_df = load_summary()
    image_df = load_standardized_images()

    chart_paths = [
        make_raw_germination_chart(summary_df),
        make_cumulative_germination_chart(summary_df),
        make_germinated_cell_count_chart(summary_df),
        make_final_comparison_chart(summary_df),
    ]

    visual_paths = []

    days = sorted(summary_df["day"].dropna().astype(int).unique())

    for day in days:
        path = make_day_image_comparison(image_df, summary_df, day)
        if path is not None:
            visual_paths.append(path)

        path = make_day_output_comparison(
            summary_df=summary_df,
            day=day,
            folder=STRICT_OVERLAY_DIR,
            keyword="overlay",
            label="Strict Overlay",
            output_suffix="strict_overlay",
        )
        if path is not None:
            visual_paths.append(path)

        path = make_day_output_comparison(
            summary_df=summary_df,
            day=day,
            folder=STRICT_MASK_DIR,
            keyword="green_mask",
            label="Strict Green Mask",
            output_suffix="strict_green_mask",
        )
        if path is not None:
            visual_paths.append(path)

    all_days_panel = make_all_days_panel(image_df, summary_df)
    visual_paths.append(all_days_panel)

    report_table_path = save_report_table(summary_df)

    print()
    print("Done.")
    print(f"Charts saved to:\n{CHART_DIR}")
    print(f"Visual comparisons saved to:\n{VISUAL_DIR}")
    print(f"Report table saved to:\n{report_table_path}")
    print()
    print("Charts:")
    for path in chart_paths:
        print(path)
    print()
    print("Visual comparisons:")
    for path in visual_paths:
        print(path)
    print()
    print(summary_df[
        [
            "filename",
            "treatment",
            "day",
            "total_cells",
            "estimated_germinated_cells",
            "estimated_germination_percent",
            "cumulative_germinated_cells",
            "cumulative_germination_percent",
        ]
    ])


if __name__ == "__main__":
    main()