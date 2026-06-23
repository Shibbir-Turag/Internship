from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np


# ============================================================
# First Trial - 04 Spatial Heatmaps
# Uses strict standardized germination detection results.
#
# Input:
# outputs/first_trial_germination/02_germination_detection_strict/
#   cell_measurements_standardized_strict.csv
#
# Output:
# outputs/first_trial_germination/04_spatial_heatmaps/
# ============================================================


ROOT_DIR = Path.cwd()

BASE_OUTPUT_DIR = ROOT_DIR / "outputs" / "first_trial_germination"

INPUT_CSV = (
    BASE_OUTPUT_DIR
    / "02_germination_detection_strict"
    / "cell_measurements_standardized_strict.csv"
)

OUTPUT_DIR = BASE_OUTPUT_DIR / "04_spatial_heatmaps"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

ROWS = 7
COLS = 10
TOTAL_CELLS = ROWS * COLS


def bool_to_int(value):
    if isinstance(value, bool):
        return int(value)

    value = str(value).strip().lower()

    if value in ["true", "1", "yes", "y"]:
        return 1

    return 0


def make_grid(group):
    grid = np.zeros((ROWS, COLS))

    for _, row in group.iterrows():
        r = int(row["row"]) - 1
        c = int(row["col"]) - 1
        grid[r, c] = bool_to_int(row["germinated_estimate"])

    return grid


def make_single_heatmap(group, treatment, day):
    grid = make_grid(group)

    germinated_count = int(grid.sum())
    percent = germinated_count / TOTAL_CELLS * 100

    plt.figure(figsize=(10, 7))
    plt.imshow(grid, vmin=0, vmax=1)

    for r in range(ROWS):
        for c in range(COLS):
            label = "G" if grid[r, c] == 1 else "-"
            plt.text(c, r, label, ha="center", va="center", fontsize=9)

    plt.title(
        f"First Trial Spatial Germination Heatmap\n"
        f"{treatment} - Day {day} | "
        f"{germinated_count}/{TOTAL_CELLS} cells ({percent:.1f}%)"
    )

    plt.xlabel("Column")
    plt.ylabel("Row")
    plt.xticks(range(COLS), range(1, COLS + 1))
    plt.yticks(range(ROWS), range(1, ROWS + 1))
    plt.colorbar(label="0 = not detected, 1 = germinated")
    plt.tight_layout()

    safe_treatment = treatment.lower().replace(" ", "_")
    output_path = OUTPUT_DIR / f"first_trial_{safe_treatment}_day_{int(day):02d}_spatial_heatmap.png"

    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_comparison_heatmap(df, day):
    day_df = df[df["day"].astype(int) == int(day)]

    microbes = day_df[day_df["treatment"] == "Microbes"]
    no_microbes = day_df[day_df["treatment"] == "No Microbes"]

    if microbes.empty or no_microbes.empty:
        print(f"Skipping Day {day}: missing one treatment.")
        return None

    fig, axes = plt.subplots(1, 2, figsize=(16, 7))

    for ax, group, treatment in [
        (axes[0], microbes, "Microbes"),
        (axes[1], no_microbes, "No Microbes"),
    ]:
        grid = make_grid(group)
        germinated_count = int(grid.sum())
        percent = germinated_count / TOTAL_CELLS * 100

        ax.imshow(grid, vmin=0, vmax=1)

        for r in range(ROWS):
            for c in range(COLS):
                label = "G" if grid[r, c] == 1 else "-"
                ax.text(c, r, label, ha="center", va="center", fontsize=8)

        ax.set_title(
            f"{treatment}\n"
            f"{germinated_count}/{TOTAL_CELLS} cells ({percent:.1f}%)"
        )

        ax.set_xlabel("Column")
        ax.set_ylabel("Row")
        ax.set_xticks(range(COLS))
        ax.set_xticklabels(range(1, COLS + 1))
        ax.set_yticks(range(ROWS))
        ax.set_yticklabels(range(1, ROWS + 1))

    fig.suptitle(f"First Trial Spatial Comparison - Day {day}", fontsize=16)
    plt.tight_layout()

    output_path = OUTPUT_DIR / f"first_trial_day_{int(day):02d}_microbes_vs_no_microbes_spatial_heatmap.png"
    plt.savefig(output_path, dpi=300)
    plt.close()

    return output_path


def make_row_column_summary(df):
    summary_rows = []

    for (treatment, day), group in df.groupby(["treatment", "day"]):
        for row_number, row_group in group.groupby("row"):
            germinated = row_group["germinated_estimate"].apply(bool_to_int).sum()
            total = len(row_group)

            summary_rows.append(
                {
                    "treatment": treatment,
                    "day": int(day),
                    "axis": "row",
                    "axis_number": int(row_number),
                    "germinated_cells": int(germinated),
                    "total_cells": total,
                    "germination_percent": round(germinated / total * 100, 2),
                }
            )

        for col_number, col_group in group.groupby("col"):
            germinated = col_group["germinated_estimate"].apply(bool_to_int).sum()
            total = len(col_group)

            summary_rows.append(
                {
                    "treatment": treatment,
                    "day": int(day),
                    "axis": "column",
                    "axis_number": int(col_number),
                    "germinated_cells": int(germinated),
                    "total_cells": total,
                    "germination_percent": round(germinated / total * 100, 2),
                }
            )

    summary_df = pd.DataFrame(summary_rows)
    output_path = OUTPUT_DIR / "first_trial_row_column_germination_summary.csv"
    summary_df.to_csv(output_path, index=False)

    return output_path


def main():
    if not INPUT_CSV.exists():
        raise FileNotFoundError(
            f"Could not find:\n{INPUT_CSV}\n\n"
            "Run first_trial_02_detect_germination_standardized.py first."
        )

    df = pd.read_csv(INPUT_CSV)

    required = [
        "filename",
        "treatment",
        "day",
        "row",
        "col",
        "germinated_estimate",
    ]

    missing = [col for col in required if col not in df.columns]

    if missing:
        raise ValueError(f"Missing columns: {missing}")

    output_paths = []

    for (treatment, day), group in df.groupby(["treatment", "day"]):
        output_paths.append(make_single_heatmap(group, treatment, int(day)))

    for day in sorted(df["day"].dropna().astype(int).unique()):
        path = make_comparison_heatmap(df, int(day))
        if path is not None:
            output_paths.append(path)

    row_col_summary_path = make_row_column_summary(df)

    print()
    print("Done.")
    print(f"Spatial heatmaps saved to:\n{OUTPUT_DIR}")
    print(f"Row/column summary saved to:\n{row_col_summary_path}")
    print()
    print("Created heatmaps:")

    for path in output_paths:
        print(path)


if __name__ == "__main__":
    main()