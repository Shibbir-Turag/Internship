from __future__ import annotations

"""
SCRIPT 05 — THIRD TRIAL RGB GROWTH-RATE AND TREATMENT COMPARISON

Purpose
-------
Use Script 04 visible-emergence and RGB green-cover outputs to calculate
Trial 3 tray-level and treatment-level growth metrics.

This script DOES:
- read Script 04 tray and cell results
- preserve observed Day 7 RGB green-cover values
- create adjusted/imputed Day 7 RGB green-cover values for likely bug-eaten cells
- calculate visible-emergence rate and RGB green-cover growth rate using real elapsed days
- compare Microbes vs No Microbes
- compare Ideal vs Heat vs Moisture
- compare Microbes x Treatment groups
- compare fixed Inside/Outside placement for Ideal and Moisture trays
- create a dedicated Ideal/Moisture Inside-vs-Outside comparison table
- summarise Heat and Moisture phase responses descriptively
- create CSV, Excel, chart, PDF, and Word report outputs

This script DOES NOT:
- alter Script 04 outputs
- alter the original images
- calculate NDVI or NDRE
- mix observed and adjusted Day 7 values silently

Important Day 7 rule
--------------------
Observed Day 7 values are kept exactly as Script 04 measured them.

If a cell had visible green evidence before Day 7 but is missing on Day 7,
the script creates a separate adjusted estimate:

    adjusted Day 7 = previous observed green cover + previous growth slope x elapsed days

The adjusted estimate is clearly flagged with:
- day7_imputed
- imputation_reason
- imputation_method
- previous_growth_rate_pp_per_day

Important Inside/Outside rule
-----------------------------
Inside vs Outside comparison is only valid for fixed-position Ideal and Moisture
trays. Heat trays are excluded from this direct comparison because their
environment changed during the trial.

New output added
----------------
A short Word report is generated at:

    outputs/Third trial/05_RGB_Growth_Rate_Treatment_Comparison/_reports/
        rgb_growth_treatment_short_report.docx
"""

import argparse
import json
import math
import re
from pathlib import Path

import matplotlib
matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


# ============================================================
# 1) PATHS
# ============================================================

PROJECT_ROOT = Path(
    r"C:\Users\tshib\OneDrive\Desktop\Internship"
)

SCRIPT04_REPORTS = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "04_Visible_Emergence"
    / "_reports"
)

TRAY_SUMMARY_CSV = (
    SCRIPT04_REPORTS
    / "visible_emergence_tray_summary.csv"
)

CELL_RESULTS_CSV = (
    SCRIPT04_REPORTS
    / "visible_emergence_cell_results.csv"
)

FIRST_EMERGENCE_CSV = (
    SCRIPT04_REPORTS
    / "first_emergence_summary.csv"
)

OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Third trial"
    / "05_RGB_Growth_Rate_Treatment_Comparison"
)

CHARTS_ROOT = OUTPUT_ROOT / "charts"
REPORTS_ROOT = OUTPUT_ROOT / "_reports"
CONFIG_ROOT = OUTPUT_ROOT / "_config"


# ============================================================
# 2) SETTINGS
# ============================================================

EXPECTED_CELLS = 70

EXPECTED_OBSERVATION_DAYS = [
    1,
    2,
    3,
    4,
    5,
    6,
    7,
]

DAY_LABELS = {
    1: "Day 1",
    2: "Day 2",
    3: "Day 3",
    4: "Day 4",
    5: "Day 5",
    6: "Day 6",
    7: "Day 7",
}

DATE_MAP = {
    1: "2026-06-29",
    2: "2026-06-30",
    3: "2026-07-01",
    4: "2026-07-02",
    5: "2026-07-03",
    6: "2026-07-04",
    7: "2026-07-07",
}

ELAPSED_DAY_MAP = {
    1: 0,
    2: 1,
    3: 2,
    4: 3,
    5: 4,
    6: 5,
    7: 8,
}

PREVIOUS_PHOTO_INTERVAL_MAP = {
    1: "",
    2: 1,
    3: 1,
    4: 1,
    5: 1,
    6: 1,
    7: 3,
}

MICROBE_ORDER = [
    "No Microbes",
    "Microbes",
]

TREATMENT_ORDER = [
    "Ideal",
    "Heat",
    "Moisture",
]

FIXED_ENVIRONMENT_TREATMENTS = [
    "Ideal",
    "Moisture",
]

PERFORMANCE_COMPONENTS_OBSERVED = [
    "final_tracked_emergence_percent",
    "final_observed_green_cover_percent",
    "observed_green_cover_rate_day1_to_day7_pp_per_day",
    "emergence_rate_day1_to_day7_pp_per_day",
]

PERFORMANCE_COMPONENTS_ADJUSTED = [
    "final_tracked_emergence_percent",
    "final_adjusted_green_cover_percent",
    "adjusted_green_cover_rate_day1_to_day7_pp_per_day",
    "emergence_rate_day1_to_day7_pp_per_day",
]


# ============================================================
# 3) OPTIONAL REPORT IMPORTS
# ============================================================

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        Image as PDFImage,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
        Table,
        TableStyle,
    )

    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False


try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    DOCX_AVAILABLE = True
except Exception:
    DOCX_AVAILABLE = False


# ============================================================
# 4) GENERAL HELPERS
# ============================================================

def natural_key(value: object):
    return [
        int(part) if part.isdigit() else part.casefold()
        for part in re.split(r"(\d+)", str(value))
    ]


def yes_no(value: bool) -> str:
    return "Yes" if value else "No"


def parse_yes(value: object) -> bool:
    return str(value).strip().casefold() in {
        "yes",
        "y",
        "true",
        "1",
        "p",
    }


def require_columns(
    dataframe: pd.DataFrame,
    required_columns: list[str],
    source_name: str,
) -> None:
    missing = [
        column
        for column in required_columns
        if column not in dataframe.columns
    ]

    if missing:
        raise ValueError(
            f"{source_name} is missing required column(s): "
            + ", ".join(missing)
        )


def safe_numeric(
    dataframe: pd.DataFrame,
    columns: list[str],
) -> pd.DataFrame:
    dataframe = dataframe.copy()

    for column in columns:
        if column in dataframe.columns:
            dataframe[column] = pd.to_numeric(
                dataframe[column],
                errors="coerce",
            )

    return dataframe


def safe_round_dataframe(
    dataframe: pd.DataFrame,
    decimals: int = 4,
) -> pd.DataFrame:
    output = dataframe.copy()

    numeric_columns = output.select_dtypes(
        include=["number"]
    ).columns

    output[numeric_columns] = output[numeric_columns].round(decimals)

    return output


def minmax_score(series: pd.Series) -> pd.Series:
    values = pd.to_numeric(
        series,
        errors="coerce",
    )

    valid = values.dropna()

    if valid.empty:
        return pd.Series(
            [math.nan] * len(values),
            index=values.index,
        )

    minimum = valid.min()
    maximum = valid.max()

    if maximum == minimum:
        return pd.Series(
            [
                50.0 if pd.notna(value) else math.nan
                for value in values
            ],
            index=values.index,
        )

    return (
        values
        - minimum
    ) / (
        maximum
        - minimum
    ) * 100.0


def inverse_minmax_score(series: pd.Series) -> pd.Series:
    values = pd.to_numeric(
        series,
        errors="coerce",
    )

    valid = values.dropna()

    if valid.empty:
        return pd.Series(
            [math.nan] * len(values),
            index=values.index,
        )

    minimum = valid.min()
    maximum = valid.max()

    if maximum == minimum:
        return pd.Series(
            [
                50.0 if pd.notna(value) else math.nan
                for value in values
            ],
            index=values.index,
        )

    return (
        maximum
        - values
    ) / (
        maximum
        - minimum
    ) * 100.0


def mean_or_nan(
    dataframe: pd.DataFrame,
    column: str,
):
    if dataframe.empty or column not in dataframe.columns:
        return math.nan

    return pd.to_numeric(
        dataframe[column],
        errors="coerce",
    ).mean()


def count_unique_or_zero(
    dataframe: pd.DataFrame,
    column: str,
) -> int:
    if dataframe.empty or column not in dataframe.columns:
        return 0

    return int(dataframe[column].nunique())


def format_number(
    value: object,
    decimals: int = 2,
) -> str:
    try:
        if pd.isna(value):
            return "N/A"

        return f"{float(value):.{decimals}f}"
    except Exception:
        return "N/A"


# ============================================================
# 5) LOAD SCRIPT 04 OUTPUTS
# ============================================================

def load_script04_outputs() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if not TRAY_SUMMARY_CSV.exists():
        raise FileNotFoundError(
            f"Missing Script 04 tray summary:\n{TRAY_SUMMARY_CSV}"
        )

    if not CELL_RESULTS_CSV.exists():
        raise FileNotFoundError(
            f"Missing Script 04 cell results:\n{CELL_RESULTS_CSV}"
        )

    tray = pd.read_csv(TRAY_SUMMARY_CSV)
    cell = pd.read_csv(CELL_RESULTS_CSV)

    if FIRST_EMERGENCE_CSV.exists():
        first = pd.read_csv(FIRST_EMERGENCE_CSV)
    else:
        first = pd.DataFrame()

    require_columns(
        tray,
        [
            "day_order",
            "day",
            "calendar_date",
            "days_since_day1",
            "days_since_previous_photo",
            "tray",
            "tray_no",
            "microbe_status",
            "treatment",
            "label_environment",
            "observed_environment",
            "tracked_emergence_percent",
            "mean_rgb_green_cover_percent",
            "status",
        ],
        "visible_emergence_tray_summary.csv",
    )

    require_columns(
        cell,
        [
            "day_order",
            "day",
            "calendar_date",
            "days_since_day1",
            "tray",
            "tray_no",
            "cell_id",
            "raw_current_green_evidence",
            "tracked_visible_emerged",
            "green_area_percent",
            "microbe_status",
            "treatment",
            "label_environment",
            "observed_environment",
        ],
        "visible_emergence_cell_results.csv",
    )

    tray = safe_numeric(
        tray,
        [
            "day_order",
            "days_since_planting",
            "days_since_day1",
            "days_since_previous_photo",
            "tray_no",
            "raw_green_cells",
            "raw_green_percent",
            "tracked_emerged_cells",
            "tracked_emergence_percent",
            "newly_emerged_today",
            "newly_emerged_percent",
            "carried_forward_cells",
            "mean_green_area_percent",
            "mean_rgb_green_cover_percent",
            "possible_day7_bug_eaten_cells",
        ],
    )

    cell = safe_numeric(
        cell,
        [
            "day_order",
            "days_since_planting",
            "days_since_day1",
            "days_since_previous_photo",
            "tray_no",
            "cell_id",
            "row",
            "column",
            "green_pixels",
            "largest_green_component",
            "green_area_percent",
            "rgb_green_cover_percent",
            "zone_pixels",
        ],
    )

    tray["status"] = tray["status"].astype(str).str.upper()

    pass_tray = tray.loc[
        tray["status"].eq("PASS")
    ].copy()

    pass_keys = pass_tray[
        [
            "day_order",
            "tray_no",
            "capture_id",
        ]
    ].drop_duplicates()

    cell = cell.merge(
        pass_keys,
        on=[
            "day_order",
            "tray_no",
            "capture_id",
        ],
        how="inner",
    )

    for frame in [
        pass_tray,
        cell,
    ]:
        frame["day_order"] = frame["day_order"].astype(int)

        frame["calendar_date"] = frame.apply(
            lambda row: (
                DATE_MAP.get(int(row["day_order"]), "")
                if pd.isna(row.get("calendar_date"))
                or str(row.get("calendar_date")).strip() == ""
                else str(row.get("calendar_date")).strip()
            ),
            axis=1,
        )

        frame["days_since_day1"] = pd.to_numeric(
            frame["days_since_day1"],
            errors="coerce",
        )

        frame["days_since_day1"] = frame.apply(
            lambda row: (
                ELAPSED_DAY_MAP.get(int(row["day_order"]), math.nan)
                if pd.isna(row["days_since_day1"])
                else row["days_since_day1"]
            ),
            axis=1,
        )

        frame["days_since_previous_photo"] = pd.to_numeric(
            frame.get("days_since_previous_photo", math.nan),
            errors="coerce",
        )

        frame["days_since_previous_photo"] = frame.apply(
            lambda row: (
                PREVIOUS_PHOTO_INTERVAL_MAP.get(int(row["day_order"]), math.nan)
                if pd.isna(row["days_since_previous_photo"])
                else row["days_since_previous_photo"]
            ),
            axis=1,
        )

        frame["microbe_status"] = frame["microbe_status"].astype(str).str.strip()
        frame["treatment"] = frame["treatment"].astype(str).str.strip()
        frame["label_environment"] = frame["label_environment"].astype(str).str.strip()
        frame["observed_environment"] = frame["observed_environment"].astype(str).str.strip()

        frame["environment_group"] = np.where(
            frame["treatment"].str.casefold().eq("heat"),
            "Dynamic Heat",
            frame["label_environment"],
        )

        frame["microbe_treatment"] = (
            frame["microbe_status"]
            + " | "
            + frame["treatment"]
        )

        frame["treatment_environment"] = (
            frame["treatment"]
            + " | "
            + frame["environment_group"]
        )

        frame["microbe_environment"] = (
            frame["microbe_status"]
            + " | "
            + frame["environment_group"]
        )

    return pass_tray, cell, first


# ============================================================
# 6) DAY 7 ADJUSTMENT / IMPUTATION
# ============================================================

def estimate_adjusted_day7_for_group(
    group: pd.DataFrame,
) -> pd.DataFrame:
    group = group.sort_values(
        "day_order"
    ).copy()

    group["observed_green_area_percent"] = group["green_area_percent"]
    group["adjusted_green_area_percent"] = group["green_area_percent"]
    group["day7_imputed"] = "No"
    group["imputation_reason"] = ""
    group["imputation_method"] = ""
    group["previous_growth_rate_pp_per_day"] = math.nan
    group["previous_observed_green_percent"] = math.nan
    group["previous_observed_day_order"] = math.nan
    group["previous_observed_days_since_day1"] = math.nan

    if 7 not in set(group["day_order"]):
        return group

    day7_index = group.index[
        group["day_order"].eq(7)
    ][0]

    day7_row = group.loc[day7_index]

    raw_day7 = parse_yes(
        day7_row.get("raw_current_green_evidence", "")
    )

    possible_flag = parse_yes(
        day7_row.get("possible_day7_bug_eaten", "")
    )

    prior_rows = group.loc[
        group["day_order"] < 7
    ].copy()

    prior_raw_positive = prior_rows.loc[
        prior_rows["raw_current_green_evidence"].apply(parse_yes)
    ].copy()

    had_prior_visible_crop = not prior_raw_positive.empty

    should_impute = (
        possible_flag
        or (
            had_prior_visible_crop
            and not raw_day7
        )
    )

    if not should_impute:
        return group

    prior_positive_green = prior_rows.loc[
        pd.to_numeric(
            prior_rows["green_area_percent"],
            errors="coerce",
        ).fillna(0) > 0
    ].copy()

    if prior_positive_green.empty:
        return group

    prior_positive_green = prior_positive_green.sort_values(
        "day_order"
    )

    last_row = prior_positive_green.iloc[-1]
    last_green = float(last_row["green_area_percent"])
    last_elapsed = float(last_row["days_since_day1"])
    last_day_order = int(last_row["day_order"])

    slope = 0.0

    if len(prior_positive_green) >= 2:
        previous_row = prior_positive_green.iloc[-2]
        previous_green = float(previous_row["green_area_percent"])
        previous_elapsed = float(previous_row["days_since_day1"])

        elapsed_difference = max(
            last_elapsed
            - previous_elapsed,
            1e-9,
        )

        raw_slope = (
            last_green
            - previous_green
        ) / elapsed_difference

        slope = max(
            0.0,
            raw_slope,
        )

    day7_elapsed = float(day7_row["days_since_day1"])

    elapsed_to_day7 = max(
        day7_elapsed
        - last_elapsed,
        0.0,
    )

    estimated = (
        last_green
        + slope * elapsed_to_day7
    )

    observed_day7 = float(
        day7_row["green_area_percent"]
    )

    adjusted = min(
        100.0,
        max(
            observed_day7,
            estimated,
            last_green,
        ),
    )

    group.loc[
        day7_index,
        "adjusted_green_area_percent",
    ] = adjusted

    group.loc[
        day7_index,
        "day7_imputed",
    ] = "Yes"

    group.loc[
        day7_index,
        "imputation_reason",
    ] = (
        "Day 7 cell had no current green evidence but showed visible green "
        "evidence earlier; treated as possible bug-eaten/missing crop."
    )

    group.loc[
        day7_index,
        "imputation_method",
    ] = (
        "Estimated from the latest prior observed green cover and the previous "
        "positive growth rate. Negative previous growth was not projected."
    )

    group.loc[
        day7_index,
        "previous_growth_rate_pp_per_day",
    ] = slope

    group.loc[
        day7_index,
        "previous_observed_green_percent",
    ] = last_green

    group.loc[
        day7_index,
        "previous_observed_day_order",
    ] = last_day_order

    group.loc[
        day7_index,
        "previous_observed_days_since_day1",
    ] = last_elapsed

    return group


def create_adjusted_cell_table(
    cell: pd.DataFrame,
) -> pd.DataFrame:
    adjusted_groups = []

    for (_tray_no, _cell_id), group in cell.groupby(
        [
            "tray_no",
            "cell_id",
        ],
        sort=True,
    ):
        adjusted_groups.append(
            estimate_adjusted_day7_for_group(group)
        )

    adjusted = pd.concat(
        adjusted_groups,
        ignore_index=True,
    )

    adjusted["adjusted_raw_current_green_evidence"] = np.where(
        adjusted["adjusted_green_area_percent"] > 0,
        "Yes",
        "No",
    )

    adjusted["adjustment_difference_pp"] = (
        adjusted["adjusted_green_area_percent"]
        - adjusted["observed_green_area_percent"]
    )

    return adjusted.sort_values(
        [
            "day_order",
            "tray_no",
            "cell_id",
        ]
    ).reset_index(drop=True)


# ============================================================
# 7) TRAY DAILY METRICS
# ============================================================

def create_tray_daily_metrics(
    adjusted_cell: pd.DataFrame,
) -> pd.DataFrame:
    adjusted_cell = adjusted_cell.copy()

    adjusted_cell["raw_green_bool"] = adjusted_cell[
        "raw_current_green_evidence"
    ].apply(parse_yes)

    adjusted_cell["tracked_bool"] = adjusted_cell[
        "tracked_visible_emerged"
    ].apply(parse_yes)

    if "newly_emerged_today" in adjusted_cell.columns:
        adjusted_cell["newly_emerged_bool"] = adjusted_cell[
            "newly_emerged_today"
        ].apply(parse_yes)
    else:
        adjusted_cell["newly_emerged_bool"] = False

    group_columns = [
        "day_order",
        "day",
        "calendar_date",
        "days_since_day1",
        "days_since_previous_photo",
        "tray",
        "tray_no",
        "microbe_status",
        "treatment",
        "label_environment",
        "observed_environment",
        "environment_group",
        "microbe_treatment",
        "treatment_environment",
        "microbe_environment",
        "heat_phase",
        "moisture_phase",
        "moisture_watered_today",
        "capture_id",
    ]

    group_columns = [
        column
        for column in group_columns
        if column in adjusted_cell.columns
    ]

    daily = (
        adjusted_cell.groupby(
            group_columns,
            as_index=False,
        )
        .agg(
            cells_processed=("cell_id", "count"),
            raw_green_cells=("raw_green_bool", "sum"),
            tracked_emerged_cells=("tracked_bool", "sum"),
            newly_emerged_cells=("newly_emerged_bool", "sum"),
            observed_green_cover_percent=("observed_green_area_percent", "mean"),
            adjusted_green_cover_percent=("adjusted_green_area_percent", "mean"),
            day7_imputed_cells=("day7_imputed", lambda s: int((s == "Yes").sum())),
            mean_adjustment_difference_pp=("adjustment_difference_pp", "mean"),
        )
    )

    invalid = daily.loc[
        daily["cells_processed"] != EXPECTED_CELLS
    ]

    if not invalid.empty:
        raise ValueError(
            "Each tray/day must have 70 cell records. Invalid records:\n"
            + invalid[
                [
                    "day_order",
                    "tray_no",
                    "cells_processed",
                ]
            ].to_string(index=False)
        )

    daily["raw_green_percent"] = (
        daily["raw_green_cells"]
        / EXPECTED_CELLS
        * 100.0
    )

    daily["tracked_emergence_percent"] = (
        daily["tracked_emerged_cells"]
        / EXPECTED_CELLS
        * 100.0
    )

    daily["newly_emerged_percent"] = (
        daily["newly_emerged_cells"]
        / EXPECTED_CELLS
        * 100.0
    )

    return daily.sort_values(
        [
            "day_order",
            "tray_no",
        ]
    ).reset_index(drop=True)


# ============================================================
# 8) TRAY GROWTH METRICS
# ============================================================

def value_for_day(
    rows_by_day: dict[int, pd.Series],
    day_order: int,
    column: str,
):
    row = rows_by_day.get(day_order)

    if row is None:
        return math.nan

    return row.get(
        column,
        math.nan,
    )


def elapsed_for_day(
    rows_by_day: dict[int, pd.Series],
    day_order: int,
):
    row = rows_by_day.get(day_order)

    if row is None:
        return math.nan

    return pd.to_numeric(
        row.get("days_since_day1", math.nan),
        errors="coerce",
    )


def rate_between_days(
    rows_by_day: dict[int, pd.Series],
    start_day: int,
    end_day: int,
    column: str,
):
    start_value = value_for_day(
        rows_by_day,
        start_day,
        column,
    )

    end_value = value_for_day(
        rows_by_day,
        end_day,
        column,
    )

    start_elapsed = elapsed_for_day(
        rows_by_day,
        start_day,
    )

    end_elapsed = elapsed_for_day(
        rows_by_day,
        end_day,
    )

    if (
        pd.isna(start_value)
        or pd.isna(end_value)
        or pd.isna(start_elapsed)
        or pd.isna(end_elapsed)
        or end_elapsed <= start_elapsed
    ):
        return math.nan

    return (
        end_value
        - start_value
    ) / (
        end_elapsed
        - start_elapsed
    )


def t50_elapsed_days(
    group: pd.DataFrame,
    emergence_column: str = "tracked_emergence_percent",
):
    series = group.sort_values(
        "days_since_day1"
    )[
        [
            "day_order",
            "days_since_day1",
            emergence_column,
        ]
    ].dropna()

    if series.empty:
        return math.nan

    if series[emergence_column].max() < 50:
        return math.nan

    previous_elapsed = None
    previous_value = None

    for row in series.itertuples(index=False):
        elapsed = float(row.days_since_day1)
        value = float(getattr(row, emergence_column))

        if value >= 50:
            if previous_elapsed is None:
                return elapsed

            if value == previous_value:
                return elapsed

            fraction = (
                50.0
                - previous_value
            ) / (
                value
                - previous_value
            )

            return previous_elapsed + fraction * (
                elapsed
                - previous_elapsed
            )

        previous_elapsed = elapsed
        previous_value = value

    return math.nan


def create_first_emergence_metrics(
    adjusted_cell: pd.DataFrame,
) -> pd.DataFrame:
    if "first_visible_emergence_day_order" not in adjusted_cell.columns:
        return pd.DataFrame()

    first = adjusted_cell[
        [
            "tray_no",
            "cell_id",
            "first_visible_emergence_day_order",
            "first_visible_emergence_date",
        ]
    ].drop_duplicates(
        subset=[
            "tray_no",
            "cell_id",
        ]
    ).copy()

    first["first_visible_emergence_day_order"] = pd.to_numeric(
        first["first_visible_emergence_day_order"],
        errors="coerce",
    )

    elapsed_lookup = dict(
        zip(
            adjusted_cell["day_order"].astype(int),
            adjusted_cell["days_since_day1"],
        )
    )

    first["first_visible_emergence_elapsed_days"] = first[
        "first_visible_emergence_day_order"
    ].map(
        elapsed_lookup
    )

    first = first.dropna(
        subset=[
            "first_visible_emergence_day_order",
        ]
    )

    if first.empty:
        return pd.DataFrame()

    return (
        first.groupby(
            "tray_no",
            as_index=False,
        )
        .agg(
            mean_first_visible_emergence_day_order=(
                "first_visible_emergence_day_order",
                "mean",
            ),
            median_first_visible_emergence_day_order=(
                "first_visible_emergence_day_order",
                "median",
            ),
            mean_first_visible_emergence_elapsed_days=(
                "first_visible_emergence_elapsed_days",
                "mean",
            ),
            median_first_visible_emergence_elapsed_days=(
                "first_visible_emergence_elapsed_days",
                "median",
            ),
        )
    )


def create_tray_growth_metrics(
    tray_daily: pd.DataFrame,
    adjusted_cell: pd.DataFrame,
) -> pd.DataFrame:
    records = []

    for tray_no, group in tray_daily.groupby("tray_no"):
        group = group.sort_values("day_order")

        rows_by_day = {
            int(row["day_order"]): row
            for _, row in group.iterrows()
        }

        reference = group.iloc[0]

        final_day = int(group["day_order"].max())
        final_row = rows_by_day[final_day]

        record = {
            "tray_no": int(tray_no),
            "tray": str(reference["tray"]),
            "microbe_status": str(reference["microbe_status"]),
            "treatment": str(reference["treatment"]),
            "label_environment": str(reference["label_environment"]),
            "environment_group": str(reference["environment_group"]),
            "microbe_treatment": str(reference["microbe_treatment"]),
            "treatment_environment": str(reference["treatment_environment"]),
            "microbe_environment": str(reference["microbe_environment"]),
            "available_day_count": int(group["day_order"].nunique()),
            "final_day_order": final_day,
            "final_calendar_date": str(final_row["calendar_date"]),
            "final_days_since_day1": float(final_row["days_since_day1"]),

            "day1_tracked_emergence_percent": value_for_day(rows_by_day, 1, "tracked_emergence_percent"),
            "final_tracked_emergence_percent": value_for_day(rows_by_day, 7, "tracked_emergence_percent"),

            "day1_observed_green_cover_percent": value_for_day(rows_by_day, 1, "observed_green_cover_percent"),
            "final_observed_green_cover_percent": value_for_day(rows_by_day, 7, "observed_green_cover_percent"),

            "day1_adjusted_green_cover_percent": value_for_day(rows_by_day, 1, "adjusted_green_cover_percent"),
            "final_adjusted_green_cover_percent": value_for_day(rows_by_day, 7, "adjusted_green_cover_percent"),

            "day7_imputed_cells": value_for_day(rows_by_day, 7, "day7_imputed_cells"),
            "day7_mean_adjustment_difference_pp": value_for_day(rows_by_day, 7, "mean_adjustment_difference_pp"),

            "emergence_rate_day1_to_day7_pp_per_day": rate_between_days(rows_by_day, 1, 7, "tracked_emergence_percent"),
            "observed_green_cover_rate_day1_to_day7_pp_per_day": rate_between_days(rows_by_day, 1, 7, "observed_green_cover_percent"),
            "adjusted_green_cover_rate_day1_to_day7_pp_per_day": rate_between_days(rows_by_day, 1, 7, "adjusted_green_cover_percent"),

            "observed_green_cover_rate_day1_to_day6_pp_per_day": rate_between_days(rows_by_day, 1, 6, "observed_green_cover_percent"),
            "adjusted_green_cover_rate_day1_to_day6_pp_per_day": rate_between_days(rows_by_day, 1, 6, "adjusted_green_cover_percent"),

            "observed_green_cover_rate_day6_to_day7_pp_per_day": rate_between_days(rows_by_day, 6, 7, "observed_green_cover_percent"),
            "adjusted_green_cover_rate_day6_to_day7_pp_per_day": rate_between_days(rows_by_day, 6, 7, "adjusted_green_cover_percent"),

            "observed_day7_change_from_day6_pp": (
                value_for_day(rows_by_day, 7, "observed_green_cover_percent")
                - value_for_day(rows_by_day, 6, "observed_green_cover_percent")
            ),
            "adjusted_day7_change_from_day6_pp": (
                value_for_day(rows_by_day, 7, "adjusted_green_cover_percent")
                - value_for_day(rows_by_day, 6, "adjusted_green_cover_percent")
            ),

            "t50_elapsed_days": t50_elapsed_days(group),
        }

        records.append(record)

    metrics = pd.DataFrame(records)

    first_metrics = create_first_emergence_metrics(
        adjusted_cell
    )

    if not first_metrics.empty:
        metrics = metrics.merge(
            first_metrics,
            on="tray_no",
            how="left",
            validate="one_to_one",
        )

    for component in PERFORMANCE_COMPONENTS_OBSERVED:
        metrics[f"{component}_observed_score"] = (
            minmax_score(metrics[component])
            if component in metrics.columns
            else math.nan
        )

    observed_score_columns = [
        f"{component}_observed_score"
        for component in PERFORMANCE_COMPONENTS_OBSERVED
    ]

    if "t50_elapsed_days" in metrics.columns:
        metrics["t50_observed_score"] = inverse_minmax_score(
            metrics["t50_elapsed_days"]
        )
        observed_score_columns.append("t50_observed_score")

    metrics["overall_observed_rgb_score"] = metrics[
        observed_score_columns
    ].mean(
        axis=1,
        skipna=True,
    )

    metrics["overall_observed_rgb_rank"] = metrics[
        "overall_observed_rgb_score"
    ].rank(
        ascending=False,
        method="min",
    ).astype("Int64")

    for component in PERFORMANCE_COMPONENTS_ADJUSTED:
        metrics[f"{component}_adjusted_score"] = (
            minmax_score(metrics[component])
            if component in metrics.columns
            else math.nan
        )

    adjusted_score_columns = [
        f"{component}_adjusted_score"
        for component in PERFORMANCE_COMPONENTS_ADJUSTED
    ]

    if "t50_elapsed_days" in metrics.columns:
        metrics["t50_adjusted_score"] = inverse_minmax_score(
            metrics["t50_elapsed_days"]
        )
        adjusted_score_columns.append("t50_adjusted_score")

    metrics["overall_adjusted_rgb_score"] = metrics[
        adjusted_score_columns
    ].mean(
        axis=1,
        skipna=True,
    )

    metrics["overall_adjusted_rgb_rank"] = metrics[
        "overall_adjusted_rgb_score"
    ].rank(
        ascending=False,
        method="min",
    ).astype("Int64")

    return metrics.sort_values(
        [
            "overall_adjusted_rgb_rank",
            "tray_no",
        ],
        na_position="last",
    ).reset_index(drop=True)


# ============================================================
# 9) GROUP TABLES
# ============================================================

def create_group_daily(
    tray_daily: pd.DataFrame,
    group_column: str,
    group_type: str,
) -> pd.DataFrame:
    result = (
        tray_daily.groupby(
            [
                group_column,
                "day_order",
                "day",
                "calendar_date",
                "days_since_day1",
            ],
            as_index=False,
        )
        .agg(
            tray_count=("tray_no", "nunique"),
            mean_tracked_emergence_percent=("tracked_emergence_percent", "mean"),
            sd_tracked_emergence_percent=("tracked_emergence_percent", "std"),
            mean_observed_green_cover_percent=("observed_green_cover_percent", "mean"),
            sd_observed_green_cover_percent=("observed_green_cover_percent", "std"),
            mean_adjusted_green_cover_percent=("adjusted_green_cover_percent", "mean"),
            sd_adjusted_green_cover_percent=("adjusted_green_cover_percent", "std"),
            mean_newly_emerged_cells=("newly_emerged_cells", "mean"),
            mean_day7_imputed_cells=("day7_imputed_cells", "mean"),
        )
        .rename(
            columns={
                group_column: "group",
            }
        )
    )

    result["group_type"] = group_type

    for column in [
        "sd_tracked_emergence_percent",
        "sd_observed_green_cover_percent",
        "sd_adjusted_green_cover_percent",
    ]:
        result[column] = result[column].fillna(0.0)

    return result


def create_group_growth(
    tray_metrics: pd.DataFrame,
    group_column: str,
    group_type: str,
) -> pd.DataFrame:
    metric_columns = [
        "final_tracked_emergence_percent",
        "final_observed_green_cover_percent",
        "final_adjusted_green_cover_percent",
        "day7_imputed_cells",
        "day7_mean_adjustment_difference_pp",
        "emergence_rate_day1_to_day7_pp_per_day",
        "observed_green_cover_rate_day1_to_day7_pp_per_day",
        "adjusted_green_cover_rate_day1_to_day7_pp_per_day",
        "observed_green_cover_rate_day6_to_day7_pp_per_day",
        "adjusted_green_cover_rate_day6_to_day7_pp_per_day",
        "t50_elapsed_days",
        "mean_first_visible_emergence_elapsed_days",
        "overall_observed_rgb_score",
        "overall_adjusted_rgb_score",
    ]

    aggregations = {
        "tray_count": (
            "tray_no",
            "nunique",
        )
    }

    for metric in metric_columns:
        if metric in tray_metrics.columns:
            aggregations[f"mean_{metric}"] = (
                metric,
                "mean",
            )
            aggregations[f"sd_{metric}"] = (
                metric,
                "std",
            )

    output = (
        tray_metrics.groupby(
            group_column,
            as_index=False,
        )
        .agg(**aggregations)
        .rename(
            columns={
                group_column: "group",
            }
        )
    )

    output["group_type"] = group_type

    for column in output.columns:
        if column.startswith("sd_"):
            output[column] = output[column].fillna(0.0)

    return output


def create_all_group_tables(
    tray_daily: pd.DataFrame,
    tray_metrics: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    group_specs = [
        ("microbe_status", "Microbe Status"),
        ("treatment", "Treatment Type"),
        ("microbe_treatment", "Microbe x Treatment"),
        ("treatment_environment", "Treatment x Environment"),
        ("microbe_environment", "Microbe x Environment"),
    ]

    daily_parts = []
    growth_parts = []

    for group_column, group_type in group_specs:
        daily_parts.append(
            create_group_daily(
                tray_daily,
                group_column,
                group_type,
            )
        )

        growth_parts.append(
            create_group_growth(
                tray_metrics,
                group_column,
                group_type,
            )
        )

    group_daily = pd.concat(
        daily_parts,
        ignore_index=True,
    )

    group_growth = pd.concat(
        growth_parts,
        ignore_index=True,
    )

    return group_daily, group_growth


# ============================================================
# 10) DEDICATED IDEAL/MOISTURE INSIDE VS OUTSIDE COMPARISON
# ============================================================

def create_inside_outside_comparison(
    tray_metrics: pd.DataFrame,
) -> pd.DataFrame:
    rows = []

    for treatment in FIXED_ENVIRONMENT_TREATMENTS:
        subset = tray_metrics.loc[
            tray_metrics["treatment"].astype(str).str.casefold().eq(
                treatment.casefold()
            )
        ].copy()

        inside = subset.loc[
            subset["label_environment"].astype(str).str.casefold().eq("inside")
        ].copy()

        outside = subset.loc[
            subset["label_environment"].astype(str).str.casefold().eq("outside")
        ].copy()

        row = {
            "treatment": treatment,

            "inside_tray_count": count_unique_or_zero(inside, "tray_no"),
            "outside_tray_count": count_unique_or_zero(outside, "tray_no"),

            "inside_trays": ", ".join(inside["tray"].astype(str).tolist()),
            "outside_trays": ", ".join(outside["tray"].astype(str).tolist()),

            "inside_final_emergence_percent": mean_or_nan(
                inside,
                "final_tracked_emergence_percent",
            ),
            "outside_final_emergence_percent": mean_or_nan(
                outside,
                "final_tracked_emergence_percent",
            ),

            "inside_final_observed_green_cover_percent": mean_or_nan(
                inside,
                "final_observed_green_cover_percent",
            ),
            "outside_final_observed_green_cover_percent": mean_or_nan(
                outside,
                "final_observed_green_cover_percent",
            ),

            "inside_final_adjusted_green_cover_percent": mean_or_nan(
                inside,
                "final_adjusted_green_cover_percent",
            ),
            "outside_final_adjusted_green_cover_percent": mean_or_nan(
                outside,
                "final_adjusted_green_cover_percent",
            ),

            "inside_emergence_rate_pp_per_day": mean_or_nan(
                inside,
                "emergence_rate_day1_to_day7_pp_per_day",
            ),
            "outside_emergence_rate_pp_per_day": mean_or_nan(
                outside,
                "emergence_rate_day1_to_day7_pp_per_day",
            ),

            "inside_adjusted_green_cover_rate_pp_per_day": mean_or_nan(
                inside,
                "adjusted_green_cover_rate_day1_to_day7_pp_per_day",
            ),
            "outside_adjusted_green_cover_rate_pp_per_day": mean_or_nan(
                outside,
                "adjusted_green_cover_rate_day1_to_day7_pp_per_day",
            ),

            "inside_t50_elapsed_days": mean_or_nan(
                inside,
                "t50_elapsed_days",
            ),
            "outside_t50_elapsed_days": mean_or_nan(
                outside,
                "t50_elapsed_days",
            ),

            "inside_overall_adjusted_rgb_score": mean_or_nan(
                inside,
                "overall_adjusted_rgb_score",
            ),
            "outside_overall_adjusted_rgb_score": mean_or_nan(
                outside,
                "overall_adjusted_rgb_score",
            ),
        }

        row["inside_minus_outside_final_emergence_pp"] = (
            row["inside_final_emergence_percent"]
            - row["outside_final_emergence_percent"]
        )

        row["inside_minus_outside_adjusted_green_cover_pp"] = (
            row["inside_final_adjusted_green_cover_percent"]
            - row["outside_final_adjusted_green_cover_percent"]
        )

        row["inside_minus_outside_adjusted_growth_rate_pp_per_day"] = (
            row["inside_adjusted_green_cover_rate_pp_per_day"]
            - row["outside_adjusted_green_cover_rate_pp_per_day"]
        )

        row["inside_minus_outside_overall_score"] = (
            row["inside_overall_adjusted_rgb_score"]
            - row["outside_overall_adjusted_rgb_score"]
        )

        if pd.isna(row["inside_minus_outside_adjusted_green_cover_pp"]):
            interpretation = "Insufficient data for comparison."
        elif row["inside_minus_outside_adjusted_green_cover_pp"] > 0:
            interpretation = (
                "Inside performed higher than Outside for adjusted Day 7 RGB green-cover."
            )
        elif row["inside_minus_outside_adjusted_green_cover_pp"] < 0:
            interpretation = (
                "Outside performed higher than Inside for adjusted Day 7 RGB green-cover."
            )
        else:
            interpretation = (
                "Inside and Outside were equal for adjusted Day 7 RGB green-cover."
            )

        row["interpretation"] = interpretation

        rows.append(row)

    return pd.DataFrame(rows)


# ============================================================
# 11) PHASE TABLES
# ============================================================

def create_phase_tables(
    tray_daily: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    heat = tray_daily.loc[
        tray_daily["treatment"].astype(str).str.casefold().eq("heat")
    ].copy()

    moisture = tray_daily.loc[
        tray_daily["treatment"].astype(str).str.casefold().eq("moisture")
    ].copy()

    if not heat.empty and "heat_phase" in heat.columns:
        heat_phase = (
            heat.groupby(
                [
                    "heat_phase",
                    "microbe_status",
                    "day_order",
                    "day",
                    "days_since_day1",
                ],
                as_index=False,
            )
            .agg(
                tray_count=("tray_no", "nunique"),
                mean_tracked_emergence_percent=("tracked_emergence_percent", "mean"),
                mean_observed_green_cover_percent=("observed_green_cover_percent", "mean"),
                mean_adjusted_green_cover_percent=("adjusted_green_cover_percent", "mean"),
                mean_day7_imputed_cells=("day7_imputed_cells", "mean"),
            )
            .sort_values(
                [
                    "day_order",
                    "microbe_status",
                ]
            )
            .reset_index(drop=True)
        )
    else:
        heat_phase = pd.DataFrame()

    if not moisture.empty and "moisture_phase" in moisture.columns:
        moisture_phase = (
            moisture.groupby(
                [
                    "moisture_phase",
                    "moisture_watered_today",
                    "microbe_status",
                    "label_environment",
                    "day_order",
                    "day",
                    "days_since_day1",
                ],
                as_index=False,
            )
            .agg(
                tray_count=("tray_no", "nunique"),
                mean_tracked_emergence_percent=("tracked_emergence_percent", "mean"),
                mean_observed_green_cover_percent=("observed_green_cover_percent", "mean"),
                mean_adjusted_green_cover_percent=("adjusted_green_cover_percent", "mean"),
                mean_day7_imputed_cells=("day7_imputed_cells", "mean"),
            )
            .sort_values(
                [
                    "day_order",
                    "microbe_status",
                    "label_environment",
                ]
            )
            .reset_index(drop=True)
        )
    else:
        moisture_phase = pd.DataFrame()

    return heat_phase, moisture_phase


# ============================================================
# 12) CHARTS
# ============================================================

def save_trend_chart(
    daily: pd.DataFrame,
    group_type: str,
    groups: list[str],
    value_column: str,
    sd_column: str,
    title: str,
    y_label: str,
    output_path: Path,
) -> None:
    subset = daily.loc[
        daily["group_type"].eq(group_type)
    ].copy()

    if subset.empty:
        return

    figure, axis = plt.subplots(
        figsize=(11.5, 6.8)
    )

    for group in groups:
        series = subset.loc[
            subset["group"].eq(group)
        ].sort_values("days_since_day1")

        if series.empty:
            continue

        axis.errorbar(
            series["days_since_day1"],
            series[value_column],
            yerr=series[sd_column] if sd_column in series.columns else None,
            marker="o",
            linewidth=2,
            capsize=4,
            label=f"{group} (n={int(series['tray_count'].max())})",
        )

    ticks = (
        subset[
            [
                "day",
                "days_since_day1",
            ]
        ]
        .drop_duplicates()
        .sort_values("days_since_day1")
    )

    axis.set_title(title)
    axis.set_xlabel("Elapsed days since Day 1 image")
    axis.set_ylabel(y_label)
    axis.set_xticks(ticks["days_since_day1"])
    axis.set_xticklabels(ticks["day"])
    axis.grid(True, axis="y", alpha=0.30)
    axis.legend(loc="best")

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def save_day7_observed_adjusted_chart(
    tray_metrics: pd.DataFrame,
    output_path: Path,
) -> None:
    frame = tray_metrics.sort_values("tray_no").copy()

    x = np.arange(len(frame))
    width = 0.38

    figure, axis = plt.subplots(
        figsize=(12, 6.5)
    )

    axis.bar(
        x - width / 2,
        frame["final_observed_green_cover_percent"],
        width,
        label="Observed Day 7",
    )

    axis.bar(
        x + width / 2,
        frame["final_adjusted_green_cover_percent"],
        width,
        label="Adjusted Day 7",
    )

    axis.set_title("Day 7 RGB green-cover: observed vs adjusted")
    axis.set_xlabel("Tray")
    axis.set_ylabel("Mean RGB green-cover (%)")
    axis.set_xticks(x)
    axis.set_xticklabels(frame["tray"], rotation=30, ha="right")
    axis.grid(True, axis="y", alpha=0.30)
    axis.legend(loc="best")

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def save_tray_ranking_chart(
    tray_metrics: pd.DataFrame,
    output_path: Path,
) -> None:
    frame = tray_metrics.dropna(
        subset=["overall_adjusted_rgb_score"]
    ).sort_values(
        "overall_adjusted_rgb_score"
    ).copy()

    if frame.empty:
        return

    frame["label"] = (
        frame["tray"].astype(str)
        + " — "
        + frame["microbe_status"].astype(str)
        + " | "
        + frame["treatment"].astype(str)
        + " | "
        + frame["environment_group"].astype(str)
    )

    figure, axis = plt.subplots(
        figsize=(12.5, 7.2)
    )

    axis.barh(
        frame["label"],
        frame["overall_adjusted_rgb_score"],
    )

    axis.set_title("Trial 3 tray ranking by adjusted RGB performance score")
    axis.set_xlabel("Adjusted RGB performance score")
    axis.set_ylabel("Tray")
    axis.grid(True, axis="x", alpha=0.30)

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def save_bug_imputed_cells_chart(
    tray_metrics: pd.DataFrame,
    output_path: Path,
) -> None:
    frame = tray_metrics.sort_values("tray_no").copy()

    if "day7_imputed_cells" not in frame.columns:
        return

    figure, axis = plt.subplots(
        figsize=(11.5, 6.2)
    )

    axis.bar(
        frame["tray"],
        frame["day7_imputed_cells"].fillna(0),
    )

    axis.set_title("Possible Day 7 bug-eaten/missing cells flagged by tray")
    axis.set_xlabel("Tray")
    axis.set_ylabel("Cells flagged")
    axis.grid(True, axis="y", alpha=0.30)
    axis.tick_params(axis="x", rotation=30)

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def save_inside_outside_green_chart(
    comparison: pd.DataFrame,
    output_path: Path,
) -> None:
    if comparison.empty:
        return

    x = np.arange(len(comparison))
    width = 0.36

    figure, axis = plt.subplots(
        figsize=(9.5, 5.8)
    )

    axis.bar(
        x - width / 2,
        comparison["inside_final_adjusted_green_cover_percent"],
        width,
        label="Inside",
    )

    axis.bar(
        x + width / 2,
        comparison["outside_final_adjusted_green_cover_percent"],
        width,
        label="Outside",
    )

    axis.set_title("Ideal and Moisture: Inside vs Outside adjusted Day 7 RGB green-cover")
    axis.set_xlabel("Treatment")
    axis.set_ylabel("Adjusted Day 7 RGB green-cover (%)")
    axis.set_xticks(x)
    axis.set_xticklabels(comparison["treatment"])
    axis.grid(True, axis="y", alpha=0.30)
    axis.legend(loc="best")

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def save_inside_outside_emergence_chart(
    comparison: pd.DataFrame,
    output_path: Path,
) -> None:
    if comparison.empty:
        return

    x = np.arange(len(comparison))
    width = 0.36

    figure, axis = plt.subplots(
        figsize=(9.5, 5.8)
    )

    axis.bar(
        x - width / 2,
        comparison["inside_final_emergence_percent"],
        width,
        label="Inside",
    )

    axis.bar(
        x + width / 2,
        comparison["outside_final_emergence_percent"],
        width,
        label="Outside",
    )

    axis.set_title("Ideal and Moisture: Inside vs Outside final visible emergence")
    axis.set_xlabel("Treatment")
    axis.set_ylabel("Final tracked visible emergence (%)")
    axis.set_xticks(x)
    axis.set_xticklabels(comparison["treatment"])
    axis.grid(True, axis="y", alpha=0.30)
    axis.legend(loc="best")

    figure.tight_layout()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    figure.savefig(output_path, dpi=220)
    plt.close(figure)


def create_charts(
    group_daily: pd.DataFrame,
    tray_metrics: pd.DataFrame,
    inside_outside_comparison: pd.DataFrame,
) -> dict[str, Path]:
    CHARTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    charts = {}

    path = CHARTS_ROOT / "01_microbes_visible_emergence_trend.png"
    save_trend_chart(
        group_daily,
        "Microbe Status",
        MICROBE_ORDER,
        "mean_tracked_emergence_percent",
        "sd_tracked_emergence_percent",
        "Tracked visible emergence: Microbes vs No Microbes",
        "Mean tracked visible emergence (%)",
        path,
    )
    charts["microbes_emergence"] = path

    path = CHARTS_ROOT / "02_microbes_adjusted_green_cover_trend.png"
    save_trend_chart(
        group_daily,
        "Microbe Status",
        MICROBE_ORDER,
        "mean_adjusted_green_cover_percent",
        "sd_adjusted_green_cover_percent",
        "Adjusted RGB green-cover: Microbes vs No Microbes",
        "Mean adjusted RGB green-cover (%)",
        path,
    )
    charts["microbes_green"] = path

    path = CHARTS_ROOT / "03_treatment_visible_emergence_trend.png"
    save_trend_chart(
        group_daily,
        "Treatment Type",
        TREATMENT_ORDER,
        "mean_tracked_emergence_percent",
        "sd_tracked_emergence_percent",
        "Tracked visible emergence by treatment type",
        "Mean tracked visible emergence (%)",
        path,
    )
    charts["treatment_emergence"] = path

    path = CHARTS_ROOT / "04_treatment_adjusted_green_cover_trend.png"
    save_trend_chart(
        group_daily,
        "Treatment Type",
        TREATMENT_ORDER,
        "mean_adjusted_green_cover_percent",
        "sd_adjusted_green_cover_percent",
        "Adjusted RGB green-cover by treatment type",
        "Mean adjusted RGB green-cover (%)",
        path,
    )
    charts["treatment_green"] = path

    path = CHARTS_ROOT / "05_day7_observed_vs_adjusted_green_cover_by_tray.png"
    save_day7_observed_adjusted_chart(
        tray_metrics,
        path,
    )
    charts["day7_observed_adjusted"] = path

    path = CHARTS_ROOT / "06_adjusted_rgb_tray_ranking.png"
    save_tray_ranking_chart(
        tray_metrics,
        path,
    )
    charts["tray_ranking"] = path

    path = CHARTS_ROOT / "07_day7_possible_bug_eaten_cells_by_tray.png"
    save_bug_imputed_cells_chart(
        tray_metrics,
        path,
    )
    charts["bug_cells"] = path

    path = CHARTS_ROOT / "08_inside_outside_adjusted_green_cover_ideal_moisture.png"
    save_inside_outside_green_chart(
        inside_outside_comparison,
        path,
    )
    charts["inside_outside_green"] = path

    path = CHARTS_ROOT / "09_inside_outside_emergence_ideal_moisture.png"
    save_inside_outside_emergence_chart(
        inside_outside_comparison,
        path,
    )
    charts["inside_outside_emergence"] = path

    return charts


# ============================================================
# 13) EXCEL AND CSV OUTPUTS
# ============================================================

def style_workbook(path: Path) -> None:
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


def save_tables(
    adjusted_cell: pd.DataFrame,
    tray_daily: pd.DataFrame,
    tray_metrics: pd.DataFrame,
    group_daily: pd.DataFrame,
    group_growth: pd.DataFrame,
    heat_phase: pd.DataFrame,
    moisture_phase: pd.DataFrame,
    inside_outside_comparison: pd.DataFrame,
) -> dict[str, Path]:
    REPORTS_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    paths = {
        "adjusted_cell": REPORTS_ROOT / "cell_growth_with_day7_adjustment.csv",
        "tray_daily": REPORTS_ROOT / "tray_daily_rgb_metrics.csv",
        "tray_metrics": REPORTS_ROOT / "tray_growth_metrics.csv",
        "group_daily": REPORTS_ROOT / "group_daily_metrics.csv",
        "group_growth": REPORTS_ROOT / "group_growth_metrics.csv",
        "heat_phase": REPORTS_ROOT / "heat_phase_response.csv",
        "moisture_phase": REPORTS_ROOT / "moisture_phase_response.csv",
        "inside_outside": REPORTS_ROOT / "inside_outside_comparison_ideal_moisture.csv",
        "day7_imputed": REPORTS_ROOT / "possible_day7_bug_eaten_cells.csv",
        "excel": REPORTS_ROOT / "rgb_growth_treatment_report.xlsx",
    }

    adjusted_cell.to_csv(paths["adjusted_cell"], index=False)
    tray_daily.to_csv(paths["tray_daily"], index=False)
    tray_metrics.to_csv(paths["tray_metrics"], index=False)
    group_daily.to_csv(paths["group_daily"], index=False)
    group_growth.to_csv(paths["group_growth"], index=False)
    heat_phase.to_csv(paths["heat_phase"], index=False)
    moisture_phase.to_csv(paths["moisture_phase"], index=False)
    inside_outside_comparison.to_csv(paths["inside_outside"], index=False)

    day7_imputed = adjusted_cell.loc[
        adjusted_cell["day7_imputed"].eq("Yes")
    ].copy()

    day7_imputed.to_csv(paths["day7_imputed"], index=False)

    readme = pd.DataFrame(
        {
            "Notes": [
                "This workbook summarises Trial 3 RGB visible-emergence and green-cover growth metrics.",
                "Observed Day 7 values come directly from Script 04 and are not overwritten.",
                "Adjusted Day 7 values are created only for cells that were visible before Day 7 but missing on Day 7.",
                "day7_imputed = Yes identifies adjusted records.",
                "Growth rates use real elapsed days since the Day 1 image, not folder number.",
                "Day 7 is 8 elapsed days after Day 1 because images were skipped on 05/07/2026 and 06/07/2026.",
                "The Inside vs Outside sheet compares only Ideal and Moisture trays because Heat trays had dynamic movement.",
                "These are descriptive image-derived results, not formal statistical proof.",
                "NDVI/NDRE are not included here; they are handled by later multispectral scripts.",
            ]
        }
    )

    with pd.ExcelWriter(
        paths["excel"],
        engine="openpyxl",
    ) as writer:
        safe_round_dataframe(tray_metrics).to_excel(
            writer,
            sheet_name="Tray Growth Metrics",
            index=False,
        )

        safe_round_dataframe(tray_daily).to_excel(
            writer,
            sheet_name="Tray Daily Metrics",
            index=False,
        )

        safe_round_dataframe(group_daily).to_excel(
            writer,
            sheet_name="Group Daily Metrics",
            index=False,
        )

        safe_round_dataframe(group_growth).to_excel(
            writer,
            sheet_name="Group Growth Metrics",
            index=False,
        )

        safe_round_dataframe(inside_outside_comparison).to_excel(
            writer,
            sheet_name="Inside Outside Compare",
            index=False,
        )

        safe_round_dataframe(day7_imputed).to_excel(
            writer,
            sheet_name="Day7 Imputed Cells",
            index=False,
        )

        safe_round_dataframe(heat_phase).to_excel(
            writer,
            sheet_name="Heat Phase Response",
            index=False,
        )

        safe_round_dataframe(moisture_phase).to_excel(
            writer,
            sheet_name="Moisture Phase Response",
            index=False,
        )

        readme.to_excel(
            writer,
            sheet_name="Read Me",
            index=False,
        )

    style_workbook(paths["excel"])

    return paths


# ============================================================
# 14) WORD REPORT
# ============================================================

def add_docx_table(
    document,
    dataframe: pd.DataFrame,
    columns: list[str],
    max_rows: int = 12,
):
    frame = dataframe[columns].head(max_rows).copy()
    frame = safe_round_dataframe(frame, 3)

    table = document.add_table(
        rows=1,
        cols=len(columns),
    )

    table.style = "Table Grid"

    header_cells = table.rows[0].cells

    for index, column in enumerate(columns):
        header_cells[index].text = column.replace("_", " ").title()

    for _, row in frame.iterrows():
        cells = table.add_row().cells

        for index, column in enumerate(columns):
            cells[index].text = str(row[column])


def add_picture_if_exists(
    document,
    path: Path | None,
    width_inches: float = 6.3,
):
    if path is None:
        return

    if not Path(path).exists():
        return

    document.add_picture(
        str(path),
        width=Inches(width_inches),
    )


def best_row_by_metric(
    dataframe: pd.DataFrame,
    metric: str,
):
    if dataframe.empty or metric not in dataframe.columns:
        return None

    frame = dataframe.dropna(
        subset=[metric]
    )

    if frame.empty:
        return None

    return frame.loc[
        frame[metric].idxmax()
    ]


def group_summary_sentence(
    group_growth: pd.DataFrame,
    group_type: str,
    metric: str,
    metric_label: str,
) -> str:
    subset = group_growth.loc[
        group_growth["group_type"].eq(group_type)
    ].copy()

    if subset.empty or metric not in subset.columns:
        return (
            f"No valid {group_type.lower()} summary was available for "
            f"{metric_label}."
        )

    subset = subset.dropna(
        subset=[metric]
    )

    if subset.empty:
        return (
            f"No valid {group_type.lower()} summary was available for "
            f"{metric_label}."
        )

    best = subset.loc[
        subset[metric].idxmax()
    ]

    return (
        f"For {metric_label}, the highest mean value in the {group_type.lower()} "
        f"comparison was recorded by {best['group']} "
        f"({format_number(best[metric])})."
    )


def describe_chart_file(
    document,
    chart_name: str,
    chart_path: Path | None,
    description: str,
):
    document.add_paragraph(
        f"{chart_name}: {description}"
    )

    if chart_path is not None:
        document.add_paragraph(
            f"Chart file: {Path(chart_path).name}"
        )


def describe_output_file(
    document,
    filename: str,
    description: str,
):
    paragraph = document.add_paragraph()
    paragraph.add_run(filename).bold = True
    paragraph.add_run(f": {description}")


def create_word_report(
    output_path: Path,
    tray_metrics: pd.DataFrame,
    group_growth: pd.DataFrame,
    inside_outside_comparison: pd.DataFrame,
    adjusted_cell: pd.DataFrame,
    charts: dict[str, Path],
):
    if not DOCX_AVAILABLE:
        print(
            "WARNING: python-docx is not installed. Word report was skipped."
        )
        print(
            "Install it with: pip install python-docx"
        )
        return None

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    document = Document()

    styles = document.styles

    styles["Normal"].font.name = "Times New Roman"
    styles["Normal"].font.size = Pt(11)

    for style_name in [
        "Title",
        "Heading 1",
        "Heading 2",
        "Heading 3",
    ]:
        if style_name in styles:
            styles[style_name].font.name = "Times New Roman"

    title = document.add_heading(
        "Trial 3 RGB Growth-Rate and Treatment Comparison — Short Report",
        level=0,
    )

    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_paragraph(
        "This report summarises the Trial 3 RGB visible-emergence and "
        "green-cover analysis. The purpose of this script was to convert the "
        "cell-level outputs from Script 04 into tray-level and treatment-level "
        "growth comparisons. The report also documents the generated CSV, Excel, "
        "chart, PDF, and Word outputs so that the results can be interpreted "
        "and reused in the final internship report."
    )

    document.add_paragraph(
        "The analysis is based on visible seedling evidence detected from RGB "
        "images. Therefore, the results should be interpreted as visible "
        "emergence and image-derived green-cover measurements, not as direct "
        "biological germination measurements below the soil surface."
    )

    document.add_heading(
        "1. Input data and processing summary",
        level=1,
    )

    tray_count = int(tray_metrics["tray_no"].nunique()) if not tray_metrics.empty else 0
    imputed_count = int(adjusted_cell["day7_imputed"].eq("Yes").sum()) if "day7_imputed" in adjusted_cell.columns else 0
    cell_record_count = len(adjusted_cell)

    document.add_paragraph(
        f"The RGB comparison used {tray_count} trays, with each tray containing "
        f"70 grid cells. The adjusted cell table contained {cell_record_count} "
        f"cell-day records. The analysis covered the corrected seven observation "
        f"days from Day 1 (29/06/2026) to Day 7 (07/07/2026)."
    )

    document.add_paragraph(
        f"A total of {imputed_count} Day 7 cell records were flagged for adjusted "
        f"estimation. These were cells where earlier visible green evidence was "
        f"detected but Day 7 showed missing or no current green evidence. The "
        f"adjustment was made only in the separate adjusted output columns; the "
        f"observed Day 7 values were preserved unchanged."
    )

    document.add_paragraph(
        "The main metrics calculated in this script were final tracked visible "
        "emergence percentage, observed RGB green-cover percentage, adjusted RGB "
        "green-cover percentage, Day 1 to Day 7 growth rates, Day 6 to Day 7 "
        "change, possible Day 7 missing cells, and an overall descriptive RGB "
        "performance score."
    )

    document.add_heading(
        "2. Tray-level performance summary",
        level=1,
    )

    best = best_row_by_metric(
        tray_metrics,
        "overall_adjusted_rgb_score",
    )

    if best is not None:
        document.add_paragraph(
            f"The highest adjusted RGB performance score was recorded by "
            f"{best['tray']} ({best['microbe_status']} | {best['treatment']} | "
            f"{best['environment_group']}). Its adjusted RGB score was "
            f"{format_number(best['overall_adjusted_rgb_score'])}. This ranking "
            f"combines final visible emergence, final adjusted RGB green-cover, "
            f"growth-rate behaviour, and emergence timing into one descriptive "
            f"score."
        )

    document.add_paragraph(
        "The table below lists the tray-level ranking using the adjusted RGB "
        "performance score. This table is useful for identifying the trays that "
        "performed best overall after accounting for possible Day 7 missing crop "
        "issues."
    )

    tray_columns = [
        "tray",
        "microbe_status",
        "treatment",
        "environment_group",
        "final_tracked_emergence_percent",
        "final_adjusted_green_cover_percent",
        "day7_imputed_cells",
        "overall_adjusted_rgb_score",
        "overall_adjusted_rgb_rank",
    ]

    tray_columns = [
        column
        for column in tray_columns
        if column in tray_metrics.columns
    ]

    add_docx_table(
        document,
        tray_metrics.sort_values(
            "overall_adjusted_rgb_rank",
            na_position="last",
        ),
        tray_columns,
        max_rows=12,
    )

    add_picture_if_exists(
        document,
        charts.get("tray_ranking"),
    )

    document.add_heading(
        "3. Microbes vs No Microbes comparison",
        level=1,
    )

    document.add_paragraph(
        group_summary_sentence(
            group_growth,
            "Microbe Status",
            "mean_final_adjusted_green_cover_percent",
            "final adjusted RGB green-cover",
        )
    )

    document.add_paragraph(
        group_summary_sentence(
            group_growth,
            "Microbe Status",
            "mean_overall_adjusted_rgb_score",
            "overall adjusted RGB performance score",
        )
    )

    document.add_paragraph(
        "This comparison groups trays by Microbes and No Microbes. It helps "
        "evaluate whether the microbial treatment was associated with stronger "
        "visible emergence or higher green-cover development across the trial."
    )

    microbe_group = group_growth.loc[
        group_growth["group_type"].eq("Microbe Status")
    ].copy()

    if not microbe_group.empty:
        microbe_columns = [
            "group",
            "tray_count",
            "mean_final_tracked_emergence_percent",
            "mean_final_adjusted_green_cover_percent",
            "mean_adjusted_green_cover_rate_day1_to_day7_pp_per_day",
            "mean_overall_adjusted_rgb_score",
        ]

        microbe_columns = [
            column
            for column in microbe_columns
            if column in microbe_group.columns
        ]

        add_docx_table(
            document,
            microbe_group,
            microbe_columns,
            max_rows=5,
        )

    add_picture_if_exists(
        document,
        charts.get("microbes_green"),
    )

    document.add_paragraph(
        "The Microbes vs No Microbes green-cover trend chart shows how the "
        "mean adjusted RGB green-cover changed across the seven observation "
        "days. This chart is useful for checking whether one group developed "
        "faster or maintained stronger green-cover over time."
    )

    document.add_heading(
        "4. Ideal vs Heat vs Moisture comparison",
        level=1,
    )

    document.add_paragraph(
        group_summary_sentence(
            group_growth,
            "Treatment Type",
            "mean_final_adjusted_green_cover_percent",
            "final adjusted RGB green-cover",
        )
    )

    document.add_paragraph(
        group_summary_sentence(
            group_growth,
            "Treatment Type",
            "mean_overall_adjusted_rgb_score",
            "overall adjusted RGB performance score",
        )
    )

    document.add_paragraph(
        "This comparison groups trays by treatment type: Ideal, Heat, and "
        "Moisture. The Ideal treatment represents trays kept fully inside or "
        "outside without the scheduled heat or moisture stress. Heat trays were "
        "moved according to the heat-treatment schedule, while Moisture trays "
        "followed the restricted watering schedule."
    )

    treatment_group = group_growth.loc[
        group_growth["group_type"].eq("Treatment Type")
    ].copy()

    if not treatment_group.empty:
        treatment_columns = [
            "group",
            "tray_count",
            "mean_final_tracked_emergence_percent",
            "mean_final_adjusted_green_cover_percent",
            "mean_adjusted_green_cover_rate_day1_to_day7_pp_per_day",
            "mean_overall_adjusted_rgb_score",
        ]

        treatment_columns = [
            column
            for column in treatment_columns
            if column in treatment_group.columns
        ]

        add_docx_table(
            document,
            treatment_group,
            treatment_columns,
            max_rows=10,
        )

    add_picture_if_exists(
        document,
        charts.get("treatment_green"),
    )

    document.add_paragraph(
        "The treatment green-cover chart compares Ideal, Heat, and Moisture "
        "groups over time. It is useful for observing whether stress treatments "
        "reduced growth, delayed emergence, or produced recovery patterns later "
        "in the trial."
    )

    document.add_heading(
        "5. Inside vs Outside comparison for Ideal and Moisture trays",
        level=1,
    )

    document.add_paragraph(
        "A direct Inside vs Outside comparison was performed only for Ideal and "
        "Moisture trays because these treatments had fixed placement. Heat trays "
        "were excluded from this direct comparison because their location changed "
        "during the heat-treatment schedule. This prevents dynamic Heat trays "
        "from being incorrectly treated as permanently inside or outside."
    )

    if not inside_outside_comparison.empty:
        inside_columns = [
            "treatment",
            "inside_tray_count",
            "outside_tray_count",
            "inside_final_emergence_percent",
            "outside_final_emergence_percent",
            "inside_minus_outside_final_emergence_pp",
            "inside_final_adjusted_green_cover_percent",
            "outside_final_adjusted_green_cover_percent",
            "inside_minus_outside_adjusted_green_cover_pp",
            "inside_minus_outside_adjusted_growth_rate_pp_per_day",
            "interpretation",
        ]

        inside_columns = [
            column
            for column in inside_columns
            if column in inside_outside_comparison.columns
        ]

        add_docx_table(
            document,
            inside_outside_comparison,
            inside_columns,
            max_rows=5,
        )

        for _, row in inside_outside_comparison.iterrows():
            document.add_paragraph(
                f"For {row['treatment']} trays, the inside-minus-outside "
                f"difference in adjusted Day 7 RGB green-cover was "
                f"{format_number(row.get('inside_minus_outside_adjusted_green_cover_pp'))} "
                f"percentage points. Interpretation: {row.get('interpretation', '')}"
            )

    add_picture_if_exists(
        document,
        charts.get("inside_outside_green"),
    )

    document.add_paragraph(
        "The Inside vs Outside green-cover chart compares the adjusted final "
        "green-cover percentage for Ideal and Moisture trays. This chart directly "
        "answers whether fixed inside placement or fixed outside placement was "
        "associated with stronger visible crop development within those treatment "
        "types."
    )

    add_picture_if_exists(
        document,
        charts.get("inside_outside_emergence"),
    )

    document.add_paragraph(
        "The Inside vs Outside emergence chart compares final tracked visible "
        "emergence. It is different from the green-cover chart because emergence "
        "measures how many cells showed visible seedlings, while green-cover "
        "measures the relative amount of green plant area inside the cell zones."
    )

    document.add_heading(
        "6. Observed vs adjusted Day 7 results",
        level=1,
    )

    document.add_paragraph(
        "The observed Day 7 results represent what was actually visible in the "
        "final images. The adjusted Day 7 values are a separate scenario created "
        "for growth-rate comparison when a cell had earlier visible crop but no "
        "green evidence on Day 7. This adjustment is important because some "
        "seedlings may have been eaten by bugs or otherwise disappeared before "
        "the final image set."
    )

    document.add_paragraph(
        "The observed and adjusted values are not mixed silently. The adjusted "
        "cell table contains columns such as day7_imputed, imputation_reason, "
        "imputation_method, previous_growth_rate_pp_per_day, and "
        "adjustment_difference_pp so that every estimated value can be traced."
    )

    add_picture_if_exists(
        document,
        charts.get("day7_observed_adjusted"),
    )

    document.add_paragraph(
        "The observed vs adjusted Day 7 chart shows how much the final tray-level "
        "green-cover changed after applying the Day 7 missing-crop adjustment. "
        "Large differences indicate trays where many cells had earlier growth but "
        "lower visible green area in the final image."
    )

    add_picture_if_exists(
        document,
        charts.get("bug_cells"),
    )

    document.add_paragraph(
        "The possible bug-eaten cell chart shows the number of Day 7 cells flagged "
        "for adjustment in each tray. This chart helps identify which trays were "
        "most affected by missing final-day crop evidence."
    )

    document.add_heading(
        "7. Description of generated CSV files",
        level=1,
    )

    describe_output_file(
        document,
        "cell_growth_with_day7_adjustment.csv",
        "The full cell-level dataset after Day 7 adjustment processing. It preserves observed green-cover values and adds adjusted green-cover columns, imputation flags, and imputation reasons for cells affected by possible Day 7 crop disappearance.",
    )

    describe_output_file(
        document,
        "tray_daily_rgb_metrics.csv",
        "A tray-by-day summary table. It contains daily visible emergence percentage, observed green-cover percentage, adjusted green-cover percentage, newly emerged cells, and Day 7 imputed cell counts for each tray.",
    )

    describe_output_file(
        document,
        "tray_growth_metrics.csv",
        "A tray-level growth summary. It includes final emergence, final observed and adjusted green-cover, Day 1 to Day 7 growth rates, Day 6 to Day 7 changes, T50 timing, and overall RGB performance scores.",
    )

    describe_output_file(
        document,
        "group_daily_metrics.csv",
        "A group-by-day trend table. It summarises daily group averages for Microbes vs No Microbes, treatment groups, treatment-by-environment groups, and other comparison categories.",
    )

    describe_output_file(
        document,
        "group_growth_metrics.csv",
        "A group-level final comparison table. It summarises final performance and growth-rate metrics for each treatment or comparison group.",
    )

    describe_output_file(
        document,
        "inside_outside_comparison_ideal_moisture.csv",
        "A dedicated comparison table for Ideal and Moisture trays only. It compares Inside vs Outside final emergence, adjusted green-cover, growth rate, and overall adjusted RGB score.",
    )

    describe_output_file(
        document,
        "possible_day7_bug_eaten_cells.csv",
        "A filtered table containing only cells that were flagged for Day 7 adjustment because they had earlier visible green evidence but were missing or not visible on Day 7.",
    )

    describe_output_file(
        document,
        "heat_phase_response.csv",
        "A descriptive summary of Heat tray performance across the heat-treatment phases, including inside-before-heat, outside heat exposure, and inside recovery.",
    )

    describe_output_file(
        document,
        "moisture_phase_response.csv",
        "A descriptive summary of Moisture tray performance across watering and dry phases.",
    )

    document.add_heading(
        "8. Description of Excel workbook",
        level=1,
    )

    document.add_paragraph(
        "The Excel workbook rgb_growth_treatment_report.xlsx combines the main "
        "CSV outputs into one file with multiple sheets. It is the easiest file "
        "to inspect manually because it contains tray-level, group-level, "
        "inside-vs-outside, Day 7 adjustment, heat-phase, and moisture-phase "
        "outputs in one place."
    )

    describe_output_file(
        document,
        "Tray Growth Metrics sheet",
        "Contains one row per tray and is the main sheet for ranking trays by final performance, growth rate, and adjusted RGB score.",
    )

    describe_output_file(
        document,
        "Tray Daily Metrics sheet",
        "Contains one row per tray per observation day and is useful for checking daily progression.",
    )

    describe_output_file(
        document,
        "Group Daily Metrics sheet",
        "Contains group-level daily trends used to create line charts.",
    )

    describe_output_file(
        document,
        "Group Growth Metrics sheet",
        "Contains final group-level comparison metrics for Microbes, treatments, and treatment-environment groups.",
    )

    describe_output_file(
        document,
        "Inside Outside Compare sheet",
        "Contains the dedicated Ideal and Moisture Inside vs Outside comparison.",
    )

    describe_output_file(
        document,
        "Day7 Imputed Cells sheet",
        "Contains only the cells where adjusted Day 7 values were estimated.",
    )

    describe_output_file(
        document,
        "Heat Phase Response sheet",
        "Summarises performance of Heat trays across the heat-treatment movement phases.",
    )

    describe_output_file(
        document,
        "Moisture Phase Response sheet",
        "Summarises performance of Moisture trays across watering and non-watering phases.",
    )

    document.add_heading(
        "9. Description of generated charts",
        level=1,
    )

    describe_chart_file(
        document,
        "01_microbes_visible_emergence_trend.png",
        charts.get("microbes_emergence"),
        "Shows the tracked visible emergence trend for Microbes and No Microbes groups across observation days.",
    )

    describe_chart_file(
        document,
        "02_microbes_adjusted_green_cover_trend.png",
        charts.get("microbes_green"),
        "Shows adjusted RGB green-cover over time for Microbes and No Microbes groups.",
    )

    describe_chart_file(
        document,
        "03_treatment_visible_emergence_trend.png",
        charts.get("treatment_emergence"),
        "Shows tracked visible emergence trends for Ideal, Heat, and Moisture treatments.",
    )

    describe_chart_file(
        document,
        "04_treatment_adjusted_green_cover_trend.png",
        charts.get("treatment_green"),
        "Shows adjusted RGB green-cover trends for Ideal, Heat, and Moisture treatments.",
    )

    describe_chart_file(
        document,
        "05_day7_observed_vs_adjusted_green_cover_by_tray.png",
        charts.get("day7_observed_adjusted"),
        "Compares observed Day 7 green-cover with adjusted Day 7 green-cover for each tray.",
    )

    describe_chart_file(
        document,
        "06_adjusted_rgb_tray_ranking.png",
        charts.get("tray_ranking"),
        "Ranks trays by overall adjusted RGB performance score.",
    )

    describe_chart_file(
        document,
        "07_day7_possible_bug_eaten_cells_by_tray.png",
        charts.get("bug_cells"),
        "Shows how many cells in each tray were flagged as possible Day 7 missing or bug-eaten cells.",
    )

    describe_chart_file(
        document,
        "08_inside_outside_adjusted_green_cover_ideal_moisture.png",
        charts.get("inside_outside_green"),
        "Compares Inside vs Outside adjusted Day 7 RGB green-cover for Ideal and Moisture trays.",
    )

    describe_chart_file(
        document,
        "09_inside_outside_emergence_ideal_moisture.png",
        charts.get("inside_outside_emergence"),
        "Compares Inside vs Outside final visible emergence for Ideal and Moisture trays.",
    )

    document.add_heading(
        "10. Interpretation and limitations",
        level=1,
    )

    document.add_paragraph(
        "The results are useful for identifying visual growth trends and treatment "
        "differences in Trial 3. However, they should be interpreted carefully. "
        "The measurements are based on image-derived visible green area, not "
        "direct plant biomass or calibrated reflectance. Lighting, image quality, "
        "crop overlap, and cell visibility can affect RGB green-cover estimates."
    )

    document.add_paragraph(
        "The Day 7 adjustment is a practical correction for likely bug-eaten or "
        "missing seedlings. It should be used only for adjusted growth-rate "
        "comparison, while the observed Day 7 results should remain the source "
        "for reporting what was actually visible in the images."
    )

    document.add_paragraph(
        "The Inside vs Outside comparison is restricted to Ideal and Moisture "
        "trays because Heat trays were moved between environments. Including "
        "Heat trays in a fixed Inside vs Outside comparison would be misleading."
    )

    document.add_paragraph(
        "Because each treatment group contains a small number of trays, these "
        "results should be treated as descriptive evidence rather than formal "
        "statistical proof. The later multispectral scripts should be used to "
        "check whether NDVI and NDRE trends support the RGB results."
    )

    document.save(output_path)

    return output_path
# ============================================================
# 15) PDF REPORT
# ============================================================

def make_pdf_table(
    dataframe: pd.DataFrame,
    columns: list[str],
    max_rows: int = 20,
):
    if not REPORTLAB_AVAILABLE:
        return None

    styles = getSampleStyleSheet()

    header_style = ParagraphStyle(
        "HeaderCell",
        parent=styles["BodyText"],
        fontSize=7,
        leading=8,
        textColor=colors.white,
        alignment=1,
        wordWrap="CJK",
    )

    cell_style = ParagraphStyle(
        "BodyCell",
        parent=styles["BodyText"],
        fontSize=7,
        leading=8,
        wordWrap="CJK",
    )

    frame = dataframe[columns].head(max_rows).copy()
    frame = safe_round_dataframe(frame, 3)

    data = [
        [
            Paragraph(str(column), header_style)
            for column in columns
        ]
    ]

    for _, row in frame.iterrows():
        data.append(
            [
                Paragraph(str(row[column]), cell_style)
                for column in columns
            ]
        )

    table = Table(
        data,
        repeatRows=1,
    )

    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D1D5DB")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
            ]
        )
    )

    return table


def create_pdf_report(
    output_path: Path,
    tray_metrics: pd.DataFrame,
    group_growth: pd.DataFrame,
    inside_outside_comparison: pd.DataFrame,
    charts: dict[str, Path],
):
    if not REPORTLAB_AVAILABLE:
        print(
            "WARNING: reportlab is not installed. PDF report was skipped."
        )
        return None

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    document = SimpleDocTemplate(
        str(output_path),
        pagesize=landscape(A4),
        rightMargin=0.35 * inch,
        leftMargin=0.35 * inch,
        topMargin=0.35 * inch,
        bottomMargin=0.35 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(
        Paragraph(
            "Trial 3 RGB Growth-Rate and Treatment Comparison Report",
            styles["Title"],
        )
    )

    story.append(Spacer(1, 10))

    story.append(
        Paragraph(
            "This report summarises Trial 3 visible-emergence and RGB green-cover "
            "growth results. Observed Day 7 values are preserved, while adjusted "
            "Day 7 values are reported separately for likely missing or bug-eaten seedlings.",
            styles["BodyText"],
        )
    )

    story.append(Spacer(1, 12))

    top_columns = [
        "tray",
        "microbe_status",
        "treatment",
        "environment_group",
        "final_tracked_emergence_percent",
        "final_observed_green_cover_percent",
        "final_adjusted_green_cover_percent",
        "day7_imputed_cells",
        "overall_adjusted_rgb_score",
        "overall_adjusted_rgb_rank",
    ]

    top_columns = [
        column
        for column in top_columns
        if column in tray_metrics.columns
    ]

    story.append(
        Paragraph(
            "Tray-level adjusted RGB ranking",
            styles["Heading2"],
        )
    )

    story.append(
        make_pdf_table(
            tray_metrics.sort_values(
                "overall_adjusted_rgb_rank",
                na_position="last",
            ),
            top_columns,
            max_rows=12,
        )
    )

    story.append(PageBreak())

    story.append(
        Paragraph(
            "Ideal and Moisture Inside vs Outside comparison",
            styles["Heading2"],
        )
    )

    inside_columns = [
        "treatment",
        "inside_tray_count",
        "outside_tray_count",
        "inside_final_emergence_percent",
        "outside_final_emergence_percent",
        "inside_minus_outside_final_emergence_pp",
        "inside_final_adjusted_green_cover_percent",
        "outside_final_adjusted_green_cover_percent",
        "inside_minus_outside_adjusted_green_cover_pp",
        "interpretation",
    ]

    inside_columns = [
        column
        for column in inside_columns
        if column in inside_outside_comparison.columns
    ]

    story.append(
        make_pdf_table(
            inside_outside_comparison,
            inside_columns,
            max_rows=5,
        )
    )

    figure_items = [
        ("Microbes adjusted green-cover trend", charts.get("microbes_green")),
        ("Treatment adjusted green-cover trend", charts.get("treatment_green")),
        ("Inside vs Outside adjusted green-cover", charts.get("inside_outside_green")),
        ("Day 7 observed vs adjusted green-cover", charts.get("day7_observed_adjusted")),
        ("Adjusted RGB tray ranking", charts.get("tray_ranking")),
        ("Possible Day 7 bug-eaten cells", charts.get("bug_cells")),
    ]

    for title, path in figure_items:
        if path is None or not Path(path).exists():
            continue

        story.append(PageBreak())
        story.append(
            Paragraph(
                title,
                styles["Heading2"],
            )
        )
        story.append(Spacer(1, 6))
        story.append(
            PDFImage(
                str(path),
                width=10.5 * inch,
                height=5.9 * inch,
            )
        )

    document.build(story)

    return output_path


# ============================================================
# 16) SETTINGS
# ============================================================

def save_settings(
    path: Path,
):
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    settings = {
        "purpose": "Trial 3 RGB growth-rate and treatment comparison",
        "script04_tray_summary_csv": str(TRAY_SUMMARY_CSV),
        "script04_cell_results_csv": str(CELL_RESULTS_CSV),
        "output_root": str(OUTPUT_ROOT),
        "expected_cells_per_tray": EXPECTED_CELLS,
        "expected_observation_days": EXPECTED_OBSERVATION_DAYS,
        "corrected_day1_photo_date": "2026-06-29",
        "day7_photo_date": "2026-07-07",
        "inside_outside_comparison_policy": (
            "Direct Inside vs Outside comparison is only performed for Ideal and Moisture trays. "
            "Heat trays are excluded because their environment changed during the trial."
        ),
        "observed_day7_policy": (
            "Observed Day 7 values from Script 04 are preserved exactly."
        ),
        "adjusted_day7_policy": (
            "If a cell had prior green evidence but no current green evidence "
            "on Day 7, an adjusted Day 7 green-cover value is estimated from "
            "the most recent prior green-cover value and previous positive "
            "growth slope."
        ),
        "word_report": (
            "A short Word report is generated for this comparison script."
        ),
        "statistical_warning": (
            "This script produces descriptive image-derived comparisons only; "
            "it is not a formal statistical test."
        ),
    }

    with path.open(
        "w",
        encoding="utf-8",
    ) as file:
        json.dump(
            settings,
            file,
            indent=2,
        )

    return path


# ============================================================
# 17) MAIN
# ============================================================

def run_analysis(args) -> int:
    print("\nSCRIPT 05 — THIRD TRIAL RGB GROWTH-RATE AND TREATMENT COMPARISON")
    print("=" * 78)
    print(f"Script 04 reports folder:\n{SCRIPT04_REPORTS}")
    print(f"\nOutput folder:\n{OUTPUT_ROOT}")

    tray_summary, cell_results, first_summary = load_script04_outputs()

    print(f"\nLoaded Script 04 tray rows: {len(tray_summary)}")
    print(f"Loaded Script 04 cell rows: {len(cell_results)}")

    day_counts = (
        tray_summary.groupby("day_order")["tray_no"]
        .nunique()
        .reset_index(name="tray_count")
    )

    print("\nTray count by observation day:")
    print(day_counts.to_string(index=False))

    missing_days = set(EXPECTED_OBSERVATION_DAYS) - set(
        tray_summary["day_order"].astype(int).unique()
    )

    if missing_days:
        print(
            "\nWARNING: Missing expected observation days: "
            + ", ".join(
                f"Day {day}"
                for day in sorted(missing_days)
            )
        )

    adjusted_cell = create_adjusted_cell_table(
        cell_results
    )

    tray_daily = create_tray_daily_metrics(
        adjusted_cell
    )

    tray_metrics = create_tray_growth_metrics(
        tray_daily,
        adjusted_cell,
    )

    group_daily, group_growth = create_all_group_tables(
        tray_daily,
        tray_metrics,
    )

    heat_phase, moisture_phase = create_phase_tables(
        tray_daily
    )

    inside_outside_comparison = create_inside_outside_comparison(
        tray_metrics
    )

    charts = create_charts(
        group_daily,
        tray_metrics,
        inside_outside_comparison,
    )

    table_paths = save_tables(
        adjusted_cell,
        tray_daily,
        tray_metrics,
        group_daily,
        group_growth,
        heat_phase,
        moisture_phase,
        inside_outside_comparison,
    )

    pdf_path = REPORTS_ROOT / "rgb_growth_treatment_summary.pdf"

    create_pdf_report(
        pdf_path,
        tray_metrics,
        group_growth,
        inside_outside_comparison,
        charts,
    )

    docx_path = REPORTS_ROOT / "rgb_growth_treatment_short_report.docx"

    create_word_report(
        docx_path,
        tray_metrics,
        group_growth,
        inside_outside_comparison,
        adjusted_cell,
        charts,
    )

    settings_path = save_settings(
        CONFIG_ROOT / "rgb_growth_treatment_settings.json"
    )

    print("\n" + "=" * 78)
    print("SCRIPT 05 FINISHED")
    print("=" * 78)

    day7_imputed_cells = int(
        adjusted_cell["day7_imputed"].eq("Yes").sum()
    )

    print(f"Day 7 imputed cell records: {day7_imputed_cells}")

    if not tray_metrics.empty:
        best = tray_metrics.sort_values(
            "overall_adjusted_rgb_rank",
            na_position="last",
        ).iloc[0]

        print(
            "\nTop adjusted RGB tray:"
            f"\n  {best['tray']} — {best['microbe_status']} | "
            f"{best['treatment']} | {best['environment_group']}"
            f"\n  Adjusted RGB score: {best['overall_adjusted_rgb_score']:.2f}"
        )

    print("\nMain output files:")
    print(f"Adjusted cell table:\n{table_paths['adjusted_cell']}")
    print(f"\nTray daily metrics:\n{table_paths['tray_daily']}")
    print(f"\nTray growth metrics:\n{table_paths['tray_metrics']}")
    print(f"\nGroup growth metrics:\n{table_paths['group_growth']}")
    print(f"\nInside vs Outside comparison:\n{table_paths['inside_outside']}")
    print(f"\nDay 7 imputed cells:\n{table_paths['day7_imputed']}")
    print(f"\nExcel report:\n{table_paths['excel']}")

    if REPORTLAB_AVAILABLE:
        print(f"\nPDF report:\n{pdf_path}")
    else:
        print("\nPDF report skipped because reportlab is not installed.")

    if DOCX_AVAILABLE:
        print(f"\nWord report:\n{docx_path}")
    else:
        print("\nWord report skipped because python-docx is not installed.")

    print(f"\nSettings:\n{settings_path}")
    print(f"\nCharts folder:\n{CHARTS_ROOT}")

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Trial 3 Script 05: RGB growth-rate and treatment comparison."
        )
    )

    args = parser.parse_args()

    return run_analysis(args)


if __name__ == "__main__":
    raise SystemExit(main())