from __future__ import annotations

"""
SCRIPT 10 - SECOND TRIAL REPORT FIGURE PACKAGE

Purpose
-------
Collect the final report-ready outputs from the Second Trial workflow into one
clean package folder.

This script does not perform new analysis. It copies selected charts, reports,
Excel files, and summary CSV files from Scripts 05, 08 and 09.

Outputs
-------
outputs/Second Trial/10_Report_Figure_Package/
    figures/
    reports/
    tables/
    _report_index/
        report_package_manifest.csv
        report_package_index.pdf
        README_report_package.txt

Reason
------
The earlier scripts produce many intermediate files. This script creates a clean
report package so the final internship report can be written without searching
through multiple output folders.
"""

import argparse
import csv
import json
import shutil
from datetime import datetime
from pathlib import Path

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


# ============================================================
# 1) PATHS — CHANGE PROJECT_ROOT ONLY WHEN REUSING
# ============================================================

PROJECT_ROOT = Path(r"C:\Users\tshib\OneDrive\Desktop\Internship")

SECOND_TRIAL_OUTPUT_ROOT = (
    PROJECT_ROOT
    / "outputs"
    / "Second Trial"
)

SCRIPT05_ROOT = (
    SECOND_TRIAL_OUTPUT_ROOT
    / "05_Treatment_Growth_Visuals"
)

SCRIPT08_ROOT = (
    SECOND_TRIAL_OUTPUT_ROOT
    / "08_MS_Treatment_Comparison"
)

SCRIPT09_ROOT = (
    SECOND_TRIAL_OUTPUT_ROOT
    / "09_Second_Trial_Synthesis"
)

OUTPUT_ROOT = (
    SECOND_TRIAL_OUTPUT_ROOT
    / "10_Report_Figure_Package"
)

FIGURES_ROOT = OUTPUT_ROOT / "figures"
REPORTS_ROOT = OUTPUT_ROOT / "reports"
TABLES_ROOT = OUTPUT_ROOT / "tables"
INDEX_ROOT = OUTPUT_ROOT / "_report_index"


# ============================================================
# 2) FILE COLLECTION PLAN
# ============================================================

FIGURE_ITEMS = [
    {
        "label": "RGB emergence trend by treatment",
        "target_name": "01_rgb_emergence_trend.png",
        "source_candidates": [
            SCRIPT05_ROOT / "charts" / "01_microbes_vs_no_microbes_emergence.png",
            SCRIPT05_ROOT / "charts" / "01_treatment_tracked_emergence_trend.png",
        ],
    },
    {
        "label": "RGB green-cover trend by treatment",
        "target_name": "02_rgb_green_cover_treatment.png",
        "source_candidates": [
            SCRIPT05_ROOT / "charts" / "02_microbes_vs_no_microbes_green_cover.png",
            SCRIPT05_ROOT / "charts" / "02_treatment_rgb_green_cover_trend.png",
        ],
    },
    {
        "label": "RGB emergence by treatment and environment",
        "target_name": "03_rgb_interaction_emergence.png",
        "source_candidates": [
            SCRIPT05_ROOT / "charts" / "03_interaction_emergence.png",
            SCRIPT05_ROOT / "charts" / "03_interaction_tracked_emergence_trend.png",
        ],
    },
    {
        "label": "RGB green-cover by treatment and environment",
        "target_name": "04_rgb_interaction_green_cover.png",
        "source_candidates": [
            SCRIPT05_ROOT / "charts" / "04_interaction_green_cover.png",
            SCRIPT05_ROOT / "charts" / "04_interaction_rgb_green_cover_trend.png",
        ],
    },
    {
        "label": "MS NDVI treatment trend",
        "target_name": "05_ms_ndvi_treatment.png",
        "source_candidates": [
            SCRIPT08_ROOT / "charts" / "01_microbes_vs_no_microbes_ndvi.png",
        ],
    },
    {
        "label": "MS NDRE treatment trend",
        "target_name": "06_ms_ndre_treatment.png",
        "source_candidates": [
            SCRIPT08_ROOT / "charts" / "02_microbes_vs_no_microbes_ndre.png",
        ],
    },
    {
        "label": "MS NDVI by treatment and environment",
        "target_name": "07_ms_interaction_ndvi.png",
        "source_candidates": [
            SCRIPT08_ROOT / "charts" / "03_interaction_ndvi.png",
        ],
    },
    {
        "label": "MS NDRE by treatment and environment",
        "target_name": "08_ms_interaction_ndre.png",
        "source_candidates": [
            SCRIPT08_ROOT / "charts" / "04_interaction_ndre.png",
        ],
    },
    {
        "label": "Final tray performance ranking",
        "target_name": "09_final_tray_ranking.png",
        "source_candidates": [
            SCRIPT09_ROOT / "charts" / "01_final_tray_score_ranking.png",
        ],
    },
    {
        "label": "Final treatment group ranking",
        "target_name": "10_final_group_ranking.png",
        "source_candidates": [
            SCRIPT09_ROOT / "charts" / "02_final_group_score_ranking.png",
        ],
    },
    {
        "label": "Day 5 RGB green cover versus NDVI",
        "target_name": "11_rgb_vs_ndvi.png",
        "source_candidates": [
            SCRIPT09_ROOT / "charts" / "03_day5_rgb_green_cover_vs_ndvi.png",
        ],
    },
    {
        "label": "Day 5 RGB green cover versus NDRE",
        "target_name": "12_rgb_vs_ndre.png",
        "source_candidates": [
            SCRIPT09_ROOT / "charts" / "04_day5_rgb_green_cover_vs_ndre.png",
        ],
    },
    {
        "label": "Treatment summary heatmap",
        "target_name": "13_treatment_summary_heatmap.png",
        "source_candidates": [
            SCRIPT09_ROOT / "charts" / "05_treatment_summary_heatmap.png",
        ],
    },
]

REPORT_ITEMS = [
    {
        "label": "Final Second Trial synthesis PDF report",
        "target_name": "01_second_trial_synthesis_report.pdf",
        "source_candidates": [
            SCRIPT09_ROOT / "_reports" / "second_trial_synthesis_report.pdf",
        ],
    },
    {
        "label": "Multispectral treatment comparison PDF report",
        "target_name": "02_ms_treatment_visual_summary.pdf",
        "source_candidates": [
            SCRIPT08_ROOT / "_reports" / "ms_treatment_visual_summary.pdf",
        ],
    },
    {
        "label": "Final Second Trial master Excel workbook",
        "target_name": "03_second_trial_master_summary.xlsx",
        "source_candidates": [
            SCRIPT09_ROOT / "_reports" / "second_trial_master_summary.xlsx",
        ],
    },
    {
        "label": "Multispectral treatment comparison Excel workbook",
        "target_name": "04_ms_treatment_comparison_report.xlsx",
        "source_candidates": [
            SCRIPT08_ROOT / "_reports" / "ms_treatment_comparison_report.xlsx",
        ],
    },
    {
        "label": "RGB treatment growth Excel workbook",
        "target_name": "05_treatment_growth_report.xlsx",
        "source_candidates": [
            SCRIPT05_ROOT / "_reports" / "treatment_growth_report.xlsx",
        ],
    },
]

TABLE_ITEMS = [
    {
        "label": "Final master summary CSV",
        "target_name": "01_final_master_summary.csv",
        "source_candidates": [
            SCRIPT09_ROOT / "_reports" / "second_trial_master_summary.csv",
        ],
    },
    {
        "label": "Final group summary CSV",
        "target_name": "02_final_group_summary.csv",
        "source_candidates": [
            SCRIPT09_ROOT / "_reports" / "second_trial_group_summary.csv",
        ],
    },
    {
        "label": "RGB tray growth metrics CSV",
        "target_name": "03_rgb_tray_growth_metrics.csv",
        "source_candidates": [
            SCRIPT05_ROOT / "_reports" / "tray_growth_metrics.csv",
        ],
    },
    {
        "label": "RGB group daily metrics CSV",
        "target_name": "04_rgb_group_daily_metrics.csv",
        "source_candidates": [
            SCRIPT05_ROOT / "_reports" / "group_daily_metrics.csv",
        ],
    },
    {
        "label": "MS tray index metrics CSV",
        "target_name": "05_ms_tray_index_metrics.csv",
        "source_candidates": [
            SCRIPT08_ROOT / "_reports" / "ms_tray_index_metrics.csv",
        ],
    },
    {
        "label": "MS group growth rates CSV",
        "target_name": "06_ms_group_growth_rates.csv",
        "source_candidates": [
            SCRIPT08_ROOT / "_reports" / "ms_group_growth_rates.csv",
        ],
    },
]


# ============================================================
# 3) HELPERS
# ============================================================

def ensure_folders() -> None:
    for folder in [
        FIGURES_ROOT,
        REPORTS_ROOT,
        TABLES_ROOT,
        INDEX_ROOT,
    ]:
        folder.mkdir(parents=True, exist_ok=True)


def find_existing_source(candidates: list[Path]) -> Path | None:
    for path in candidates:
        if path.exists() and path.is_file():
            return path

    return None


def copy_item(
    item: dict,
    target_folder: Path,
    category: str,
    overwrite: bool,
) -> dict:
    source = find_existing_source(item["source_candidates"])
    target_path = target_folder / item["target_name"]

    record = {
        "category": category,
        "label": item["label"],
        "target_name": item["target_name"],
        "source_path": "",
        "target_path": str(target_path),
        "status": "",
        "notes": "",
    }

    if source is None:
        record["status"] = "MISSING"
        record["notes"] = (
            "No source candidate was found. Check whether the previous "
            "script was completed and whether the output filename exists."
        )
        return record

    record["source_path"] = str(source)

    if target_path.exists() and not overwrite:
        record["status"] = "SKIPPED_EXISTING"
        record["notes"] = "Target file already exists. Use --overwrite to replace it."
        return record

    shutil.copy2(source, target_path)

    record["status"] = "COPIED"
    record["notes"] = "Copied successfully."
    return record


def write_manifest(path: Path, rows: list[dict]) -> None:
    fieldnames = [
        "category",
        "label",
        "target_name",
        "source_path",
        "target_path",
        "status",
        "notes",
    ]

    path.parent.mkdir(parents=True, exist_ok=True)

    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def count_status(rows: list[dict], status: str) -> int:
    return sum(row["status"] == status for row in rows)


def make_pdf_table(rows: list[dict]):
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

    data = [
        [
            Paragraph("Category", header_style),
            Paragraph("Item", header_style),
            Paragraph("File", header_style),
            Paragraph("Status", header_style),
            Paragraph("Notes", header_style),
        ]
    ]

    for row in rows:
        data.append(
            [
                Paragraph(str(row["category"]), cell_style),
                Paragraph(str(row["label"]), cell_style),
                Paragraph(str(row["target_name"]), cell_style),
                Paragraph(str(row["status"]), cell_style),
                Paragraph(str(row["notes"]), cell_style),
            ]
        )

    table = Table(
        data,
        repeatRows=1,
        colWidths=[
            1.1 * inch,
            2.4 * inch,
            2.4 * inch,
            1.2 * inch,
            3.5 * inch,
        ],
    )

    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D1D5DB")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F9FAFB")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    return table


def create_pdf_index(
    path: Path,
    manifest_rows: list[dict],
    figure_records: list[dict],
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    document = SimpleDocTemplate(
        str(path),
        pagesize=landscape(A4),
        rightMargin=0.35 * inch,
        leftMargin=0.35 * inch,
        topMargin=0.35 * inch,
        bottomMargin=0.35 * inch,
    )

    styles = getSampleStyleSheet()
    story = []

    copied_count = count_status(manifest_rows, "COPIED")
    skipped_count = count_status(manifest_rows, "SKIPPED_EXISTING")
    missing_count = count_status(manifest_rows, "MISSING")

    story.append(
        Paragraph(
            "Second Trial Report Figure Package Index",
            styles["Title"],
        )
    )

    story.append(Spacer(1, 10))

    story.append(
        Paragraph(
            "This PDF indexes the final report-ready package for the Second Trial. "
            "The package collects selected figures, PDF reports, Excel workbooks, "
            "and CSV tables from the completed RGB and multispectral workflows.",
            styles["BodyText"],
        )
    )

    story.append(Spacer(1, 10))

    story.append(
        Paragraph(
            f"Created: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>"
            f"Copied files: {copied_count}<br/>"
            f"Skipped existing files: {skipped_count}<br/>"
            f"Missing files: {missing_count}",
            styles["BodyText"],
        )
    )

    story.append(Spacer(1, 14))

    story.append(
        Paragraph(
            "Package Manifest",
            styles["Heading2"],
        )
    )

    story.append(make_pdf_table(manifest_rows))

    successful_figures = [
        row
        for row in figure_records
        if row["status"] in {"COPIED", "SKIPPED_EXISTING"}
        and (FIGURES_ROOT / row["target_name"]).exists()
    ]

    if successful_figures:
        story.append(PageBreak())

        story.append(
            Paragraph(
                "Collected Figure Preview",
                styles["Heading2"],
            )
        )

        first = True

        for row in successful_figures:
            image_path = FIGURES_ROOT / row["target_name"]

            if not image_path.exists():
                continue

            if not first:
                story.append(PageBreak())

            first = False

            story.append(
                Paragraph(
                    row["label"],
                    styles["Heading3"],
                )
            )

            story.append(Spacer(1, 6))

            story.append(
                PDFImage(
                    str(image_path),
                    width=9.8 * inch,
                    height=5.8 * inch,
                )
            )

    document.build(story)


def create_readme(path: Path, manifest_rows: list[dict]) -> None:
    copied_count = count_status(manifest_rows, "COPIED")
    skipped_count = count_status(manifest_rows, "SKIPPED_EXISTING")
    missing_count = count_status(manifest_rows, "MISSING")

    text = f"""SECOND TRIAL REPORT FIGURE PACKAGE
=================================

Created:
{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}

Purpose:
This folder contains the final report-ready figures, reports, Excel workbooks,
and CSV tables for the Second Trial.

Folder structure:
- figures: selected report-ready PNG charts
- reports: final PDF and Excel reports
- tables: final CSV summary tables
- _report_index: package manifest, PDF index, and this README

Summary:
- Copied files: {copied_count}
- Skipped existing files: {skipped_count}
- Missing files: {missing_count}

Recommended final report sources:
1. reports/01_second_trial_synthesis_report.pdf
2. reports/02_ms_treatment_visual_summary.pdf
3. reports/03_second_trial_master_summary.xlsx
4. tables/01_final_master_summary.csv
5. figures/09_final_tray_ranking.png
6. figures/10_final_group_ranking.png
7. figures/13_treatment_summary_heatmap.png

Important interpretation:
The Second Trial outputs are image-derived and descriptive. RGB green cover,
visible emergence, relative NDVI, and relative NDRE should be interpreted as
observed trial indicators. Treatment x Environment groups contain two trays
each, so the output supports descriptive comparison rather than formal
statistical proof.
"""

    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def save_settings(path: Path, overwrite: bool) -> None:
    settings = {
        "purpose": (
            "Collect final report-ready Second Trial outputs into one package."
        ),
        "output_root": str(OUTPUT_ROOT),
        "report_format_preference": "PDF",
        "overwrite": overwrite,
        "source_roots": {
            "script05_rgb_treatment_visuals": str(SCRIPT05_ROOT),
            "script08_ms_treatment_comparison": str(SCRIPT08_ROOT),
            "script09_second_trial_synthesis": str(SCRIPT09_ROOT),
        },
        "created_at": datetime.now().isoformat(timespec="seconds"),
    }

    path.parent.mkdir(parents=True, exist_ok=True)

    path.write_text(
        json.dumps(settings, indent=2),
        encoding="utf-8",
    )


# ============================================================
# 4) MAIN
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Script 10: collect final Second Trial figures, reports, "
            "Excel files, and CSV tables into one report package."
        )
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing files inside the report package.",
    )

    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be collected without copying files.",
    )

    args = parser.parse_args()

    print("\nSCRIPT 10 - SECOND TRIAL REPORT FIGURE PACKAGE")
    print("=" * 70)
    print(f"Script 05 source:\n{SCRIPT05_ROOT}")
    print(f"\nScript 08 source:\n{SCRIPT08_ROOT}")
    print(f"\nScript 09 source:\n{SCRIPT09_ROOT}")
    print(f"\nPackage output:\n{OUTPUT_ROOT}")

    all_items = [
        ("figure", FIGURE_ITEMS, FIGURES_ROOT),
        ("report", REPORT_ITEMS, REPORTS_ROOT),
        ("table", TABLE_ITEMS, TABLES_ROOT),
    ]

    planned_rows = []

    for category, items, target_folder in all_items:
        for item in items:
            source = find_existing_source(item["source_candidates"])

            planned_rows.append(
                {
                    "category": category,
                    "label": item["label"],
                    "target_name": item["target_name"],
                    "source_path": str(source) if source else "",
                    "target_path": str(target_folder / item["target_name"]),
                    "status": "FOUND" if source else "MISSING",
                    "notes": "Ready to copy." if source else "Source file was not found.",
                }
            )

    print("\nPlanned package items:")
    for row in planned_rows:
        print(
            f"{row['status']}: "
            f"{row['category']} | {row['target_name']} | {row['label']}"
        )

    if args.dry_run:
        print("\nDry run complete. No files copied.")
        return 0

    ensure_folders()

    manifest_rows = []
    figure_records = []

    for category, items, target_folder in all_items:
        for item in items:
            record = copy_item(
                item=item,
                target_folder=target_folder,
                category=category,
                overwrite=args.overwrite,
            )

            manifest_rows.append(record)

            if category == "figure":
                figure_records.append(record)

            print(
                f"{record['status']}: "
                f"{record['category']} | {record['target_name']}"
            )

    manifest_path = INDEX_ROOT / "report_package_manifest.csv"
    pdf_index_path = INDEX_ROOT / "report_package_index.pdf"
    readme_path = INDEX_ROOT / "README_report_package.txt"
    settings_path = INDEX_ROOT / "report_package_settings.json"

    write_manifest(
        manifest_path,
        manifest_rows,
    )

    create_pdf_index(
        pdf_index_path,
        manifest_rows,
        figure_records,
    )

    create_readme(
        readme_path,
        manifest_rows,
    )

    save_settings(
        settings_path,
        args.overwrite,
    )

    copied_count = count_status(manifest_rows, "COPIED")
    skipped_count = count_status(manifest_rows, "SKIPPED_EXISTING")
    missing_count = count_status(manifest_rows, "MISSING")

    print("\n" + "=" * 70)
    print("SCRIPT 10 FINISHED")
    print("=" * 70)
    print(f"COPIED: {copied_count}")
    print(f"SKIPPED_EXISTING: {skipped_count}")
    print(f"MISSING: {missing_count}")

    print(f"\nReport package:\n{OUTPUT_ROOT}")
    print(f"\nPackage index PDF:\n{pdf_index_path}")
    print(f"\nManifest CSV:\n{manifest_path}")

    if missing_count > 0:
        print(
            "\nWARNING: Some expected files were missing. "
            "Open the manifest CSV or package index PDF to see which ones."
        )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())