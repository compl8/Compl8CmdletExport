"""Smoke project: minimal report over the REAL star schema, proving the engine.

Generates the full semantic model (every non-pipeline/non-index SSOT table),
two measures, and two pages exercising the engine's surface: card, titled bar,
pie with Series, pivotTable, lineChart, table with columnWidth, a visual-level
measure-threshold filter, a page filter, a report-level NOT-IN filter, and
drillthrough wiring (howCreated:5 fields + auto Back button + pod binding)
to a second page that also packages the WordCloud custom visual.

Usage: py -m PowerBI.builders.build_smoke [--output-dir DIR] [--parquet-root PATH]
"""

from __future__ import annotations

import argparse
from pathlib import Path

from .expressions import col, meas
from .filters import categorical_in_filter, measure_threshold_filter, not_in_filter
from .pbi_project import ProjectSpec, build_project
from .report_layout import (
    CHART_HEIGHT,
    CHART_ROW_Y,
    PageSpec,
    TABLE_HEIGHT,
    TABLE_ROW_Y,
    full_width,
    grid_row,
    row_of_cards,
    title_rect,
)
from .tmdl_model import DEFAULT_PARQUET_ROOT, MeasureSpec
from .visual_factories import (
    bar_chart,
    card,
    line_chart,
    pie_chart,
    pivot_table,
    table,
    textbox,
    word_cloud,
)

DEFAULT_OUTPUT_DIR = Path(__file__).resolve().parents[1] / "projects" / "_smoke" / "pbix"

MEASURES = [
    MeasureSpec(
        table="fact_activity",
        name="Total Activities",
        dax="COUNTROWS ( fact_activity )",
        format_string="#,0",
        display_folder="Smoke Metrics",
        description="Distinct Activity Explorer records.",
    ),
    MeasureSpec(
        table="fact_activity_sit",
        name="Total SIT Matches",
        dax="SUM ( fact_activity_sit[match_count] )",
        format_string="#,0",
        display_folder="Smoke Metrics",
    ),
]

TOTAL_ACTIVITIES = meas("fact_activity", "Total Activities")
TOTAL_SIT_MATCHES = meas("fact_activity_sit", "Total SIT Matches")


def overview_page() -> PageSpec:
    kpi_cells = row_of_cards(4)
    chart_cells = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    table_cells = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)

    detail_table = table(
        "overview-table",
        [
            col("dim_workload", "workload", "Workload"),
            col("dim_activity_type", "activity_group", "Activity Group"),
            TOTAL_ACTIVITIES,
            TOTAL_SIT_MATCHES,
        ],
        table_cells[0],
        title="Workload Activity Summary",
        order_by=TOTAL_ACTIVITIES,
        column_widths={col("dim_workload", "workload", "Workload"): 140.0},
    )
    # Visual-level measure-threshold filter (Advanced comparison shape).
    detail_table.filters.append(measure_threshold_filter(TOTAL_ACTIVITIES, 1))

    return PageSpec(
        folder="000_Smoke_Overview",
        display_name="Smoke Overview",
        visuals=[
            textbox("overview-title", "Compl8 Builder Smoke Test", title_rect()),
            card("overview-card-activities", TOTAL_ACTIVITIES, kpi_cells[0]),
            card("overview-card-matches", TOTAL_SIT_MATCHES, kpi_cells[1]),
            bar_chart(
                "overview-bar-workload",
                col("dim_workload", "workload", "Workload"),
                [TOTAL_ACTIVITIES],
                chart_cells[0],
                title="Activities by Workload",
            ),
            pie_chart(
                "overview-pie-group",
                col("dim_activity_type", "activity_group", "Activity Group"),
                TOTAL_SIT_MATCHES,
                chart_cells[1],
                title="SIT Matches by Activity Group",
                series=col("dim_workload", "workload", "Workload"),
            ),
            line_chart(
                "overview-line-daily",
                col("dim_date", "date", "Date"),
                [TOTAL_ACTIVITIES],
                chart_cells[2],
                title="Daily Activity Trend",
            ),
            detail_table,
            pivot_table(
                "overview-pivot-risk",
                rows=[col("dim_sit", "risk_band", "Risk Band")],
                columns=[col("dim_workload", "workload", "Workload")],
                values=[TOTAL_SIT_MATCHES],
                rect=table_cells[1],
                title="Risk Band by Workload",
            ),
        ],
        page_filters=[
            categorical_in_filter(
                "dim_workload", "workload",
                ["Exchange", "SharePoint", "OneDrive", "Endpoint devices"],
            ),
        ],
    )


def drill_page() -> PageSpec:
    chart_cells = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="010_Smoke_Drill",
        display_name="Smoke Drill Detail",
        # Drillthrough target: engine emits howCreated:5 section filters,
        # the report.json pod binding, and an auto Back button.
        drillthrough_fields=[("dim_workload", "workload")],
        outspace_pane_width=362,
        visuals=[
            card("drill-card-activities", TOTAL_ACTIVITIES, row_of_cards(4)[0]),
            word_cloud(
                "drill-wordcloud-files",
                col("dim_file", "file_name", "File"),
                TOTAL_ACTIVITIES,
                chart_cells[0],
                title="File Name Cloud",
            ),
            table(
                "drill-table-evidence",
                [
                    col("dim_date", "date", "Date"),
                    col("dim_user", "user_upn", "User"),
                    col("fact_activity_detail", "item_name", "Item"),
                    col("fact_activity_detail", "application", "Application"),
                ],
                full_width(TABLE_ROW_Y, TABLE_HEIGHT),
                title="Activity Evidence",
            ),
        ],
    )


def smoke_project() -> ProjectSpec:
    return ProjectSpec(
        name="Compl8 Smoke",
        slug="compl8-smoke-v6",
        pages=[overview_page(), drill_page()],
        measures=MEASURES,
        report_filters=[
            # Advanced NOT-IN exclusion (built-in SIT names as sample values).
            not_in_filter(
                "dim_sit", "sit_name",
                [None, "Croatia Driver's License Number", "Croatia Identity Card Number"],
            ),
        ],
    )


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the smoke Power BI project.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR),
                        help="Destination PbixProj folder.")
    parser.add_argument("--parquet-root", default=DEFAULT_PARQUET_ROOT,
                        help="Default ParquetRoot parameter value.")
    parser.add_argument("--overwrite", action="store_true",
                        help="Replace an existing output folder.")
    args = parser.parse_args()

    out_dir = build_project(
        smoke_project(),
        Path(args.output_dir).resolve(),
        parquet_root=args.parquet_root,
        overwrite=args.overwrite,
    )
    print(f"Built smoke Power BI project: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
