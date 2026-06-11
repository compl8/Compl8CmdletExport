"""Activity Explorer Risk report: superset of the legacy 29-page report,
generated on the T2 engine against the star schema v6.

Usage:
    py -m PowerBI.builders.build_activity_explorer
        [--output-dir DIR] [--parquet-root PATH] [--overwrite]

Page plan (NNN nav order) and legacy -> new mapping
---------------------------------------------------

Overview group
    000_Executive_Overview   <- legacy 'Executive Overview' merged with the
                                interim report's Executive page (agg visuals)
    010_Activity_Summary     <- legacy 'Activity Summary Table' + 'Summary
                                Activity Detail' (the DLM x Workload pivot
                                appeared on both; merged)
    020_Timeline             <- legacy 'Timeline'
Risk / SIT analysis group
    100_Risk_Assessment      <- legacy 'Risk Assessment'
    110_Classifier_Analysis  <- legacy 'Classifier Analysis'
    120_Classifier_Focus     <- legacy 'Classifier Focus'
    130_File_Analysis        <- legacy 'File Analysis'
People group
    200_Department_Analysis  <- legacy 'Department Analysis'
    210_Department_Treemap   <- legacy 'TreeDept'
    220_User_Investigation   <- legacy 'User'
Locations / movement group
    300_Location_Hotspots    <- legacy 'Location'
    310_Location_Risk        <- legacy 'Location Risk'
    320_Domain_Data_Flows    <- legacy 'Domain Data Flows'
    330_Location_Domain_Flows<- legacy 'Location Domain Data Flows'
    340_Folder_Data_Flows    <- legacy 'Folder Data Flows'
    350_Domain_Graph         <- legacy 'Graph Domain Data Flows' (fallback
                                matrix kept + ForceGraph upgrade added)
    360_Device_Activity      <- legacy 'Device'
    370_USB_Breakdown        <- legacy 'USB Breakdown'
Email / policy / AI group
    400_DLP_Policy_Analysis  <- legacy 'DLP Policy Analysis'
    410_Email_Subject_Cloud  <- legacy 'Subject Heading Word Cloud'
    420_AI_View              <- legacy 'AI View'
    430_Agent_Activity       <- legacy 'Agent Activity' (upgraded onto
                                fact_copilot_interaction + dim_app_identity)
Detail + drillthrough group
    500_Activity_Detail      <- legacy 'Activity Detail'
    510_Activity_Drill       <- legacy 'Drill Through Activity'
    520_Classifier_Detail    <- legacy 'Classifier Detail'
    530_Domain_Drill         <- legacy 'DomainDrillThrough'
    540_Location_Drill       <- legacy 'LocationDrillThrough'
    550_Drill_Summary        <- legacy 'Drill Through Summary'
Terminology
    600_Terminology          <- interim report's Terminology page

Deviations from the legacy report (intentional):
- The 269-value SITName NOT-IN report filter is applied at ETL in v6 and is
  NOT replicated in-report. No other legacy report-level filters existed.
- Tenant-specific visual-filter value lists (department/domain/rule-name
  exclusions hard-coded in the old report) are not ported; the structural
  gates (TotalRisk > 100, detection thresholds, null-rule exclusion,
  removable-media activity sets) are.
- Legacy time-intelligence measures anchored to TODAY() are re-anchored to
  the latest data date (MAX(dim_date[date])).
"""

from __future__ import annotations

import argparse
from pathlib import Path

from .ae_measures import MEASURES
from .ae_pages_detail import detail_pages
from .ae_pages_flows import flows_pages
from .ae_pages_overview import overview_pages
from .pbi_project import ProjectSpec, build_project
from .report_layout import PageSpec
from .tmdl_model import DEFAULT_PARQUET_ROOT

PROJECT_NAME = "Activity Explorer Risk"
PROJECT_SLUG = "activity-explorer"  # deterministic-id namespace; keep stable

DEFAULT_OUTPUT_DIR = (
    Path(__file__).resolve().parents[1] / "projects" / "ActivityExplorer" / "pbix"
)

# Legacy 29-page report -> new page folder (superset contract; every legacy
# page's content exists on its mapped page). Tests assert this mapping is
# total and that every target page is emitted.
LEGACY_PAGE_MAPPING: dict[str, str] = {
    "Domain Data Flows": "320_Domain_Data_Flows",
    "Department Analysis": "200_Department_Analysis",
    "Location": "300_Location_Hotspots",
    "Location Risk": "310_Location_Risk",
    "Timeline": "020_Timeline",
    "Risk Assessment": "100_Risk_Assessment",
    "Location Domain Data Flows": "330_Location_Domain_Flows",
    "File Analysis": "130_File_Analysis",
    "DLP Policy Analysis": "400_DLP_Policy_Analysis",
    "AI View": "420_AI_View",
    "Graph Domain Data Flows": "350_Domain_Graph",
    "Subject Heading Word Cloud": "410_Email_Subject_Cloud",
    "Classifier Focus": "120_Classifier_Focus",
    "Agent Activity": "430_Agent_Activity",
    "Executive Overview": "000_Executive_Overview",
    "Device": "360_Device_Activity",
    "Folder Data Flows": "340_Folder_Data_Flows",
    "Activity Summary Table": "010_Activity_Summary",
    "Classifier Analysis": "110_Classifier_Analysis",
    "TreeDept": "210_Department_Treemap",
    "Classifier Detail": "520_Classifier_Detail",
    "Drill Through Activity": "510_Activity_Drill",
    "Drill Through Summary": "550_Drill_Summary",
    "DomainDrillThrough": "530_Domain_Drill",
    "LocationDrillThrough": "540_Location_Drill",
    "Activity Detail": "500_Activity_Detail",
    "Summary Activity Detail": "010_Activity_Summary",
    "USB Breakdown": "370_USB_Breakdown",
    "User": "220_User_Investigation",
}


def ae_pages() -> list[PageSpec]:
    pages = [*overview_pages(), *flows_pages(), *detail_pages()]
    folders = [page.folder for page in pages]
    if folders != sorted(folders):
        raise ValueError("page folders out of NNN nav order")
    return pages


def ae_project() -> ProjectSpec:
    return ProjectSpec(
        name=PROJECT_NAME,
        slug=PROJECT_SLUG,
        pages=ae_pages(),
        measures=MEASURES,
        # The legacy report's only report-level filter (269-value SITName
        # NOT-IN) is applied at ETL in v6 — nothing to declare here.
        report_filters=[],
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build the Activity Explorer Risk Power BI project.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR),
                        help="Destination PbixProj folder.")
    parser.add_argument("--parquet-root", default=DEFAULT_PARQUET_ROOT,
                        help="Default ParquetRoot parameter value.")
    parser.add_argument("--overwrite", action="store_true",
                        help="Replace an existing output folder.")
    args = parser.parse_args()

    out_dir = build_project(
        ae_project(),
        Path(args.output_dir).resolve(),
        parquet_root=args.parquet_root,
        overwrite=args.overwrite,
    )
    print(f"Built Activity Explorer Power BI project: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
