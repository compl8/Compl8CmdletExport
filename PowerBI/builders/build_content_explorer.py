"""Content Explorer SIT Risk report: the legacy generated 15-page report
ported onto the T2 engine, against the EXISTING CE parquet layout (declared
in ce_schema; the CE star schema is a later phase).

Usage:
    py -m PowerBI.builders.build_content_explorer
        [--output-dir DIR] [--parquet-root PATH] [--overwrite]

Page plan (nav order identical to the legacy report; 052/053 renumbered from
the legacy 054/055 duplicates so folder sort == nav order):

    000_Overview              <- legacy 000_Overview 'Overview'
    005_Area_Hotspots         <- legacy 005_Area_Hotspots
    010_SIT_Risk              <- legacy 010_SIT_Risk
    020_Location_User         <- legacy 020_Location_User
    030_Area_Drilldown        <- legacy 030_Area_Drilldown
    040_File_Drillthrough     <- legacy 040_File_Drillthrough (UPGRADE: real
                                 drillthrough wiring + Back button added)
    050_Patterns              <- legacy 050_Patterns 'Patterns'
    052_Graph_Visualizer      <- legacy 054_Graph_Visualizer
    053_Node_Neighbourhood    <- legacy 055_Node_Neighbourhood
    054_Sankey_Flows          <- legacy 054_Sankey_Flows
    055_Network_Navigator     <- legacy 055_Network_Navigator
    056_Cluster_Graph         <- legacy 056_Cluster_Graph
    057_Executive_Summary     <- legacy 057_Executive_Summary
    060_Data_Quality          <- legacy 060_Data_Quality
    070_Terminology           <- legacy 070_Terminology

Upgrades over the legacy generator (content otherwise identical):
- every visual carries a curated vcObjects title (legacy had none);
- merged Compl8 theme embedded (engine default);
- layout re-expressed on the engine grid;
- 040 gets real drillthrough fields + Back button;
- dense tables get column widths;
- measure display folders; three DimFile FILTER iterators rewritten as
  column predicates (see ce_measures).
"""

from __future__ import annotations

import argparse
from pathlib import Path

from .ce_measures import MEASURES
from .ce_pages_core import core_pages
from .ce_pages_graph import graph_pages
from .ce_schema import CE_DEFAULT_PARQUET_ROOT, ce_model_source
from .pbi_project import ProjectSpec, build_project
from .report_layout import PageSpec

PROJECT_NAME = "Content Explorer SIT Risk"
PROJECT_SLUG = "content-explorer-sit-risk"  # deterministic-id namespace; keep stable

DEFAULT_OUTPUT_DIR = (
    Path(__file__).resolve().parents[1] / "projects" / "ContentExplorerSITRisk" / "pbix"
)

# Legacy section folder -> new page folder (every legacy page is ported 1:1).
LEGACY_PAGE_MAPPING: dict[str, str] = {
    "000_Overview": "000_Overview",
    "005_Area_Hotspots": "005_Area_Hotspots",
    "010_SIT_Risk": "010_SIT_Risk",
    "020_Location_User": "020_Location_User",
    "030_Area_Drilldown": "030_Area_Drilldown",
    "040_File_Drillthrough": "040_File_Drillthrough",
    "050_Patterns": "050_Patterns",
    "054_Graph_Visualizer": "052_Graph_Visualizer",
    "055_Node_Neighbourhood": "053_Node_Neighbourhood",
    "054_Sankey_Flows": "054_Sankey_Flows",
    "055_Network_Navigator": "055_Network_Navigator",
    "056_Cluster_Graph": "056_Cluster_Graph",
    "057_Executive_Summary": "057_Executive_Summary",
    "060_Data_Quality": "060_Data_Quality",
    "070_Terminology": "070_Terminology",
}


def ce_pages() -> list[PageSpec]:
    pages = [*core_pages(), *graph_pages()]
    folders = [page.folder for page in pages]
    if folders != sorted(folders):
        raise ValueError("page folders out of NNN nav order")
    return pages


def ce_project() -> ProjectSpec:
    return ProjectSpec(
        name=PROJECT_NAME,
        slug=PROJECT_SLUG,
        pages=ce_pages(),
        measures=MEASURES,
        report_filters=[],  # the legacy report declared none
        model=ce_model_source(),
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build the Content Explorer SIT Risk Power BI project.")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR),
                        help="Destination PbixProj folder.")
    parser.add_argument("--parquet-root", default=CE_DEFAULT_PARQUET_ROOT,
                        help="Default ParquetRoot parameter value.")
    parser.add_argument("--overwrite", action="store_true",
                        help="Replace an existing output folder.")
    args = parser.parse_args()

    out_dir = build_project(
        ce_project(),
        Path(args.output_dir).resolve(),
        parquet_root=args.parquet_root,
        overwrite=args.overwrite,
    )
    print(f"Built Content Explorer SIT Risk Power BI project: {out_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
