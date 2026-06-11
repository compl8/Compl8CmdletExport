"""Activity Explorer report pages: Detail + drillthrough group (5xx) and
Terminology (600). Drillthrough pages declare their bound fields; the engine
emits the howCreated:5 section filters and pod bindings (explicit Back
buttons are placed in the page header).
"""

from __future__ import annotations

from . import ae_fields as f
from .expressions import entity_column_expression, query_aliases
from .filters import categorical_in_filter, not_in_filter
from .report_layout import (
    CHART_HEIGHT,
    CHART_ROW_Y,
    MARGIN,
    PageSpec,
    TABLE_HEIGHT,
    TABLE_ROW_Y,
    TITLE_HEIGHT,
    TITLE_Y,
    full_width,
    grid_row,
    title_rect,
)
from .visual_factories import (
    Rect,
    VisualSpec,
    back_button,
    pivot_table,
    table,
    textbox,
)

DRILL_PANE_WIDTH = 362
TALL_TABLE_Y = 64
TALL_TABLE_HEIGHT = 720 - TALL_TABLE_Y - 40  # 616


def contains_filter(table_name: str, column: str, substring: str) -> dict[str, object]:
    """Advanced Contains filter (legacy 'FolderPath contains C:\\' shape)."""
    alias = query_aliases([table_name])[table_name]
    escaped = substring.replace("'", "''")
    return {
        "expression": entity_column_expression(table_name, column),
        "filter": {
            "Version": 2,
            "From": [{"Name": alias, "Entity": table_name, "Type": 0}],
            "Where": [
                {
                    "Condition": {
                        "Contains": {
                            "Left": {
                                "Column": {
                                    "Expression": {"SourceRef": {"Source": alias}},
                                    "Property": column,
                                }
                            },
                            "Right": {"Literal": {"Value": f"'{escaped}'"}},
                        }
                    }
                }
            ],
        },
        "type": "Advanced",
        "howCreated": 1,
    }


def _drill_header(prefix: str, heading: str) -> list[VisualSpec]:
    """Explicit Back button + heading for drillthrough pages (suppresses the
    engine's auto back button by providing the actionButton ourselves)."""
    return [
        back_button(f"{prefix}-back", Rect(MARGIN, 16, 88, 32)),
        textbox(f"{prefix}-title", heading, Rect(136, TITLE_Y, 700, TITLE_HEIGHT)),
    ]


def _category_table(seed: str, title: str, rect: Rect,
                    activity_values: list[str] | None = None,
                    extra_filters: list[dict[str, object]] | None = None) -> VisualSpec:
    spec = table(seed, [f.DATE, f.ACTIVITIES_BY_SIT], rect, title=title,
                 order_by=f.ACTIVITIES_BY_SIT)
    if activity_values:
        spec.filters.append(
            categorical_in_filter("dim_activity_type", "activity", activity_values))
    for extra in extra_filters or []:
        spec.filters.append(extra)
    return spec


def activity_detail_page() -> PageSpec:
    """500: legacy 'Activity Detail' (per-category daily SIT activity tables +
    date x rule pivot). The fifth legacy table's 'FolderPath contains C:\\'
    gate is ported via an explicit Contains filter."""
    cells = grid_row(5, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="500_Activity_Detail",
        display_name="Activity Detail",
        visuals=[
            textbox("actdetail-title", "Activity Detail", title_rect()),
            *f.slicer_band("actdetail"),
            _category_table(
                "actdetail-table-email", "Sensitive Emails", cells[0],
                extra_filters=[categorical_in_filter("dim_workload", "workload",
                                                     ["Exchange"])]),
            _category_table(
                "actdetail-table-created", "Sensitive Files Created", cells[1],
                activity_values=["File created", "File created on network share",
                                 "File created on removable media"]),
            _category_table(
                "actdetail-table-usb", "Sensitive USB", cells[2],
                activity_values=["File copied to removable media",
                                 "File created on removable media"]),
            _category_table(
                "actdetail-table-cloud", "Sensitive Cloud", cells[3],
                activity_values=["File copied to cloud"]),
            _category_table(
                "actdetail-table-local", "Sensitive Files on Local Drives", cells[4],
                extra_filters=[contains_filter("dim_location", "folder_path", "C:\\")]),
            _rule_pivot(),
        ],
    )


def _rule_pivot() -> VisualSpec:
    pivot = pivot_table(
        "actdetail-pivot-rule", rows=[f.DATE], columns=[f.RULE_NAME],
        values=[f.ACTIVITIES_BY_SIT],
        rect=full_width(TABLE_ROW_Y, TABLE_HEIGHT),
        title="SIT Activities by Day and DLP Rule")
    # Legacy intent: hide the no-rule column (tenant-specific rule exclusions
    # from the old report are NOT ported).
    pivot.filters.append(not_in_filter("dim_policy", "rule_name", [None]))
    return pivot


def activity_drill_page() -> PageSpec:
    """510: legacy 'Drill Through Activity' (drill: Happened + SITName + DLM)."""
    return PageSpec(
        folder="510_Activity_Drill",
        display_name="Drill Through Activity",
        drillthrough_fields=[
            ("fact_activity", "happened_at"),
            ("dim_sit", "sit_name"),
            ("dim_sit", "qgiscf_dlm"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *_drill_header("actdrill", "Activity Drillthrough"),
            table(
                "actdrill-table",
                [f.ACTIVITY, f.USER, f.HAPPENED_AT, f.ITEM_NAME, f.RULE_NAME,
                 f.QGISCF_DLM, f.SIT_NAME],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Activity Evidence",
                column_widths={f.SIT_NAME: 260.0, f.ITEM_NAME: 240.0,
                               f.RULE_NAME: 220.0}),
        ],
    )


def classifier_detail_page() -> PageSpec:
    """520: legacy 'Classifier Detail' (drill: SITName + User)."""
    return PageSpec(
        folder="520_Classifier_Detail",
        display_name="Classifier Detail",
        drillthrough_fields=[
            ("dim_sit", "sit_name"),
            ("dim_user", "user_upn"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *_drill_header("clsdetail", "Classifier Drillthrough"),
            table(
                "clsdetail-table",
                [f.SIT_NAME, f.QGISCF_DLM, f.ACTIVITY, f.FILE_TYPE, f.USER,
                 f.TOTAL_SIT_DETECTIONS, f.ITEM_NAME],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Classifier Evidence",
                column_widths={f.SIT_NAME: 260.0, f.USER: 200.0,
                               f.ITEM_NAME: 240.0}),
        ],
    )


def domain_drill_page() -> PageSpec:
    """530: legacy 'DomainDrillThrough' (drill: SITName/Domain/Department/User)."""
    return PageSpec(
        folder="530_Domain_Drill",
        display_name="Domain Drillthrough",
        drillthrough_fields=[
            ("dim_sit", "sit_name"),
            ("dim_domain", "domain"),
            ("dim_department", "department"),
            ("dim_user", "user_upn"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *_drill_header("domdrill", "Domain Drillthrough"),
            table(
                "domdrill-table",
                [f.DOMAIN, f.DEPARTMENT, f.USER, f.SIT_NAME, f.ACTIVITIES_BY_SIT,
                 f.FOLDER_PATH, f.RULE_NAME, f.TARGET_URL,
                 f.SOURCE_LOCATION_TYPE, f.ITEM_NAME],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Domain Flow Evidence",
                column_widths={f.DOMAIN: 180.0, f.SIT_NAME: 220.0,
                               f.FOLDER_PATH: 260.0, f.TARGET_URL: 220.0}),
        ],
    )


def location_drill_page() -> PageSpec:
    """540: legacy 'LocationDrillThrough' (drill: SITName/FolderPath/
    Department/User)."""
    return PageSpec(
        folder="540_Location_Drill",
        display_name="Location Drillthrough",
        drillthrough_fields=[
            ("dim_sit", "sit_name"),
            ("dim_location", "folder_path"),
            ("dim_department", "department"),
            ("dim_user", "user_upn"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *_drill_header("locdrill", "Location Drillthrough"),
            table(
                "locdrill-table",
                [f.SIT_NAME, f.USER, f.DEPARTMENT, f.DEVICE_NAME, f.DATE,
                 f.SOURCE_FILE],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Location Evidence",
                column_widths={f.SIT_NAME: 260.0, f.USER: 220.0,
                               f.SOURCE_FILE: 240.0}),
        ],
    )


def drill_summary_page() -> PageSpec:
    """550: legacy 'Drill Through Summary' (SIT x department launch table —
    not itself a drill target, mirrors the legacy page)."""
    return PageSpec(
        folder="550_Drill_Summary",
        display_name="Drill Through Summary",
        visuals=[
            textbox("drillsum-title", "Drill Through Summary", title_rect()),
            *f.slicer_band("drillsum"),
            table(
                "drillsum-table",
                [f.SIT_NAME, f.QGISCF_DLM, f.ACTIVITIES_BY_SIT, f.DEPARTMENT],
                full_width(CHART_ROW_Y, 720 - CHART_ROW_Y - 40),
                title="SIT and Department Summary (right-click to drill)",
                order_by=f.ACTIVITIES_BY_SIT,
                column_widths={f.SIT_NAME: 320.0}),
        ],
    )


_TERMINOLOGY = [
    ("Activity grain",
     "Activity: one raw Activity Explorer record exported from Microsoft Purview.\n"
     "SIT Activity Event: one activity matched to one SIT. An activity with three"
     " SITs contributes three SIT activity events.\n"
     "SIT Match: the count reported inside the activity's SensitiveInfoTypeData payload."),
    ("Risk",
     "Risk score: numeric value loaded from the SIT risk workbook.\n"
     "Risk band: Low, Medium, High, Critical, or Unrated.\n"
     "Risk pressure: match count multiplied by risk score and then summed —"
     " it intentionally combines severity and volume."),
    ("Unrated SITs and classifications",
     "Unrated SIT: a SIT detected in Activity Explorer that was not mapped to the"
     " supplied risk workbook. It is kept visible because treating unknown mappings"
     " as low risk hides register gaps.\n"
     "QGISCF DLM and PSPF: classification fields imported from the SIT risk workbook."),
    ("Departments",
     "Department mapping: user enrichment loaded from the GAL/department mapping"
     " supplied to the conversion script.\n"
     "Unmapped department: a user was present in Activity Explorer but no"
     " department mapping was found."),
    ("Locations",
     "Location: folder path derived from the activity FilePath field.\n"
     "Location hotspot: a folder path with elevated SIT matches, high-confidence"
     " matches, or risk pressure.\n"
     "Path depth: number of path segments in the derived folder path."),
    ("Aggregates vs detail",
     "Overview pages use the daily aggregate tables (department, user, location,"
     " activity type x SIT x day) for speed.\n"
     "Detail and drillthrough pages use the raw fact tables and should be used"
     " after filtering, not as the primary exploration surface for a very large export."),
]


def terminology_page() -> PageSpec:
    """600: Terminology (adopted from the interim 9-page report)."""
    visuals: list[VisualSpec] = [
        textbox("terms-title", "Terminology", title_rect()),
    ]
    rows = [grid_row(2, 64 + index * 212, 200) for index in range(3)]
    cells = [cell for row in rows for cell in row]
    for index, (heading, body) in enumerate(_TERMINOLOGY):
        visuals.append(textbox(
            f"terms-text-{index}", f"{heading}\n{body}", cells[index],
            font_size=11, bold=False))
    return PageSpec(
        folder="600_Terminology",
        display_name="Terminology",
        visuals=visuals,
    )


def detail_pages() -> list[PageSpec]:
    return [
        activity_detail_page(),
        activity_drill_page(),
        classifier_detail_page(),
        domain_drill_page(),
        location_drill_page(),
        drill_summary_page(),
        terminology_page(),
    ]
