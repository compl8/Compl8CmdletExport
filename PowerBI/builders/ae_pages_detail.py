"""Activity Explorer report pages: Detail + drillthrough group (5xx) and
Terminology (600). Drillthrough pages declare their bound fields; the engine
emits the howCreated:5 section filters and pod bindings (explicit Back
buttons are placed in the page header).
"""

from __future__ import annotations

from . import ae_fields as f
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


def _drill_header(prefix: str, heading: str) -> list[VisualSpec]:
    """Explicit Back button + heading for drillthrough pages (suppresses the
    engine's auto back button by providing the actionButton ourselves)."""
    return [
        back_button(f"{prefix}-back", Rect(MARGIN, 16, 88, 32)),
        textbox(f"{prefix}-title", heading, Rect(136, TITLE_Y, 700, TITLE_HEIGHT)),
    ]


def _category_table(seed: str, title: str, rect: Rect,
                    workload_values: list[str] | None = None,
                    measure=None,
                    extra_filters: list[dict[str, object]] | None = None) -> VisualSpec:
    measure = measure or f.ACTIVITIES_BY_SIT
    spec = table(seed, [f.DATE, measure], rect, title=title, order_by=measure)
    if workload_values:
        spec.filters.append(
            categorical_in_filter("dim_workload", "workload", workload_values))
    for extra in extra_filters or []:
        spec.filters.append(extra)
    return spec


def activity_detail_page() -> PageSpec:
    """500: legacy 'Activity Detail' (per-channel daily SIT activity tables +
    date x rule pivot).

    The legacy channel tables filtered on endpoint activity-type values
    ('File created', 'File copied to removable media', 'File copied to
    cloud') and a 'FolderPath contains C:\\' gate. The v6 export scope is
    cloud DLP: the only activity values present are 'DLP rule matched' /
    'Copilot Interaction' / 'Classification stamped', and no local-drive
    folder paths exist, so every one of those legacy filters matched zero
    rows. The channels are rebound to the workloads that actually carry the
    data (Exchange / SharePoint / OneDrive / Teams / Copilot); removable-media
    analysis remains on 370_USB_Breakdown for endpoint-bearing exports. The
    Copilot channel counts [Raw Activities] — Copilot interactions carry no
    SIT detections, so a SIT-grain measure would always be blank there."""
    cells = grid_row(5, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="500_Activity_Detail",
        display_name="Activity Detail",
        visuals=[
            textbox("actdetail-title", "Activity Detail", title_rect()),
            *f.slicer_band("actdetail"),
            _category_table(
                "actdetail-table-email", "Sensitive Emails", cells[0],
                workload_values=["Exchange"]),
            _category_table(
                "actdetail-table-sharepoint", "Sensitive SharePoint Files", cells[1],
                workload_values=["SharePoint"]),
            _category_table(
                "actdetail-table-onedrive", "Sensitive OneDrive Files", cells[2],
                workload_values=["OneDrive"]),
            _category_table(
                "actdetail-table-teams", "Sensitive Teams Content", cells[3],
                workload_values=["MicrosoftTeams"]),
            _category_table(
                "actdetail-table-copilot", "Copilot / AI Activity", cells[4],
                workload_values=["Copilot"], measure=f.RAW_ACTIVITIES),
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
    """510: legacy 'Drill Through Activity' (drill: Happened + SITName + DLM).

    Rebound to SIT grain so the dim_sit drill filters actually reach the
    evidence rows (dim_sit only filters fact_activity_sit): the legacy
    Happened (fact_activity.happened_at) drill field becomes dim_date.date
    (a fact_activity filter cannot propagate to fact_activity_sit, and no
    source visual binds the raw timestamp), ITEM_NAME (fact_activity_detail)
    becomes dim_file.file_name, and [Total SIT Detections] mediates the
    multi-dim implicit join."""
    return PageSpec(
        folder="510_Activity_Drill",
        display_name="Drill Through Activity",
        drillthrough_fields=[
            ("dim_date", "date"),
            ("dim_sit", "sit_name"),
            ("dim_sit", "qgiscf_dlm"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *_drill_header("actdrill", "Activity Drillthrough"),
            table(
                "actdrill-table",
                [f.ACTIVITY, f.USER, f.DATE, f.FILE_NAME, f.RULE_NAME,
                 f.QGISCF_DLM, f.SIT_NAME, f.TOTAL_SIT_DETECTIONS],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Activity Evidence",
                column_widths={f.SIT_NAME: 260.0, f.FILE_NAME: 240.0,
                               f.RULE_NAME: 220.0}),
        ],
    )


def classifier_detail_page() -> PageSpec:
    """520: legacy 'Classifier Detail' (drill: SITName + User).

    ITEM_NAME (fact_activity_detail) is rebound to dim_file.file_name:
    fact_activity_detail has no relationship path to fact_activity_sit, so
    the detail column would cross-join against the SIT-grain measure;
    dim_file carries the same item/file name and relates via file_id."""
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
                 f.TOTAL_SIT_DETECTIONS, f.FILE_NAME],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Classifier Evidence",
                column_widths={f.SIT_NAME: 260.0, f.USER: 200.0,
                               f.FILE_NAME: 240.0}),
        ],
    )


def domain_drill_page() -> PageSpec:
    """530: legacy 'DomainDrillThrough' (drill: SITName/Domain/Department/User).

    SIT-grain evidence: ITEM_NAME (fact_activity_detail) is rebound to
    dim_file.file_name; TARGET_URL and SOURCE_LOCATION_TYPE are dropped —
    they exist only on fact_activity_detail, which has no unambiguous
    relationship path to fact_activity_sit, so they would cross-join against
    the SIT-grain rows (endpoint URL detail remains on 360_Device_Activity)."""
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
                 f.FOLDER_PATH, f.RULE_NAME, f.FILE_NAME],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Domain Flow Evidence",
                column_widths={f.DOMAIN: 180.0, f.SIT_NAME: 220.0,
                               f.FOLDER_PATH: 260.0, f.FILE_NAME: 220.0}),
        ],
    )


def location_drill_page() -> PageSpec:
    """540: legacy 'LocationDrillThrough' (drill: SITName/FolderPath/
    Department/User).

    This page raised Desktop's "Can't determine relationships between the
    fields" error: the legacy table mixed bare dim columns (dim_sit/dim_user/
    dim_department/dim_date) with fact_activity_detail columns and no measure
    — there is no chain of active M:1 hops covering all of them. Rebound to
    SIT grain: DEVICE_NAME and SOURCE_FILE (fact_activity_detail) are
    replaced by dim_location.folder_path (the drilled entity) and
    dim_file.file_name, with [Activities by SIT] mediating the join."""
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
                [f.SIT_NAME, f.USER, f.DEPARTMENT, f.FOLDER_PATH, f.DATE,
                 f.FILE_NAME, f.ACTIVITIES_BY_SIT],
                full_width(TALL_TABLE_Y, TALL_TABLE_HEIGHT),
                title="Location Evidence",
                column_widths={f.SIT_NAME: 260.0, f.USER: 220.0,
                               f.FOLDER_PATH: 260.0, f.FILE_NAME: 220.0}),
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
    ("Org mapping (Division / Region / Department)",
     "Division: GAL CompanyName falling back to Department ('Unknown' when the"
     " user has no GAL row) — the primary org lens in this report.\n"
     "Region: the OU directly under the Regions OU of the user's GAL"
     " distinguished name ('Unknown' for non-Regions accounts).\n"
     "Leaver / generic account: OU-derived flags (Leavers OUs; Generic"
     " Accounts/SharedUsers pools).\n"
     "Unmapped department: a user was present in Activity Explorer but no"
     " GAL/department mapping was found."),
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
