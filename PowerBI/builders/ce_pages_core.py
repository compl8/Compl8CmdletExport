"""Content Explorer SIT Risk pages 000-050 (overview / risk / drilldown group)
plus the field shorthands and layout helpers shared with ce_pages_graph.

Every visual mirrors a visual in the legacy generated report (same bindings,
display names, and sort) and gains a curated vcObjects title; layout is
re-expressed on the engine grid. The 040 File Drillthrough page gains real
drillthrough wiring (the legacy page had none).
"""

from __future__ import annotations

from .expressions import Field, col, meas
from .report_layout import (
    CHART_HEIGHT,
    CHART_ROW_Y,
    CONTENT_WIDTH,
    GUTTER,
    KPI_HEIGHT,
    KPI_ROW_Y,
    MARGIN,
    PageSpec,
    SLICER_HEIGHT,
    SLICER_ROW_Y,
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
    bar_chart,
    card,
    scatter_chart,
    slicer,
    table,
    textbox,
    treemap,
)

DRILL_PANE_WIDTH = 362
TALL_HEIGHT = 720 - CHART_ROW_Y - 40  # hero band replacing chart+table rows
KPI_BAND_WIDTH = 896  # KPI cards stop short of a 2-slicer right band
KPI_BAND_NARROW = 768  # for pages with a 3-slicer right band
SLICER_WIDTH = 140


def m(name: str, display_name: str | None = None) -> Field:
    """Measure reference (all CE measures live on FactLocationSIT)."""
    return meas("FactLocationSIT", name, display_name)


def kpi_cells(count: int, total_width: float = KPI_BAND_WIDTH) -> list[Rect]:
    return grid_row(count, KPI_ROW_Y, KPI_HEIGHT, total_width=total_width)


def slicer_band(prefix: str, fields: list[Field]) -> list[VisualSpec]:
    """Right-aligned slicer band (slicers sit beside the KPI row)."""
    count = len(fields)
    x0 = MARGIN + CONTENT_WIDTH - count * SLICER_WIDTH - (count - 1) * GUTTER
    cells = grid_row(count, SLICER_ROW_Y, SLICER_HEIGHT,
                     x0=x0, total_width=count * SLICER_WIDTH + (count - 1) * GUTTER)
    return [
        slicer(f"{prefix}-slicer-{field.name}", field, cell, title=field.shown_as())
        for field, cell in zip(fields, cells)
    ]


def split_row(weights: tuple[int, ...], y: float, height: float) -> list[Rect]:
    """Split the content band into weighted cells (e.g. (1, 2) = third/two-thirds)."""
    total = sum(weights)
    available = CONTENT_WIDTH - GUTTER * (len(weights) - 1)
    cells: list[Rect] = []
    x = float(MARGIN)
    for weight in weights:
        width = available * weight / total
        cells.append(Rect(x, y, width, height))
        x += width + GUTTER
    return cells


def drill_header(prefix: str, heading: str) -> list[VisualSpec]:
    """Explicit Back button + heading for drillthrough pages."""
    return [
        back_button(f"{prefix}-back", Rect(MARGIN, 16, 88, 32)),
        textbox(f"{prefix}-title", heading, Rect(136, TITLE_Y, 700, TITLE_HEIGHT)),
    ]


# --- frequently used column references ----------------------------------------

SIT_NAME = col("DimSIT", "sit_name", "SIT")
SIT_CATEGORY = col("DimSIT", "category", "Category")
RISK_BAND = col("DimSIT", "risk_band", "Risk Band")
RISK_SCORE = col("DimSIT", "risk_score", "Risk Score")
QGISCF_DLM = col("DimSIT", "qgiscf_dlm", "QGISCF DLM")
PSPF = col("DimSIT", "pspf_classification", "PSPF")
LOCATION_NAME = col("DimLocation", "location_name", "Location")
LOCATION_WORKLOAD = col("DimLocation", "workload", "Workload")
AREA_PATH = col("DimArea", "area_display_path", "Area")
AREA_LOCATION = col("DimArea", "location_name", "Location")
AREA_WORKLOAD = col("DimArea", "workload", "Workload")
FOLDER_DEPTH = col("DimArea", "folder_depth", "Depth")
FILE_NAME = col("DimFile", "file_name", "File")
FILE_EXTENSION = col("DimFile", "file_extension", "Extension")
SENSITIVITY_LABEL = col("DimFile", "sensitivity_label", "Sensitivity Label")
RETENTION_LABEL = col("DimFile", "retention_label", "Retention Label")
LAST_MODIFIED = col("DimFile", "last_modified_date", "Last Modified Date")
TAG_SIT_NAME = col("FactFileSITTag", "sit_name", "SIT")
TAG_HIGH_CONF = col("FactFileSITTag", "tag_high_confidence_count", "High Confidence Count")
TAG_TOTAL = col("FactFileSITTag", "tag_total_count", "Total Count")
USER_NAME = col("DimUser", "user_display_name", "User")


def overview_page() -> PageSpec:
    """000: tenant-wide aggregate exposure overview."""
    kpis = kpi_cells(4)
    charts = split_row((1, 2), CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="000_Overview", display_name="Overview",
        visuals=[
            textbox("overview-title", "Content Explorer SIT Risk Overview", title_rect()),
            card("overview-card-items", m("Aggregate Items"), kpis[0], title="Aggregate Items"),
            card("overview-card-files", m("Files With Exported SIT", "Files With SIT"),
                 kpis[1], title="Files With SIT"),
            card("overview-card-highcrit",
                 m("High Or Critical Aggregate Items", "High/Critical Items"),
                 kpis[2], title="High/Critical Items"),
            card("overview-card-avgrisk", m("Weighted Average Risk", "Avg Risk"),
                 kpis[3], title="Average Risk Score"),
            *slicer_band("overview", [LOCATION_WORKLOAD, RISK_BAND]),
            bar_chart("overview-bar-band", RISK_BAND, [m("Aggregate Items")], charts[0],
                      title="Aggregate Items by Risk Band"),
            bar_chart("overview-bar-location", LOCATION_NAME, [m("Aggregate Items")],
                      charts[1], title="Top Locations by Aggregate Items"),
            table("overview-table-sit",
                  [SIT_NAME, SIT_CATEGORY, RISK_BAND, QGISCF_DLM, m("Aggregate Items"),
                   m("Files With Exported SIT"), m("Weighted Average Risk")],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="SIT Exposure Summary",
                  order_by=m("Aggregate Items"),
                  column_widths={SIT_NAME: 280.0, SIT_CATEGORY: 160.0}),
        ],
    )


def area_hotspots_page() -> PageSpec:
    """005: area (leaf folder) risk hotspots."""
    kpis = kpi_cells(5)
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="005_Area_Hotspots", display_name="Area Hotspots",
        visuals=[
            textbox("areahot-title", "Area Hotspots", title_rect()),
            card("areahot-card-files", m("Files With Exported SIT", "Files With SIT"),
                 kpis[0], title="Files With SIT"),
            card("areahot-card-pairs", m("Area File SIT Pairs", "File/SIT Pairs"),
                 kpis[1], title="File/SIT Pairs"),
            card("areahot-card-density", m("Area Risk Density", "Risk / File"),
                 kpis[2], title="Risk per File"),
            card("areahot-card-sitsfile", m("Area SITs Per File", "SITs / File"),
                 kpis[3], title="SITs per File"),
            card("areahot-card-highconf",
                 m("Area High Confidence Matches", "High Conf Matches"),
                 kpis[4], title="High Confidence Matches"),
            *slicer_band("areahot", [AREA_WORKLOAD, RISK_BAND]),
            treemap("areahot-treemap-location", AREA_LOCATION,
                    m("Area Risk Pressure", "Risk Pressure"), charts[0],
                    title="Risk Pressure by Location"),
            bar_chart("areahot-bar-area", AREA_PATH,
                      [m("Area Risk Density", "Risk / File")], charts[1],
                      title="Riskiest Areas (Risk per File)"),
            scatter_chart("areahot-scatter",
                          m("Files With Exported SIT", "Files With SIT"),
                          m("Area Risk Density", "Risk / File"), AREA_PATH, charts[2],
                          size=m("Area File SIT Pairs", "File/SIT Pairs"),
                          title="Area Volume vs Risk Density"),
            table("areahot-table-area",
                  [AREA_PATH, FOLDER_DEPTH, m("Files With Exported SIT", "Files With SIT"),
                   m("Area File SIT Pairs", "File/SIT Pairs"),
                   m("Area Distinct SITs", "Distinct SITs"),
                   m("Area Risk Pressure", "Risk Pressure"),
                   m("Area Risk Density", "Risk / File"),
                   m("Area SITs Per File", "SITs / File"),
                   m("High Or Critical Area File SIT Pairs", "High/Critical Pairs"),
                   m("Area Sensitivity Label Coverage %", "Label Coverage")],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="Area Hotspot Detail",
                  order_by=m("Area Risk Density"),
                  column_widths={AREA_PATH: 300.0}),
        ],
    )


def sit_risk_page() -> PageSpec:
    """010: SIT-centric risk profile."""
    kpis = kpi_cells(4)
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="010_SIT_Risk", display_name="SIT Risk",
        visuals=[
            textbox("sitrisk-title", "SIT Risk", title_rect()),
            card("sitrisk-card-critical", m("Critical Aggregate Items", "Critical Items"),
                 kpis[0], title="Critical Items"),
            card("sitrisk-card-files",
                 m("High Or Critical Detected Files", "High/Critical Files"),
                 kpis[1], title="High/Critical Files"),
            card("sitrisk-card-unrated", m("Unrated SITs"), kpis[2], title="Unrated SITs"),
            card("sitrisk-card-unlabelled",
                 m("Critical Unlabelled Files", "Critical Unlabelled"),
                 kpis[3], title="Critical Unlabelled Files"),
            *slicer_band("sitrisk", [SIT_CATEGORY, QGISCF_DLM]),
            bar_chart("sitrisk-bar-sit", SIT_NAME, [m("High Confidence Matches")],
                      charts[0], title="SITs by High Confidence Matches"),
            bar_chart("sitrisk-bar-dlm", QGISCF_DLM, [m("Aggregate Items")], charts[1],
                      title="Aggregate Items by Classification (DLM)"),
            table("sitrisk-table-sit",
                  [SIT_NAME, SIT_CATEGORY, RISK_SCORE, RISK_BAND, PSPF, QGISCF_DLM,
                   m("Aggregate Items"), m("Detected Files By Location"),
                   m("High Confidence Matches")],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="SIT Risk Detail",
                  order_by=m("Aggregate Items"),
                  column_widths={SIT_NAME: 280.0}),
        ],
    )


def location_user_page() -> PageSpec:
    """020: location and user exposure."""
    kpis = kpi_cells(4)
    charts = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    tables = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="020_Location_User", display_name="Location And User",
        visuals=[
            textbox("locuser-title", "Location And User Exposure", title_rect()),
            card("locuser-card-locations", m("Locations"), kpis[0], title="Locations"),
            card("locuser-card-users", m("Users On Files", "Users"), kpis[1],
                 title="Users on Files"),
            card("locuser-card-files", m("User Files"), kpis[2], title="User Files"),
            card("locuser-card-tags", m("User File SIT Tags", "User SIT Tags"), kpis[3],
                 title="User SIT Tags"),
            *slicer_band("locuser", [LOCATION_WORKLOAD,
                                     col("DimLocation", "location_type", "Location Type")]),
            bar_chart("locuser-bar-location", LOCATION_NAME, [m("Aggregate Items")],
                      charts[0], title="Locations by Aggregate Items"),
            bar_chart("locuser-bar-user", USER_NAME, [m("User Files")], charts[1],
                      title="Users by Files Touched"),
            table("locuser-table-location",
                  [LOCATION_NAME, LOCATION_WORKLOAD,
                   col("DimLocation", "location_type", "Location Type"),
                   col("DimLocation", "owner_candidate", "Owner Candidate"),
                   m("Aggregate Items"), m("High Or Critical Aggregate Items"),
                   m("Files With Exported SIT")],
                  tables[0], title="Location Exposure", order_by=m("Aggregate Items"),
                  column_widths={LOCATION_NAME: 200.0}),
            table("locuser-table-user",
                  [USER_NAME, col("BridgeFileUser", "user_role", "Role"),
                   m("User Files"), m("User File SIT Tags")],
                  tables[1], title="User Exposure", order_by=m("User Files")),
        ],
    )


def area_drilldown_page() -> PageSpec:
    """030: area drilldown to SIT composition + file evidence."""
    kpis = kpi_cells(4, KPI_BAND_NARROW)
    halves = grid_row(2, CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="030_Area_Drilldown", display_name="Area Drilldown",
        visuals=[
            textbox("areadrill-title", "Area Drilldown", title_rect()),
            card("areadrill-card-files", m("Files With Exported SIT", "Files With SIT"),
                 kpis[0], title="Files With SIT"),
            card("areadrill-card-pairs", m("Area File SIT Pairs", "File/SIT Pairs"),
                 kpis[1], title="File/SIT Pairs"),
            card("areadrill-card-sits", m("Area Distinct SITs", "Distinct SITs"),
                 kpis[2], title="Distinct SITs"),
            card("areadrill-card-pressure", m("Area Risk Pressure", "Risk Pressure"),
                 kpis[3], title="Risk Pressure"),
            *slicer_band("areadrill",
                         [AREA_LOCATION, col("DimArea", "area_level_1", "Level 1"),
                          col("DimArea", "area_level_2", "Level 2")]),
            table("areadrill-table-area",
                  [AREA_PATH, FOLDER_DEPTH,
                   m("Files With Exported SIT", "Files With SIT"),
                   m("Area File SIT Pairs", "File/SIT Pairs"),
                   m("Area Risk Density", "Risk / File"),
                   m("Area SITs Per File", "SITs / File"),
                   m("Area High Confidence Matches", "High Conf Matches")],
                  halves[0], title="Areas (Risk per File)",
                  order_by=m("Area Risk Density"), column_widths={AREA_PATH: 240.0}),
            table("areadrill-table-sit",
                  [SIT_NAME, RISK_BAND, SIT_CATEGORY,
                   m("Area File SIT Pairs", "File/SIT Pairs"),
                   m("Area Total Matches", "Matches"),
                   m("Area High Confidence Matches", "High Conf Matches"),
                   m("Area Risk Pressure", "Risk Pressure")],
                  halves[1], title="SIT Composition of Selected Area",
                  order_by=m("Area File SIT Pairs"), column_widths={SIT_NAME: 220.0}),
            table("areadrill-table-files",
                  [FILE_NAME, col("DimFile", "folder_path", "Folder"), TAG_SIT_NAME,
                   RISK_BAND, FILE_EXTENSION, SENSITIVITY_LABEL, TAG_HIGH_CONF,
                   TAG_TOTAL, LAST_MODIFIED],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="File Evidence",
                  column_widths={FILE_NAME: 220.0,
                                 col("DimFile", "folder_path", "Folder"): 260.0,
                                 TAG_SIT_NAME: 200.0}),
        ],
    )


def file_drillthrough_page() -> PageSpec:
    """040: file drillthrough — UPGRADE: real drillthrough wiring (the legacy
    page declared none). Drill fields: file / SIT / location identifiers and
    file extension, all bound on this page's evidence visuals."""
    kpis = kpi_cells(4)
    charts = split_row((1, 2), CHART_ROW_Y, CHART_HEIGHT)
    return PageSpec(
        folder="040_File_Drillthrough", display_name="File Drillthrough",
        drillthrough_fields=[
            ("DimFile", "file_name"),
            ("DimSIT", "sit_name"),
            ("DimLocation", "location_name"),
            ("DimFile", "file_extension"),
        ],
        outspace_pane_width=DRILL_PANE_WIDTH,
        visuals=[
            *drill_header("filedrill", "File Drillthrough"),
            card("filedrill-card-files", m("Files With Exported SIT", "Files With SIT"),
                 kpis[0], title="Files With SIT"),
            card("filedrill-card-rows", m("File SIT Tag Rows", "SIT Tag Rows"), kpis[1],
                 title="SIT Tag Rows"),
            card("filedrill-card-unique", m("Unique Files"), kpis[2], title="Unique Files"),
            card("filedrill-card-highcrit",
                 m("High Or Critical Detected Files", "High/Critical Files"),
                 kpis[3], title="High/Critical Files"),
            *slicer_band("filedrill", [RISK_BAND, FILE_EXTENSION]),
            bar_chart("filedrill-bar-ext", FILE_EXTENSION,
                      [m("Files With Exported SIT")], charts[0],
                      title="Files With SIT by Extension"),
            bar_chart("filedrill-bar-label", SENSITIVITY_LABEL,
                      [m("Files With Exported SIT")], charts[1],
                      title="Files With SIT by Sensitivity Label"),
            table("filedrill-table-files",
                  [FILE_NAME, LOCATION_NAME, TAG_SIT_NAME, RISK_BAND, FILE_EXTENSION,
                   SENSITIVITY_LABEL, RETENTION_LABEL, TAG_HIGH_CONF, TAG_TOTAL,
                   LAST_MODIFIED],
                  full_width(TABLE_ROW_Y, TABLE_HEIGHT), title="File Evidence",
                  column_widths={FILE_NAME: 220.0, LOCATION_NAME: 160.0,
                                 TAG_SIT_NAME: 200.0, SENSITIVITY_LABEL: 140.0}),
        ],
    )


def patterns_page() -> PageSpec:
    """050: pattern finder (category / extension / depth mixes)."""
    kpis = kpi_cells(4)
    charts = grid_row(3, CHART_ROW_Y, CHART_HEIGHT)
    bottoms = grid_row(2, TABLE_ROW_Y, TABLE_HEIGHT)
    return PageSpec(
        folder="050_Patterns", display_name="Patterns",
        visuals=[
            textbox("patterns-title", "Pattern Finder", title_rect()),
            card("patterns-card-areas", m("Areas"), kpis[0], title="Areas"),
            card("patterns-card-pressure", m("Area Risk Pressure", "Risk Pressure"),
                 kpis[1], title="Risk Pressure"),
            card("patterns-card-density", m("Area Risk Density", "Risk / File"),
                 kpis[2], title="Risk per File"),
            card("patterns-card-coverage",
                 m("Area Sensitivity Label Coverage %", "Area Label Coverage"),
                 kpis[3], title="Area Label Coverage"),
            *slicer_band("patterns", [AREA_WORKLOAD, RISK_BAND]),
            treemap("patterns-treemap-category",
                    col("DimSIT", "category", "SIT Category"),
                    m("Area Risk Pressure", "Risk Pressure"), charts[0],
                    title="Risk Pressure by SIT Category"),
            treemap("patterns-treemap-ext", FILE_EXTENSION,
                    m("Files With Exported SIT", "Files With SIT"), charts[1],
                    title="Files With SIT by Extension"),
            bar_chart("patterns-bar-depth",
                      col("DimArea", "folder_depth", "Folder Depth"),
                      [m("Area File SIT Pairs", "File/SIT Pairs")], charts[2],
                      title="File/SIT Pairs by Folder Depth"),
            scatter_chart("patterns-scatter",
                          m("Files With Exported SIT", "Files With SIT"),
                          m("Area Risk Density", "Risk / File"), AREA_LOCATION,
                          bottoms[0],
                          size=m("Area High Confidence Matches", "High Conf Matches"),
                          title="Location Volume vs Risk Density"),
            table("patterns-table-folder",
                  [AREA_LOCATION, col("DimArea", "folder_name", "Folder"),
                   col("DimArea", "folder_path", "Folder Path"),
                   m("Files With Exported SIT", "Files With SIT"),
                   m("Area File SIT Pairs", "File/SIT Pairs"),
                   m("Area Distinct SITs", "Distinct SITs"),
                   m("Area Risk Density", "Risk / File"),
                   m("Area High Confidence Matches", "High Conf Matches")],
                  bottoms[1], title="Folder Pattern Detail",
                  order_by=m("Area Risk Density"),
                  column_widths={col("DimArea", "folder_path", "Folder Path"): 200.0}),
        ],
    )


def core_pages() -> list[PageSpec]:
    return [
        overview_page(),
        area_hotspots_page(),
        sit_risk_page(),
        location_user_page(),
        area_drilldown_page(),
        file_drillthrough_page(),
        patterns_page(),
    ]
