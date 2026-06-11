"""Layout grid, page specs, and Report/ folder emission.

Canvas geometry is named (no magic numbers in page definitions): 1280x720
canvas, 32px side margins, fixed title/KPI/chart/table rows matching the
established report style.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field as dataclass_field
from pathlib import Path

from .expressions import hex_id
from .filters import drillthrough_field_filter, drillthrough_pod_parameters
from .visual_factories import (
    Rect,
    VisualSpec,
    back_button,
    visual_config,
    visual_container_json,
)

# --- canvas grid -------------------------------------------------------------

CANVAS_WIDTH = 1280
CANVAS_HEIGHT = 720
MARGIN = 32
GUTTER = 12
CONTENT_WIDTH = CANVAS_WIDTH - 2 * MARGIN  # 1216

TITLE_Y = 18
TITLE_HEIGHT = 34
KPI_ROW_Y = 64
KPI_HEIGHT = 88
SLICER_ROW_Y = 56
SLICER_HEIGHT = 96
CHART_ROW_Y = 184
CHART_HEIGHT = 224
TABLE_ROW_Y = 436
TABLE_HEIGHT = CANVAS_HEIGHT - TABLE_ROW_Y - 40  # 244


def grid_row(count: int, y: float, height: float, *,
             x0: float = MARGIN, total_width: float = CONTENT_WIDTH,
             gutter: float = GUTTER) -> list[Rect]:
    """Split a horizontal band into `count` equal cells."""
    if count < 1:
        raise ValueError("count must be >= 1")
    width = (total_width - gutter * (count - 1)) / count
    return [Rect(x0 + index * (width + gutter), y, width, height) for index in range(count)]


def row_of_cards(count: int, *, y: float = KPI_ROW_Y, height: float = KPI_HEIGHT) -> list[Rect]:
    """KPI card row cells (caps card width at 200px, golden-report style)."""
    cells = grid_row(count, y, height)
    return [Rect(cell.x, cell.y, min(cell.width, 200.0), cell.height) for cell in cells]


def title_rect(width: float = 760) -> Rect:
    return Rect(MARGIN, TITLE_Y, width, TITLE_HEIGHT)


def full_width(y: float, height: float) -> Rect:
    return Rect(MARGIN, y, CONTENT_WIDTH, height)


# --- pages -------------------------------------------------------------------

@dataclass
class PageSpec:
    """A report page. Drillthrough pages declare their bound fields and get
    the howCreated:5 section filters plus an auto Back button."""

    folder: str  # e.g. "000_Overview" (section directory name)
    display_name: str
    visuals: list[VisualSpec] = dataclass_field(default_factory=list)
    page_filters: list[dict[str, object]] = dataclass_field(default_factory=list)
    drillthrough_fields: list[tuple[str, str]] = dataclass_field(default_factory=list)  # (table, column)
    outspace_pane_width: int | None = None
    width: int = CANVAS_WIDTH
    height: int = CANVAS_HEIGHT

    def is_drillthrough(self) -> bool:
        return bool(self.drillthrough_fields)


def section_name(project_slug: str, page: PageSpec) -> str:
    return hex_id(f"{project_slug}/section/{page.folder}")


def _section_config(page: PageSpec) -> dict[str, object]:
    if page.outspace_pane_width is None:
        return {}
    return {
        "objects": {
            "outspacePane": [
                {
                    "properties": {
                        "width": {
                            "expr": {"Literal": {"Value": f"{page.outspace_pane_width}L"}}
                        }
                    }
                }
            ]
        }
    }


def _named_page_filters(project_slug: str, page: PageSpec) -> list[dict[str, object]]:
    """Section filters.json: page filters then drillthrough field filters."""
    named: list[dict[str, object]] = []
    for index, page_filter in enumerate(page.page_filters):
        named.append({
            "name": hex_id(f"{project_slug}/page-filter/{page.folder}/{index}"),
            **page_filter,
        })
    for table, column in page.drillthrough_fields:
        named.append({
            "name": hex_id(f"{project_slug}/drill-filter/{page.folder}/{table}.{column}"),
            **drillthrough_field_filter(table, column),
        })
    return named


def page_pod(project_slug: str, page: PageSpec, ordinal: int) -> dict[str, object]:
    """report.json pod for a page; drillthrough pages bind their filters."""
    pod: dict[str, object] = {
        "boundSection": section_name(project_slug, page),
        "config": "{}",
        "name": hex_id(f"{project_slug}/pod/{page.folder}"),
    }
    if page.is_drillthrough():
        drill_named = [
            entry for entry in _named_page_filters(project_slug, page)
            if entry.get("howCreated") == 5
        ]
        parameters = drillthrough_pod_parameters(
            [(entry["name"], entry) for entry in drill_named]
        )
        for index, parameter in enumerate(parameters):
            parameter["name"] = hex_id(f"{project_slug}/pod-param/{page.folder}/{index}")
            # parameter order: name, boundFilter, fieldExpr (golden shape)
            parameters[index] = {
                "name": parameter["name"],
                "boundFilter": parameter["boundFilter"],
                "fieldExpr": parameter["fieldExpr"],
            }
        pod["parameters"] = json.dumps(parameters)
        pod["type"] = 1
    else:
        pod["referenceScope"] = 1
    return pod


def write_section(report_dir: Path, project_slug: str, page: PageSpec, ordinal: int) -> None:
    section_dir = report_dir / "sections" / page.folder
    write_json(section_dir / "section.json", {
        "displayName": page.display_name,
        "displayOption": 1,
        "height": page.height,
        "name": section_name(project_slug, page),
        "ordinal": ordinal,
        "width": page.width,
    })
    write_json(section_dir / "config.json", _section_config(page))
    write_json(section_dir / "filters.json", _named_page_filters(project_slug, page))

    visuals = list(page.visuals)
    if page.is_drillthrough() and not any(v.visual_type == "actionButton" for v in visuals):
        visuals.insert(0, back_button(f"{page.folder}/auto-back"))

    for index, visual in enumerate(visuals):
        z = visual.z if visual.z is not None else (index + 1) * 1000
        config = visual_config(visual, project_slug, z=z, tab_order=z)
        name = str(config["name"])
        folder = section_dir / "visualContainers" / f"{index:05d}_{visual.visual_type} ({name[:5]})"
        write_json(folder / "config.json", config)
        write_json(folder / "visualContainer.json",
                   visual_container_json(config, project_slug, visual.seed))
        named_filters = [
            {"name": hex_id(f"{project_slug}/visual-filter/{visual.seed}/{f_index}"), **visual_filter}
            for f_index, visual_filter in enumerate(visual.filters)
        ]
        write_json(folder / "filters.json", named_filters)


def named_report_filters(project_slug: str,
                         report_filters: list[dict[str, object]]) -> list[dict[str, object]]:
    return [
        {"name": hex_id(f"{project_slug}/report-filter/{index}"), **report_filter}
        for index, report_filter in enumerate(report_filters)
    ]


def report_config(theme_collection: dict[str, object], active_section_index: int = 0) -> dict[str, object]:
    return {
        "version": "5.70",
        "themeCollection": theme_collection,
        "activeSectionIndex": active_section_index,
        "defaultDrillFilterOtherVisuals": True,
        "linguisticSchemaSyncVersion": 2,
        "settings": {
            "useNewFilterPaneExperience": True,
            "allowChangeFilterTypes": True,
            "useStylableVisualContainerHeader": True,
            "queryLimitOption": 6,
            "useEnhancedTooltips": True,
            "exportDataMode": 1,
            "useDefaultAggregateDisplayName": True,
        },
        "objects": {
            "section": [
                {"properties": {"verticalAlignment": {"expr": {"Literal": {"Value": "'Top'"}}}}}
            ],
            "outspacePane": [
                {"properties": {"expanded": {"expr": {"Literal": {"Value": "false"}}}}}
            ],
        },
    }


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2), encoding="utf-8", newline="\n")


def write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8", newline="\n")
