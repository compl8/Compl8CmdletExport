"""Visual factories: build VisualSpec objects and their Layout config dicts.

Factories return a `VisualSpec`; the project writer resolves deterministic
ids (namespaced by project slug) and writes config.json/visualContainer.json/
filters.json per visual. Titles are emitted in the schema-correct
`vcObjects.title[].properties.text.expr.Literal` shape (golden_vc_title.json)
— never as a bare `singleVisual["title"]` key.
"""

from __future__ import annotations

from dataclasses import dataclass, field as dataclass_field

from .expressions import (
    Field,
    column_properties,
    hex_id,
    literal_expr,
    numeric_id,
    prototype_query,
    solid_color_expr,
)

# Vendored custom visual GUIDs (PowerBI/CustomVisuals/<guid>/).
FORCE_GRAPH_GUID = "ForceGraph1449359463895"
SANKEY_GUID = "sankey02300D1BE6F5427989F3DE31CCA9E0F32020"
WORD_CLOUD_GUID = "WordCloud1447959067750"
CUSTOM_VISUAL_GUIDS = (FORCE_GRAPH_GUID, SANKEY_GUID, WORD_CLOUD_GUID)


@dataclass(frozen=True)
class Rect:
    x: float
    y: float
    width: float
    height: float


@dataclass
class VisualSpec:
    seed: str
    visual_type: str
    rect: Rect
    fields: list[Field] = dataclass_field(default_factory=list)
    projections: dict[str, list[dict[str, object]]] = dataclass_field(default_factory=dict)
    title: str | None = None
    objects: dict[str, object] | None = None
    vc_objects: dict[str, object] | None = None
    order_by: Field | None = None
    filters: list[dict[str, object]] = dataclass_field(default_factory=list)
    z: int | None = None
    how_created: str | None = None


def _refs(fields: list[Field] | None, active_first: bool = False) -> list[dict[str, object]]:
    entries: list[dict[str, object]] = []
    for index, item in enumerate(fields or []):
        entry: dict[str, object] = {"queryRef": item.query_ref()}
        if active_first and index == 0:
            entry["active"] = True
        entries.append(entry)
    return entries


def _one(field: Field, active: bool = False) -> list[dict[str, object]]:
    return _refs([field], active_first=active)


# --- vcObjects helpers (theme-consistent defaults; polish lands in T3/T6) ---

def vc_title(text: str) -> dict[str, object]:
    escaped = text.replace("'", "''")
    text_property = {"text": {"expr": {"Literal": {"Value": f"'{escaped}'"}}}}
    return {"title": [{"properties": text_property}]}


def vc_background(color: str, transparency: int = 0) -> dict[str, object]:
    return {
        "background": [
            {
                "properties": {
                    "show": literal_expr(True),
                    "color": solid_color_expr(color),
                    "transparency": literal_expr(transparency),
                }
            }
        ]
    }


def vc_border(color: str = "#E5E7EB", radius: int = 4) -> dict[str, object]:
    return {
        "border": [
            {
                "properties": {
                    "show": literal_expr(True),
                    "color": solid_color_expr(color),
                    "radius": literal_expr(radius),
                }
            }
        ]
    }


def vc_shadow() -> dict[str, object]:
    return {"dropShadow": [{"properties": {"show": literal_expr(True)}}]}


def merge_vc(*parts: dict[str, object] | None) -> dict[str, object]:
    merged: dict[str, object] = {}
    for part in parts:
        if part:
            merged.update(part)
    return merged


def table_column_widths(widths: dict[Field, float]) -> dict[str, object]:
    """objects.columnWidth entries (golden_column_width.json shape)."""
    return {
        "columnWidth": [
            {
                "properties": {"value": literal_expr(float(width))},
                "selector": {"metadata": field.query_ref()},
            }
            for field, width in widths.items()
        ]
    }


def visual_config(
    spec: VisualSpec,
    project_slug: str,
    *,
    z: int,
    tab_order: int,
) -> dict[str, object]:
    """Resolve a VisualSpec to its config.json dict with deterministic ids."""
    name = hex_id(f"{project_slug}/visual/{spec.seed}")
    single_visual: dict[str, object] = {"visualType": spec.visual_type}
    if spec.projections:
        single_visual["projections"] = spec.projections
    if spec.fields:
        single_visual["prototypeQuery"] = prototype_query(spec.fields, spec.order_by)
    single_visual["drillFilterOtherVisuals"] = True
    if spec.order_by:
        single_visual["hasDefaultSort"] = True
    props = column_properties(spec.fields)
    if props:
        single_visual["columnProperties"] = props
    if spec.objects:
        single_visual["objects"] = spec.objects
    vc_objects = merge_vc(vc_title(spec.title) if spec.title else None, spec.vc_objects)
    if vc_objects:
        single_visual["vcObjects"] = vc_objects
    config: dict[str, object] = {
        "name": name,
        "layouts": [
            {
                "id": 0,
                "position": {
                    "x": spec.rect.x,
                    "y": spec.rect.y,
                    "z": z,
                    "width": spec.rect.width,
                    "height": spec.rect.height,
                    "tabOrder": tab_order,
                },
            }
        ],
        "singleVisual": single_visual,
    }
    if spec.how_created:
        config["howCreated"] = spec.how_created
    return config


def visual_container_json(config: dict[str, object], project_slug: str, seed: str) -> dict[str, object]:
    position = config["layouts"][0]["position"]
    return {
        "height": position["height"],
        "id": numeric_id(f"{project_slug}/visual-container/{seed}"),
        "width": position["width"],
        "x": position["x"],
        "y": position["y"],
        "z": position["z"],
    }


# --- factories ---

def textbox(seed: str, text: str, rect: Rect, *, font_size: int = 18, bold: bool = True) -> VisualSpec:
    text_style: dict[str, object] = {"fontSize": f"{font_size}pt"}
    if bold:
        text_style["fontWeight"] = "bold"
    paragraphs = [
        {"textRuns": [{"value": line, "textStyle": text_style}]}
        for line in text.split("\n")
    ]
    return VisualSpec(
        seed, "textbox", rect,
        objects={"general": [{"properties": {"paragraphs": paragraphs}}]},
    )


def card(seed: str, value: Field, rect: Rect, *, title: str | None = None,
         vc_objects: dict[str, object] | None = None) -> VisualSpec:
    return VisualSpec(
        seed, "card", rect,
        fields=[value],
        projections={"Values": _one(value)},
        title=title,
        objects={
            "categoryLabels": [{"properties": {"show": literal_expr(True), "fontSize": literal_expr(9)}}],
            "labels": [{"properties": {"fontSize": literal_expr(23)}}],
        },
        vc_objects=vc_objects,
    )


def slicer(seed: str, slicer_field: Field, rect: Rect, *, title: str | None = None,
           mode: str = "Dropdown", text_size: int = 10) -> VisualSpec:
    """Compact slicer defaults (T6 owner feedback): Dropdown data mode, small
    item/header text, and plain-click checkbox multi-select
    (strictSingleSelect false = the "Multi-select with CTRL" toggle OFF)."""
    return VisualSpec(
        seed, "slicer", rect,
        fields=[slicer_field],
        projections={"Values": _one(slicer_field)},
        title=title,
        objects={
            "data": [{"properties": {"mode": literal_expr(mode)}}],
            "selection": [{"properties": {
                "singleSelect": literal_expr(False),
                "strictSingleSelect": literal_expr(False),
                "selectAllCheckboxEnabled": literal_expr(True),
            }}],
            "items": [{"properties": {"textSize": literal_expr(text_size)}}],
            "header": [{"properties": {
                "show": literal_expr(True),
                "textSize": literal_expr(text_size),
            }}],
        },
    )


def bar_chart(seed: str, category: Field, values: list[Field], rect: Rect, *,
              title: str | None = None, series: Field | None = None,
              order_by: Field | None = None) -> VisualSpec:
    fields = [category, *values, *([series] if series else [])]
    projections = {"Category": _one(category, active=True), "Y": _refs(values)}
    if series:
        projections["Series"] = _one(series)
    return VisualSpec(
        seed, "clusteredBarChart", rect, fields=fields, projections=projections,
        title=title, order_by=order_by or values[0],
    )


def column_chart(seed: str, category: Field, values: list[Field], rect: Rect, *,
                 title: str | None = None, series: Field | None = None,
                 order_by: Field | None = None) -> VisualSpec:
    """clusteredColumnChart with multi-measure Y support."""
    fields = [category, *values, *([series] if series else [])]
    projections = {"Category": _one(category, active=True), "Y": _refs(values)}
    if series:
        projections["Series"] = _one(series)
    return VisualSpec(
        seed, "clusteredColumnChart", rect, fields=fields, projections=projections,
        title=title, order_by=order_by,
    )


def line_chart(seed: str, category: Field, values: list[Field], rect: Rect, *,
               title: str | None = None, series: Field | None = None) -> VisualSpec:
    fields = [category, *values, *([series] if series else [])]
    projections = {"Category": _one(category, active=True), "Y": _refs(values)}
    if series:
        projections["Series"] = _one(series)
    return VisualSpec(seed, "lineChart", rect, fields=fields, projections=projections, title=title)


def pie_chart(seed: str, legend: Field, value: Field, rect: Rect, *,
              title: str | None = None, series: Field | None = None,
              show_legend: bool = True) -> VisualSpec:
    """pieChart: Legend -> Category, Values -> Y, Details -> Series."""
    fields = [legend, value, *([series] if series else [])]
    projections = {"Category": _one(legend, active=True), "Y": _one(value)}
    if series:
        projections["Series"] = _one(series)
    return VisualSpec(
        seed, "pieChart", rect, fields=fields, projections=projections, title=title,
        objects={"legend": [{"properties": {"show": literal_expr(show_legend)}}]},
    )


def treemap(seed: str, group: Field, value: Field, rect: Rect, *,
            title: str | None = None, details: Field | None = None) -> VisualSpec:
    fields = [group, value, *([details] if details else [])]
    projections = {"Group": _one(group, active=True), "Values": _one(value)}
    if details:
        projections["Details"] = _one(details)
    return VisualSpec(seed, "treemap", rect, fields=fields, projections=projections,
                      title=title, order_by=value)


def scatter_chart(seed: str, x_value: Field, y_value: Field, details: Field, rect: Rect, *,
                  title: str | None = None, size: Field | None = None,
                  series: Field | None = None) -> VisualSpec:
    """scatterChart: Details -> Category, Legend -> Series."""
    fields = [x_value, y_value, details, *([size] if size else []), *([series] if series else [])]
    projections = {
        "X": _one(x_value),
        "Y": _one(y_value),
        "Category": _one(details, active=True),
    }
    if size:
        projections["Size"] = _one(size)
    if series:
        projections["Series"] = _one(series)
    return VisualSpec(seed, "scatterChart", rect, fields=fields, projections=projections, title=title)


def table(seed: str, fields: list[Field], rect: Rect, *, title: str | None = None,
          order_by: Field | None = None,
          column_widths: dict[Field, float] | None = None) -> VisualSpec:
    objects: dict[str, object] = {
        "total": [{"properties": {"totals": literal_expr(False)}}],
        "values": [{"properties": {"fontSize": literal_expr(9)}}],
        "columnHeaders": [{"properties": {"fontSize": literal_expr(9), "wordWrap": literal_expr(True)}}],
    }
    if column_widths:
        objects.update(table_column_widths(column_widths))
    return VisualSpec(
        seed, "tableEx", rect, fields=fields,
        projections={"Values": _refs(fields)},
        title=title, order_by=order_by, objects=objects,
    )


def pivot_table(seed: str, rows: list[Field], columns: list[Field], values: list[Field],
                rect: Rect, *, title: str | None = None) -> VisualSpec:
    fields = [*rows, *columns, *values]
    return VisualSpec(
        seed, "pivotTable", rect, fields=fields,
        projections={
            "Rows": _refs(rows, active_first=True),
            "Columns": _refs(columns),
            "Values": _refs(values),
        },
        title=title,
    )


def sankey(seed: str, source: Field, destination: Field, weight: Field, rect: Rect, *,
           title: str | None = None, show_link_labels: bool = False) -> VisualSpec:
    return VisualSpec(
        seed, SANKEY_GUID, rect,
        fields=[source, destination, weight],
        projections={
            "Source": _one(source, active=True),
            "Destination": _one(destination, active=True),
            "Weight": _one(weight),
        },
        title=title, order_by=weight,
        objects={
            "labels": [{"properties": {"show": literal_expr(True), "fontSize": literal_expr(9), "forceDisplay": literal_expr(False)}}],
            "linkLabels": [{"properties": {"show": literal_expr(show_link_labels)}}],
            "links": [{"properties": {"matchNodeColors": literal_expr(True), "showBorder": literal_expr(False)}}],
            "nodes": [{"properties": {"nodesWidth": literal_expr(18)}}],
            "nodeComplexSettings": [{"properties": {"linksReorder": literal_expr(False), "showResetButon": literal_expr(True)}}],
            "scaleSettings": [{"properties": {"lnScale": literal_expr(True), "provideMinHeight": literal_expr(True)}}],
        },
    )


def force_graph(seed: str, source: Field, target: Field, weight: Field, link_type: Field,
                rect: Rect, *, title: str | None = None,
                source_type: Field | None = None, target_type: Field | None = None,
                charge: int = -90, name_max_length: int = 48) -> VisualSpec:
    fields = [source, target, weight, link_type,
              *([source_type] if source_type else []),
              *([target_type] if target_type else [])]
    projections = {
        "Source": _one(source, active=True),
        "Target": _one(target, active=True),
        "Weight": _one(weight),
        "LinkType": _one(link_type, active=True),
    }
    if source_type:
        projections["SourceType"] = _one(source_type)
    if target_type:
        projections["TargetType"] = _one(target_type)
    return VisualSpec(
        seed, FORCE_GRAPH_GUID, rect,
        fields=fields,
        projections=projections,
        title=title, order_by=weight,
        objects={
            "animation": [{"properties": {"show": literal_expr(False)}}],
            "labels": [{"properties": {"show": literal_expr(True), "fontSize": literal_expr(9), "allowIntersection": literal_expr(False)}}],
            "links": [{"properties": {"showArrow": literal_expr(True), "showLabel": literal_expr(False), "thickenLink": literal_expr(True)}}],
            "nodes": [{"properties": {"displayImage": literal_expr(False), "nameMaxLength": literal_expr(name_max_length), "highlightReachableLinks": literal_expr(True)}}],
            "size": [{"properties": {"charge": literal_expr(charge), "boundedByBox": literal_expr(True)}}],
        },
    )


def word_cloud(seed: str, category: Field, value: Field, rect: Rect, *,
               title: str | None = None) -> VisualSpec:
    return VisualSpec(
        seed, WORD_CLOUD_GUID, rect,
        fields=[category, value],
        projections={"Category": _one(category, active=True), "Values": _one(value)},
        title=title,
    )


def back_button(seed: str, rect: Rect | None = None) -> VisualSpec:
    """Back actionButton (golden_back_button.json shape) for drillthrough pages."""
    return VisualSpec(
        seed, "actionButton", rect or Rect(0, 0, 100, 40),
        objects={
            "icon": [
                {
                    "properties": {"shapeType": {"expr": {"Literal": {"Value": "'back'"}}}},
                    "selector": {"id": "default"},
                }
            ]
        },
        vc_objects={
            "visualLink": [
                {
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "type": {"expr": {"Literal": {"Value": "'Back'"}}},
                    }
                }
            ]
        },
        how_created="InsertVisualButton",
    )
