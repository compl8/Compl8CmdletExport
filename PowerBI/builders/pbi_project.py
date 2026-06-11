"""Facade: assemble a complete pbi-tools PbixProj from a ProjectSpec.

The semantic model is generated from `parquet_builder.star.schema` (the SSOT)
via `tmdl_model`; report content (pages, visuals, measures, filters) is
declared by per-project builders (build_smoke now, AE/CE in T3/T4).
"""

from __future__ import annotations

import shutil
from dataclasses import dataclass, field as dataclass_field
from pathlib import Path

from parquet_builder.star.schema import model_tables

from .report_layout import (
    PageSpec,
    named_report_filters,
    page_pod,
    report_config,
    write_json,
    write_section,
    write_text,
)
from .theme import copy_custom_visuals, embed_theme, linguistic_schema_xml, theme_collection
from .tmdl_model import DEFAULT_PARQUET_ROOT, MeasureSpec, write_model
from .visual_factories import CUSTOM_VISUAL_GUIDS, VisualSpec


@dataclass
class ProjectSpec:
    """A full Power BI project: model name, measures, pages, report filters."""

    name: str  # database/model display name
    slug: str  # deterministic-id namespace (e.g. "activity-explorer-v6")
    pages: list[PageSpec] = dataclass_field(default_factory=list)
    measures: list[MeasureSpec] = dataclass_field(default_factory=list)
    report_filters: list[dict[str, object]] = dataclass_field(default_factory=list)
    active_section_index: int = 0


def used_custom_visuals(pages: list[PageSpec]) -> list[str]:
    used: list[str] = []
    for page in pages:
        for visual in page.visuals:
            if visual.visual_type in CUSTOM_VISUAL_GUIDS and visual.visual_type not in used:
                used.append(visual.visual_type)
    return used


def _diagram_layout() -> dict[str, object]:
    tables = [table.name for table in model_tables()]
    per_row = 6
    nodes = [
        {
            "location": {"x": (index % per_row) * 320, "y": (index // per_row) * 320},
            "nodeIndex": name,
            "size": {"height": 240, "width": 260},
            "zIndex": index,
        }
        for index, name in enumerate(tables)
    ]
    return {
        "version": "1.1.0",
        "diagrams": [
            {
                "ordinal": 0,
                "scrollPosition": {"x": 0, "y": 0},
                "nodes": nodes,
                "name": "All tables",
                "zoomValue": 65,
            }
        ],
    }


def build_project(
    spec: ProjectSpec,
    out_dir: Path,
    *,
    parquet_root: str = DEFAULT_PARQUET_ROOT,
    overwrite: bool = False,
) -> Path:
    """Generate the PbixProj folder; returns out_dir."""
    if out_dir.exists():
        if not overwrite:
            raise FileExistsError(f"Output folder already exists: {out_dir}")
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    # Static content only (no timestamps): two runs produce identical bytes.
    write_json(out_dir / ".pbixproj.json", {
        "version": "1.0",
        "settings": {"model": {"serializationMode": "Tmdl"}},
    })
    write_text(out_dir / "Version.txt", "1.28")
    write_json(out_dir / "ReportMetadata.json", {
        "Version": 5,
        "AutoCreatedRelationships": [],
        "FileDescription": f"{spec.name} (generated from the star-schema SSOT)",
        "CreatedFrom": "pbi-tools",
        "CreatedFromRelease": "1.2.0",
    })
    write_json(out_dir / "ReportSettings.json", {
        "Version": 4,
        "ReportSettings": {},
        "QueriesSettings": {
            "TypeDetectionEnabled": False,
            "RelationshipImportEnabled": False,
            "RunBackgroundAnalysis": False,
            "Version": "2.153.910.0",
        },
    })
    write_json(out_dir / "DiagramLayout.json", _diagram_layout())
    write_text(out_dir / "LinguisticSchema.xml", linguistic_schema_xml())

    write_model(
        out_dir / "Model",
        model_name=spec.name,
        project_slug=spec.slug,
        parquet_root=parquet_root,
        measures=spec.measures,
    )

    theme_packages = embed_theme(out_dir)
    visual_guids = used_custom_visuals(spec.pages)
    visual_packages = copy_custom_visuals(out_dir, visual_guids)

    report_dir = out_dir / "Report"
    report_json: dict[str, object] = {
        "id": 0,
        "layoutOptimization": 0,
        "pods": [
            page_pod(spec.slug, page, ordinal)
            for ordinal, page in enumerate(spec.pages)
        ],
        "resourcePackages": [*visual_packages, *theme_packages],
    }
    if visual_guids:
        report_json["publicCustomVisuals"] = visual_guids
    write_json(report_dir / "report.json", report_json)
    write_json(report_dir / "config.json",
               report_config(theme_collection(), spec.active_section_index))
    write_json(report_dir / "filters.json",
               named_report_filters(spec.slug, spec.report_filters))
    for ordinal, page in enumerate(spec.pages):
        write_section(report_dir, spec.slug, page, ordinal)
    return out_dir


__all__ = [
    "MeasureSpec",
    "PageSpec",
    "ProjectSpec",
    "VisualSpec",
    "build_project",
    "used_custom_visuals",
]
