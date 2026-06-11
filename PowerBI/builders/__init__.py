"""Power BI project-generation engine.

Public surface (consumed by per-project builders in T3/T4):

- `pbi_project.ProjectSpec` / `build_project` — assemble a full PbixProj
- `tmdl_model.MeasureSpec` — declarative DAX measures
- `report_layout.PageSpec` + layout grid constants/helpers
- `visual_factories` — visual constructors (card, bar, pie, table, ...)
- `filters` — report/page/visual filter + drillthrough shapes
- `expressions.col` / `expressions.meas` — field references
"""

from __future__ import annotations

from .expressions import Field, col, meas
from .pbi_project import ProjectSpec, build_project
from .report_layout import PageSpec
from .tmdl_model import DEFAULT_PARQUET_ROOT, MeasureSpec

__all__ = [
    "DEFAULT_PARQUET_ROOT",
    "Field",
    "MeasureSpec",
    "PageSpec",
    "ProjectSpec",
    "build_project",
    "col",
    "meas",
]
