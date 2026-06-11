"""Port assertions for the Content Explorer SIT Risk report (T4).

The legacy generated 15-page / 71-measure / 22-table report is the contract:
every legacy table and measure name must exist in the emitted TMDL, every
legacy page must be mapped, drillthrough must be wired on the File
Drillthrough page (an upgrade — the legacy page had none), every data visual
must carry a curated title, and builds must be deterministic.
"""

from __future__ import annotations

import filecmp
import json
from pathlib import Path

import pytest

from PowerBI.builders.build_content_explorer import (
    LEGACY_PAGE_MAPPING,
    ce_pages,
    ce_project,
)
from PowerBI.builders.ce_measures import LEGACY_MEASURE_NAMES, MEASURES
from PowerBI.builders.ce_schema import (
    CE_PARQUET_FILES,
    CE_RELATIONSHIPS,
    CE_TABLES,
)
from PowerBI.builders.pbi_project import build_project

EXPECTED_PAGE_COUNT = 15
EXPECTED_MEASURE_COUNT = 71
EXPECTED_TABLE_COUNT = 22
EXPECTED_RELATIONSHIP_COUNT = 21
LEGACY_VISUAL_COUNT = 164  # ported superset must be >= this

# The 22 table names of the legacy model (PBI_QueryOrder).
LEGACY_TABLE_NAMES = (
    "DimSIT", "DimQGISCFDLM", "DimLocation", "DimArea", "DimFile", "DimUser",
    "BridgeFileUser", "DimClassifier", "FactLocationSIT", "FactFileSITTag",
    "FactArea", "FactAreaSIT", "FactDetectedSITByLocation", "DimGraphCluster",
    "DimGraphNode", "DimUserDepartment", "FactGraphEdge", "FactGraphEdgeFocus",
    "FactSankeyUserLocationFlow", "FactSankeySensitivityFlow",
    "FactGraphAdjacency", "FactExecFinding",
)

DRILL_PAGE = "040_File_Drillthrough"


@pytest.fixture(scope="module")
def project_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("ce") / "pbix"
    return build_project(ce_project(), out)


def _model_tmdl_text(project_dir: Path) -> str:
    return "\n".join(
        path.read_text(encoding="utf-8")
        for path in sorted((project_dir / "Model" / "tables").glob("*.tmdl"))
    )


# --- model: 22 legacy tables, legacy parquet sources, 21 relationships -------

def test_every_legacy_table_emitted(project_dir: Path) -> None:
    assert len(LEGACY_TABLE_NAMES) == EXPECTED_TABLE_COUNT
    assert tuple(table.name for table in CE_TABLES) == LEGACY_TABLE_NAMES
    emitted = {path.stem for path in (project_dir / "Model" / "tables").glob("*.tmdl")}
    assert emitted == set(LEGACY_TABLE_NAMES)


def test_partitions_use_legacy_parquet_files(project_dir: Path) -> None:
    assert set(CE_PARQUET_FILES) == set(LEGACY_TABLE_NAMES)
    for table in CE_TABLES:
        tmdl = (project_dir / "Model" / "tables" / f"{table.name}.tmdl").read_text(
            encoding="utf-8")
        assert f'ParquetRoot & "\\{CE_PARQUET_FILES[table.name]}"' in tmdl
        for column in table.columns:
            assert f"sourceColumn: {column.name}" in tmdl, (
                f"{table.name}.{column.name} missing")


def test_relationships_emitted(project_dir: Path) -> None:
    assert len(CE_RELATIONSHIPS) == EXPECTED_RELATIONSHIP_COUNT
    tmdl = (project_dir / "Model" / "relationships.tmdl").read_text(encoding="utf-8")
    for rel in CE_RELATIONSHIPS:
        assert f"fromColumn: {rel.from_table}.{rel.from_column}" in tmdl
        assert f"toColumn: {rel.to_table}.{rel.to_column}" in tmdl
    assert "isActive: false" not in tmdl  # legacy CE model: all active


def test_parquet_root_default_is_changeme(project_dir: Path) -> None:
    expressions = (project_dir / "Model" / "expressions.tmdl").read_text(
        encoding="utf-8")
    assert "CHANGEME" in expressions  # never a machine-specific path


# --- measures: all 71 legacy names ported -------------------------------------

def test_every_legacy_measure_ported(project_dir: Path) -> None:
    assert len(LEGACY_MEASURE_NAMES) == EXPECTED_MEASURE_COUNT
    assert len({measure.name for measure in MEASURES}) == EXPECTED_MEASURE_COUNT
    tmdl = _model_tmdl_text(project_dir)
    for name in LEGACY_MEASURE_NAMES:
        bare = f"measure {name} =" in tmdl
        quoted = f"measure '{name}' =" in tmdl
        assert bare or quoted, f"measure {name!r} not emitted"


def test_no_whole_table_dimfile_filter_iterators() -> None:
    """The three legacy FILTER ( DimFile, ... ) iterators were rewritten as
    column predicates (same names, same semantics)."""
    for measure in MEASURES:
        assert "FILTER ( DimFile" not in measure.dax, measure.name


# --- pages ---------------------------------------------------------------------

def test_every_legacy_page_is_mapped(project_dir: Path) -> None:
    assert len(LEGACY_PAGE_MAPPING) == EXPECTED_PAGE_COUNT
    emitted = {path.name for path in (project_dir / "Report" / "sections").iterdir()}
    for legacy, target in LEGACY_PAGE_MAPPING.items():
        assert target in emitted, f"{legacy!r} maps to missing page {target!r}"


def test_expected_page_count(project_dir: Path) -> None:
    pages = ce_pages()
    assert len(pages) == EXPECTED_PAGE_COUNT
    emitted = list((project_dir / "Report" / "sections").iterdir())
    assert len(emitted) == EXPECTED_PAGE_COUNT


def test_pages_in_nav_order(project_dir: Path) -> None:
    folders = [page.folder for page in ce_pages()]
    assert folders == sorted(folders)
    ordinals = {}
    for section_dir in (project_dir / "Report" / "sections").iterdir():
        section = json.loads((section_dir / "section.json").read_text(encoding="utf-8"))
        ordinals[section_dir.name] = section["ordinal"]
    assert [name for name, _ in sorted(ordinals.items(), key=lambda kv: kv[1])] == folders


def test_visual_count_is_superset(project_dir: Path) -> None:
    visuals = list((project_dir / "Report" / "sections").glob(
        "*/visualContainers/*/config.json"))
    assert len(visuals) >= LEGACY_VISUAL_COUNT


# --- bindings reference declared columns and measures ---------------------------

def test_all_visual_fields_resolve() -> None:
    columns = {(table.name, column.name)
               for table in CE_TABLES for column in table.columns}
    measure_names = {(measure.table, measure.name) for measure in MEASURES}
    for page in ce_pages():
        for visual in page.visuals:
            for field in visual.fields:
                if field.kind == "column":
                    assert (field.table, field.name) in columns, (
                        f"{page.folder}/{visual.seed}: unknown column "
                        f"{field.table}.{field.name}")
                else:
                    assert (field.table, field.name) in measure_names, (
                        f"{page.folder}/{visual.seed}: unknown measure "
                        f"{field.table}.{field.name}")


# --- drillthrough upgrade on 040 ------------------------------------------------

def test_file_drillthrough_wired(project_dir: Path) -> None:
    filters = json.loads(
        (project_dir / "Report" / "sections" / DRILL_PAGE / "filters.json").read_text(
            encoding="utf-8"))
    drill = [entry for entry in filters if entry.get("howCreated") == 5]
    bound = {
        (entry["expression"]["Column"]["Expression"]["SourceRef"]["Entity"],
         entry["expression"]["Column"]["Property"])
        for entry in drill
    }
    assert ("DimFile", "file_name") in bound
    assert ("DimSIT", "sit_name") in bound
    assert ("DimLocation", "location_name") in bound

    report = json.loads((project_dir / "Report" / "report.json").read_text(
        encoding="utf-8"))
    drill_pods = [pod for pod in report["pods"] if pod.get("type") == 1]
    assert len(drill_pods) == 1
    section = json.loads(
        (project_dir / "Report" / "sections" / DRILL_PAGE / "section.json").read_text(
            encoding="utf-8"))
    assert drill_pods[0]["boundSection"] == section["name"]
    parameters = json.loads(drill_pods[0]["parameters"])
    filter_names = {entry["name"] for entry in drill}
    assert parameters and all(
        param["boundFilter"] in filter_names for param in parameters)


def test_file_drillthrough_has_back_button(project_dir: Path) -> None:
    configs = (project_dir / "Report" / "sections" / DRILL_PAGE).glob(
        "visualContainers/*/config.json")
    types = [json.loads(path.read_text(encoding="utf-8"))["singleVisual"]["visualType"]
             for path in configs]
    assert "actionButton" in types


# --- polish: every data visual carries a curated title ---------------------------

def test_every_visual_has_vc_title(project_dir: Path) -> None:
    untitled_ok = {"textbox", "actionButton"}
    for config_path in (project_dir / "Report" / "sections").glob(
            "*/visualContainers/*/config.json"):
        config = json.loads(config_path.read_text(encoding="utf-8"))
        single = config["singleVisual"]
        if single["visualType"] in untitled_ok:
            continue
        title = single.get("vcObjects", {}).get("title")
        assert title, f"missing vcObjects title: {config_path.parent.name}"
        value = title[0]["properties"]["text"]["expr"]["Literal"]["Value"]
        assert value.strip("'"), f"empty title: {config_path.parent.name}"
        assert "title" not in single  # non-schema bare key must never appear


# --- determinism ------------------------------------------------------------------

def test_two_builds_are_byte_identical(project_dir: Path, tmp_path: Path) -> None:
    second = build_project(ce_project(), tmp_path / "pbix")
    differences: list[str] = []

    def _collect(cmp: filecmp.dircmp, prefix: str = "") -> None:
        differences.extend(f"{prefix}{name}" for name in cmp.diff_files)
        differences.extend(f"{prefix}{name}" for name in cmp.left_only + cmp.right_only)
        for sub_name, sub_cmp in cmp.subdirs.items():
            _collect(sub_cmp, f"{prefix}{sub_name}/")

    _collect(filecmp.dircmp(project_dir, second))
    assert differences == [], f"non-deterministic output: {differences}"
