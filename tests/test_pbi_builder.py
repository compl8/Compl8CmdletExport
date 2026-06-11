"""Tests for the Power BI project-generation engine (PowerBI/builders).

Pure-python assertions — no pbi-tools required. Golden shape fixtures in
tests/fixtures/ were copied from the hand-built reference report (the
fidelity baseline); tests never reference the legacy repo at runtime.
"""

from __future__ import annotations

import filecmp
import json
from pathlib import Path

import pytest

from parquet_builder.star import schema
from PowerBI.builders.build_smoke import smoke_project
from PowerBI.builders.pbi_project import build_project

FIXTURES = Path(__file__).parent / "fixtures"


def _load(path: Path) -> object:
    return json.loads(path.read_text(encoding="utf-8"))


def _shape(obj: object) -> object:
    """Structure signature: dict keys and list nesting, ignoring leaf values."""
    if isinstance(obj, dict):
        return {key: _shape(value) for key, value in sorted(obj.items())}
    if isinstance(obj, list):
        return [_shape(obj[0])] if obj else []
    return type(obj).__name__


@pytest.fixture(scope="module")
def project_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("pbi") / "pbix"
    return build_project(smoke_project(), out)


def _visuals_by_type(project_dir: Path, section: str) -> dict[str, dict]:
    visuals: dict[str, dict] = {}
    for config_path in sorted(project_dir.glob(f"Report/sections/{section}/visualContainers/*/config.json")):
        config = _load(config_path)
        visuals.setdefault(config["singleVisual"]["visualType"], config)
    return visuals


# --- model: generated from the SSOT ----------------------------------------

def test_model_contains_every_loaded_ssot_table_and_column(project_dir: Path) -> None:
    for table in schema.model_tables():
        tmdl_path = project_dir / "Model" / "tables" / f"{table.name}.tmdl"
        assert tmdl_path.exists(), f"missing table tmdl: {table.name}"
        tmdl = tmdl_path.read_text(encoding="utf-8")
        assert f"table {table.name}" in tmdl
        for column in table.columns:
            assert f"sourceColumn: {column.name}" in tmdl, f"{table.name}.{column.name} missing"
            # M partition prunes to and types every SSOT column
            assert f'"{column.name}"' in tmdl
        assert f'ParquetRoot & "\\{table.name}.parquet"' in tmdl


def test_pipeline_only_and_index_tables_excluded(project_dir: Path) -> None:
    tables_dir = project_dir / "Model" / "tables"
    assert not (tables_dir / "archive_raw.tmdl").exists()
    assert not (tables_dir / "activity_record_index.tmdl").exists()
    model_tmdl = (project_dir / "Model" / "model.tmdl").read_text(encoding="utf-8")
    assert "archive_raw" not in model_tmdl
    assert "activity_record_index" not in model_tmdl


def test_relationships_emitted_including_inactive(project_dir: Path) -> None:
    tmdl = (project_dir / "Model" / "relationships.tmdl").read_text(encoding="utf-8")
    emitted = [block for block in tmdl.split("relationship ") if block.strip()]
    expected = schema.model_relationships()
    assert len(emitted) == len(expected)
    for rel in expected:
        from_ref = f"fromColumn: {rel.from_table}.{rel.from_column}"
        assert from_ref in tmdl, f"missing relationship {from_ref}"
    inactive_expected = sum(1 for rel in expected if not rel.active)
    assert tmdl.count("isActive: false") == inactive_expected
    # the index-table relationship must NOT be in the model
    assert "activity_record_index" not in tmdl


def test_date_table_marked(project_dir: Path) -> None:
    tmdl = (project_dir / "Model" / "tables" / "dim_date.tmdl").read_text(encoding="utf-8")
    assert "dataCategory: Time" in tmdl
    date_block = tmdl.split("column date\n")[1].split("column ")[0]
    assert "isKey" in date_block


def test_fact_foreign_keys_hidden(project_dir: Path) -> None:
    tmdl = (project_dir / "Model" / "tables" / "fact_activity.tmdl").read_text(encoding="utf-8")
    for fk in ("date_key", "user_id", "workload_id", "activity_id"):
        block = tmdl.split(f"column {fk}\n")[1].split("column ")[0]
        assert "isHidden" in block, f"fact_activity.{fk} not hidden"


def test_measures_emitted_with_folder_and_format(project_dir: Path) -> None:
    tmdl = (project_dir / "Model" / "tables" / "fact_activity.tmdl").read_text(encoding="utf-8")
    assert "measure 'Total Activities' =" in tmdl
    assert "displayFolder: Smoke Metrics" in tmdl
    assert "formatString: #,0" in tmdl


# --- report: golden shapes from the reference report ------------------------

def test_vcobjects_title_matches_golden_shape(project_dir: Path) -> None:
    golden = _load(FIXTURES / "golden_vc_title.json")
    bar = _visuals_by_type(project_dir, "000_Smoke_Overview")["clusteredBarChart"]
    emitted = bar["singleVisual"]["vcObjects"]["title"]
    assert _shape(emitted) == _shape(golden)
    # the non-schema bare key the legacy generator used must be gone
    assert "title" not in bar["singleVisual"]


def test_back_button_matches_golden_shape(project_dir: Path) -> None:
    golden = _load(FIXTURES / "golden_back_button.json")
    button = _visuals_by_type(project_dir, "010_Smoke_Drill")["actionButton"]
    assert _shape(button["singleVisual"]) == _shape(golden)


def test_drillthrough_filters_match_golden_shape(project_dir: Path) -> None:
    golden = _load(FIXTURES / "golden_drillthrough_filters.json")
    emitted = _load(project_dir / "Report" / "sections" / "010_Smoke_Drill" / "filters.json")
    drill = [entry for entry in emitted if entry.get("howCreated") == 5]
    assert drill, "no howCreated:5 drillthrough filters emitted"
    assert _shape(drill[0]) == _shape(golden[0])


def test_drillthrough_pod_binds_section_filters(project_dir: Path) -> None:
    report = _load(project_dir / "Report" / "report.json")
    pods = [pod for pod in report["pods"] if pod.get("type") == 1]
    assert len(pods) == 1
    section = _load(project_dir / "Report" / "sections" / "010_Smoke_Drill" / "section.json")
    assert pods[0]["boundSection"] == section["name"]
    filters = _load(project_dir / "Report" / "sections" / "010_Smoke_Drill" / "filters.json")
    filter_names = {entry["name"] for entry in filters if entry.get("howCreated") == 5}
    parameters = json.loads(pods[0]["parameters"])
    assert parameters and all(param["boundFilter"] in filter_names for param in parameters)


def test_report_not_in_filter_matches_golden_shape(project_dir: Path) -> None:
    golden = _load(FIXTURES / "golden_report_not_in_filter.json")
    emitted = _load(project_dir / "Report" / "filters.json")
    assert len(emitted) == 1
    assert _shape(emitted[0]) == _shape(golden)
    values = emitted[0]["filter"]["Where"][0]["Condition"]["Not"]["Expression"]["In"]["Values"]
    assert values[0][0]["Literal"]["Value"] == "null"


def test_column_width_matches_golden_shape(project_dir: Path) -> None:
    golden = _load(FIXTURES / "golden_column_width.json")
    table = _visuals_by_type(project_dir, "000_Smoke_Overview")["tableEx"]
    emitted = table["singleVisual"]["objects"]["columnWidth"]
    assert _shape(emitted) == _shape(golden)


def test_page_and_visual_filters_emitted(project_dir: Path) -> None:
    page_filters = _load(project_dir / "Report" / "sections" / "000_Smoke_Overview" / "filters.json")
    categorical = [entry for entry in page_filters if entry.get("type") == "Categorical"]
    assert categorical and categorical[0]["howCreated"] == 1
    table_dir = next(
        path for path in (project_dir / "Report" / "sections" / "000_Smoke_Overview" / "visualContainers").iterdir()
        if "_tableEx " in path.name
    )
    visual_filters = _load(table_dir / "filters.json")
    assert visual_filters[0]["type"] == "Advanced"
    condition = visual_filters[0]["filter"]["Where"][0]["Condition"]
    assert "Comparison" in condition


def test_theme_embedded(project_dir: Path) -> None:
    assert (project_dir / "StaticResources" / "RegisteredResources" / "Compl8.Theme.json").exists()
    assert (project_dir / "StaticResources" / "SharedResources" / "BaseThemes" / "CY26SU02.json").exists()
    config = _load(project_dir / "Report" / "config.json")
    assert config["themeCollection"]["customTheme"]["name"] == "Compl8.Theme.json"
    assert config["themeCollection"]["baseTheme"]["name"] == "CY26SU02"
    report = _load(project_dir / "Report" / "report.json")
    package_names = [entry["resourcePackage"]["name"] for entry in report["resourcePackages"]]
    assert "RegisteredResources" in package_names
    assert "SharedResources" in package_names


def test_only_used_custom_visuals_packaged(project_dir: Path) -> None:
    visuals_dir = project_dir / "CustomVisuals"
    packaged = {path.name for path in visuals_dir.iterdir()} if visuals_dir.exists() else set()
    assert packaged == {"WordCloud1447959067750"}  # smoke uses only the word cloud
    report = _load(project_dir / "Report" / "report.json")
    assert report["publicCustomVisuals"] == ["WordCloud1447959067750"]


def test_chrome_emitted(project_dir: Path) -> None:
    assert (project_dir / "LinguisticSchema.xml").exists()
    assert (project_dir / "Model" / "cultures" / "en-AU.tmdl").exists()
    model_tmdl = (project_dir / "Model" / "model.tmdl").read_text(encoding="utf-8")
    assert "ref cultureInfo en-AU" in model_tmdl
    config = _load(project_dir / "Report" / "config.json")
    assert config["settings"]["useEnhancedTooltips"] is True
    assert config["settings"]["queryLimitOption"] == 6
    drill_config = _load(project_dir / "Report" / "sections" / "010_Smoke_Drill" / "config.json")
    assert drill_config["objects"]["outspacePane"][0]["properties"]["width"]["expr"]["Literal"]["Value"] == "362L"


# --- determinism -------------------------------------------------------------

def test_two_builds_are_byte_identical(project_dir: Path, tmp_path: Path) -> None:
    second = build_project(smoke_project(), tmp_path / "pbix")
    comparison = filecmp.dircmp(project_dir, second)
    differences: list[str] = []

    def _collect(cmp: filecmp.dircmp, prefix: str = "") -> None:
        differences.extend(f"{prefix}{name}" for name in cmp.diff_files)
        differences.extend(f"{prefix}{name}" for name in cmp.left_only + cmp.right_only)
        for sub_name, sub_cmp in cmp.subdirs.items():
            _collect(sub_cmp, f"{prefix}{sub_name}/")

    _collect(comparison)
    assert differences == [], f"non-deterministic output: {differences}"
