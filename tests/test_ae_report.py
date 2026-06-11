"""Superset assertions for the Activity Explorer report (T3).

The legacy 29-page report is the contract: every legacy page must be mapped
to an emitted page, every legacy measure name must exist in the emitted TMDL,
and report-wide invariants (anchored time-intel, titled visuals, deterministic
builds, SSOT-valid bindings) must hold.
"""

from __future__ import annotations

import filecmp
import json
import re
from pathlib import Path

import pytest

from parquet_builder.star import schema
from PowerBI.builders.ae_measures import LEGACY_MEASURE_NAMES, MEASURES
from PowerBI.builders.build_activity_explorer import (
    LEGACY_PAGE_MAPPING,
    ae_pages,
    ae_project,
)
from PowerBI.builders.pbi_project import build_project

# The 29 section displayNames of the legacy report, in legacy nav order.
LEGACY_PAGE_NAMES = (
    "Domain Data Flows", "Department Analysis", "Location", "Location Risk",
    "Timeline", "Risk Assessment", "Location Domain Data Flows", "File Analysis",
    "DLP Policy Analysis", "AI View", "Graph Domain Data Flows",
    "Subject Heading Word Cloud", "Classifier Focus", "Agent Activity",
    "Executive Overview", "Device", "Folder Data Flows", "Activity Summary Table",
    "Classifier Analysis", "TreeDept", "Classifier Detail",
    "Drill Through Activity", "Drill Through Summary", "DomainDrillThrough",
    "LocationDrillThrough", "Activity Detail", "Summary Activity Detail",
    "USB Breakdown", "User",
)

EXPECTED_PAGE_COUNT = 29


@pytest.fixture(scope="module")
def project_dir(tmp_path_factory: pytest.TempPathFactory) -> Path:
    out = tmp_path_factory.mktemp("ae") / "pbix"
    return build_project(ae_project(), out)


def _model_tmdl_text(project_dir: Path) -> str:
    return "\n".join(
        path.read_text(encoding="utf-8")
        for path in sorted((project_dir / "Model" / "tables").glob("*.tmdl"))
    )


# --- superset: pages ---------------------------------------------------------

def test_every_legacy_page_is_mapped() -> None:
    assert set(LEGACY_PAGE_MAPPING) == set(LEGACY_PAGE_NAMES)


def test_mapping_targets_are_emitted_pages(project_dir: Path) -> None:
    emitted = {path.name for path in (project_dir / "Report" / "sections").iterdir()}
    for legacy, target in LEGACY_PAGE_MAPPING.items():
        assert target in emitted, f"{legacy!r} maps to missing page {target!r}"


def test_expected_page_count(project_dir: Path) -> None:
    pages = ae_pages()
    assert len(pages) == EXPECTED_PAGE_COUNT
    emitted = list((project_dir / "Report" / "sections").iterdir())
    assert len(emitted) == EXPECTED_PAGE_COUNT


def test_pages_in_nav_order(project_dir: Path) -> None:
    folders = [page.folder for page in ae_pages()]
    assert folders == sorted(folders)
    ordinals = {}
    for section_dir in (project_dir / "Report" / "sections").iterdir():
        section = json.loads((section_dir / "section.json").read_text(encoding="utf-8"))
        ordinals[section_dir.name] = section["ordinal"]
    assert [name for name, _ in sorted(ordinals.items(), key=lambda kv: kv[1])] == folders


def test_drillthrough_pages_bound(project_dir: Path) -> None:
    report = json.loads((project_dir / "Report" / "report.json").read_text(encoding="utf-8"))
    drill_pods = [pod for pod in report["pods"] if pod.get("type") == 1]
    drill_pages = [page for page in ae_pages() if page.is_drillthrough()]
    assert len(drill_pods) == len(drill_pages) == 4


# --- superset: measures ------------------------------------------------------

def test_every_legacy_measure_ported(project_dir: Path) -> None:
    assert len(LEGACY_MEASURE_NAMES) == 45
    declared = {measure.name for measure in MEASURES}
    missing = set(LEGACY_MEASURE_NAMES) - declared
    assert not missing, f"legacy measures not declared: {sorted(missing)}"
    tmdl = _model_tmdl_text(project_dir)
    for name in LEGACY_MEASURE_NAMES:
        bare = f"measure {name} =" in tmdl
        quoted = f"measure '{name}' =" in tmdl
        assert bare or quoted, f"measure {name!r} not emitted"


def test_no_today_in_any_dax(project_dir: Path) -> None:
    pattern = re.compile(r"\bTODAY\s*\(")
    for measure in MEASURES:
        assert not pattern.search(measure.dax), f"TODAY() in {measure.name!r}"
    assert not pattern.search(_model_tmdl_text(project_dir))


def test_userelationship_measures_target_inactive_relationships() -> None:
    inactive = {
        (rel.from_table, rel.from_column)
        for rel in schema.model_relationships() if not rel.active
    }
    assert ("fact_activity", "target_location_id") in inactive
    assert ("fact_activity", "originating_domain_id") in inactive
    by_name = {measure.name: measure for measure in MEASURES}
    assert "USERELATIONSHIP ( fact_activity[target_location_id]" in by_name[
        "Target Location Activities"].dax
    assert "USERELATIONSHIP ( fact_activity[originating_domain_id]" in by_name[
        "Originating Domain Activities"].dax


# --- bindings reference real SSOT columns and declared measures --------------

def test_all_visual_fields_resolve() -> None:
    columns = {
        (table.name, column.name)
        for table in schema.model_tables() for column in table.columns
    }
    measure_names = {(measure.table, measure.name) for measure in MEASURES}
    for page in ae_pages():
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


# --- polish: every data visual carries a curated title ------------------------

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


def test_sankey_pages_keep_risk_gates(project_dir: Path) -> None:
    # 410's gate is at SIT grain ([Total SIT Risk]): the word cloud groups by
    # fact_email_detail[subject], which cannot filter fact_activity's
    # [TotalRisk] (see ae_pages_flows.email_subject_cloud_page).
    gated_pages = {
        "330_Location_Domain_Flows": "TotalRisk",
        "340_Folder_Data_Flows": "TotalRisk",
        "350_Domain_Graph": "TotalRisk",
        "410_Email_Subject_Cloud": "Total SIT Risk",
    }
    for folder, gate_measure in gated_pages.items():
        gates = []
        for filters_path in (project_dir / "Report" / "sections" / folder).glob(
                "visualContainers/*/filters.json"):
            for entry in json.loads(filters_path.read_text(encoding="utf-8")):
                condition = entry.get("filter", {}).get("Where", [{}])[0].get(
                    "Condition", {})
                comparison = condition.get("Comparison", {})
                if comparison.get("Left", {}).get("Measure", {}).get(
                        "Property") == gate_measure:
                    gates.append(comparison)
        assert gates, f"{folder}: {gate_measure} gate not ported"


# --- polish: compact dropdown multi-select slicers (T6 owner feedback) --------

def test_slicers_are_compact_dropdown_multiselect(project_dir: Path) -> None:
    found = 0
    for config_path in (project_dir / "Report" / "sections").glob(
            "*/visualContainers/*/config.json"):
        single = json.loads(config_path.read_text(encoding="utf-8"))["singleVisual"]
        if single["visualType"] != "slicer":
            continue
        found += 1
        objects = single["objects"]

        def literal(group: str, prop: str) -> str:
            return objects[group][0]["properties"][prop]["expr"]["Literal"]["Value"]

        assert literal("data", "mode") == "'Dropdown'"
        assert literal("selection", "singleSelect") == "false"
        # strictSingleSelect false = "Multi-select with CTRL" OFF (plain-click
        # checkbox multi-select)
        assert literal("selection", "strictSingleSelect") == "false"
        assert literal("selection", "selectAllCheckboxEnabled") == "true"
        assert literal("items", "textSize") == "10D"
        assert literal("header", "show") == "true"
        assert literal("header", "textSize") == "10D"
    assert found, "no slicers emitted"


# --- categorical filter values match the data domain (T6 Bug B) ---------------

# Workload values the star ETL passes through verbatim from the AE export
# (dim_workload.workload). Endpoint is included for endpoint-bearing tenants.
KNOWN_WORKLOADS = {
    "Exchange", "SharePoint", "OneDrive", "MicrosoftTeams", "Copilot",
    "Endpoint",
}


def _categorical_filters(page) -> list[tuple[str, str, list[str]]]:
    """(entity, property, values) for every Categorical In filter on a page."""
    entries = list(page.page_filters)
    for visual in page.visuals:
        entries.extend(visual.filters)
    found = []
    for entry in entries:
        if entry.get("type") != "Categorical" or "filter" not in entry:
            continue
        where = entry["filter"]["Where"][0]["Condition"]
        if "In" not in where:
            continue
        column = where["In"]["Expressions"][0]["Column"]["Property"]
        entity = entry["filter"]["From"][0]["Entity"]
        values = [
            value[0]["Literal"]["Value"].strip("'")
            for value in where["In"]["Values"]
        ]
        found.append((entity, column, values))
    return found


def test_workload_filter_values_are_known() -> None:
    """Stale-value guard: workload include-lists must use real dim values
    (the legacy 'Sensitive Files Created'/'Sensitive USB'/'Sensitive Cloud'
    activity lists matched zero rows in cloud-DLP exports)."""
    seen = 0
    for page in ae_pages():
        for entity, column, values in _categorical_filters(page):
            if (entity, column) != ("dim_workload", "workload"):
                continue
            seen += 1
            unknown = set(values) - KNOWN_WORKLOADS
            assert not unknown, f"{page.folder}: unknown workloads {unknown}"
    assert seen >= 5  # the five 500_Activity_Detail channel tables


def test_usb_activity_filter_carries_both_naming_styles() -> None:
    """370_USB_Breakdown must match removable-media activities whether the
    export carries humanized display strings or raw enum names."""
    usb_page = next(page for page in ae_pages()
                    if page.folder == "370_USB_Breakdown")
    activity_filters = [
        values for entity, column, values in _categorical_filters(usb_page)
        if (entity, column) == ("dim_activity_type", "activity")
    ]
    assert activity_filters, "USB page lost its activity filter"
    values = set(activity_filters[0])
    assert "File copied to removable media" in values
    assert "FileCopiedToRemovableMedia" in values


# --- division as the primary org lens (T6 polish 3) ---------------------------

def _page(folder: str):
    return next(page for page in ae_pages() if page.folder == folder)


def _visual(folder: str, seed: str):
    return next(v for v in _page(folder).visuals if v.seed == seed)


def _slicer_columns(page) -> set[tuple[str, str]]:
    return {
        (v.fields[0].table, v.fields[0].name)
        for v in page.visuals if v.visual_type == "slicer"
    }


def test_standard_slicer_band_uses_division() -> None:
    """Department was a single wall-to-wall value on this tenant; the standard
    band's org slicer is dim_user.division."""
    page = _page("010_Activity_Summary")
    slicers = _slicer_columns(page)
    assert ("dim_user", "division") in slicers
    assert ("dim_department", "department") not in slicers


def test_org_pages_carry_region_slicer() -> None:
    for folder in ("200_Department_Analysis", "210_Department_Treemap",
                   "220_User_Investigation"):
        assert ("dim_user", "region") in _slicer_columns(_page(folder)), folder


def test_executive_division_visuals_resolve_via_dim_user() -> None:
    treemap = _visual("000_Executive_Overview", "exec-treemap-division")
    bound = {(field.table, field.name) for field in treemap.fields}
    assert ("dim_user", "division") in bound
    # risk pressure at fact grain (agg_department cannot reach dim_user)
    assert ("fact_activity_sit", "Weighted Risk Score") in bound
    summary = _visual("000_Executive_Overview", "exec-table-division")
    assert ("dim_user", "division") in {(f.table, f.name) for f in summary.fields}


def test_division_visuals_never_mix_agg_department_measures() -> None:
    """A dim_user.division/region grouping cannot filter agg_department_sit_day
    (no user_id on the agg): any visual binding those columns must take its
    measures from elsewhere."""
    for page in ae_pages():
        for visual in page.visuals:
            tables = {field.table for field in visual.fields if field.kind == "column"}
            if "dim_user" not in tables:
                continue
            measure_homes = {
                field.table for field in visual.fields if field.kind == "measure"
            }
            assert "agg_department_sit_day" not in measure_homes, (
                f"{page.folder}/{visual.seed}: agg_department measure cannot "
                "respond to a dim_user grouping")


def test_department_lens_kept_for_completeness() -> None:
    """Division is primary, but Department must remain reachable somewhere."""
    dept_bindings = []
    for page in ae_pages():
        for visual in page.visuals:
            for field in visual.fields:
                if (field.table, field.name) == ("dim_department", "department"):
                    dept_bindings.append((page.folder, visual.seed))
    assert dept_bindings, "no Department-bound visual left in the report"


def test_user_investigation_org_columns_and_flagged_table() -> None:
    evidence = _visual("220_User_Investigation", "userpage-table-evidence")
    bound = {(field.table, field.name) for field in evidence.fields}
    assert ("dim_user", "job_title") in bound
    assert ("dim_user", "is_leaver") in bound
    flagged = _visual("220_User_Investigation", "userpage-table-flagged")
    flagged_bound = {(field.table, field.name) for field in flagged.fields}
    assert ("dim_user", "is_generic_account") in flagged_bound
    assert ("fact_activity", "Flagged Account Activities") in flagged_bound
    # gated to accounts that actually generated activity
    gates = [
        entry for entry in flagged.filters
        if entry.get("filter", {}).get("Where", [{}])[0].get("Condition", {})
        .get("Comparison", {}).get("Left", {}).get("Measure", {})
        .get("Property") == "Flagged Account Activities"
    ]
    assert gates, "flagged-accounts table lost its activity gate"


def test_leaver_kpi_on_executive() -> None:
    card = _visual("000_Executive_Overview", "exec-card-leavers")
    assert {(field.table, field.name) for field in card.fields} == {
        ("fact_activity", "Leaver Activities")
    }


# --- determinism --------------------------------------------------------------

def test_two_builds_are_byte_identical(project_dir: Path, tmp_path: Path) -> None:
    second = build_project(ae_project(), tmp_path / "pbix")
    differences: list[str] = []

    def _collect(cmp: filecmp.dircmp, prefix: str = "") -> None:
        differences.extend(f"{prefix}{name}" for name in cmp.diff_files)
        differences.extend(f"{prefix}{name}" for name in cmp.left_only + cmp.right_only)
        for sub_name, sub_cmp in cmp.subdirs.items():
            _collect(sub_cmp, f"{prefix}{sub_name}/")

    _collect(filecmp.dircmp(project_dir, second))
    assert differences == [], f"non-deterministic output: {differences}"
