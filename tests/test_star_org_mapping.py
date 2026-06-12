"""Unit tests for the GAL/department loader and the org-mapping config engine.

Covers the Bug A mail-alias/department-casing behavior (T6 polish 2), the org
enrichment fields (T6 polish 3) and the config-driven sourcing engine that
replaced the hardcoded tenant-specific rules (T6 polish 4). Convert-level integration
tests (resolution order, manifest provenance) live in test_star_convert.py.
"""

from __future__ import annotations

import json
from pathlib import Path

import pyarrow.parquet as pq
import pytest

from test_star_convert import _make_export, _make_gal, _make_workbook

from parquet_builder.star import convert as convert_module
from parquet_builder.star.convert import convert, resolve_org_mapping
from parquet_builder.star.enrich import (
    EnrichmentError,
    RiskLookup,
    load_department_mapping,
)
from parquet_builder.star.finalize import write_dimensions
from parquet_builder.star.org_mapping import (
    default_org_mapping,
    load_org_mapping,
    ou_contains,
    ou_name_in,
    ou_under,
    parse_dn_ous,
)
from parquet_builder.star.registry import IdRegistry


def _make_org_mapping_config(path: Path) -> Path:
    """Zava-style override: Division from CompanyName, Department fallback.

    Deliberately partial — every other field must keep its built-in default."""
    path.write_text(json.dumps({
        "_Description": "test org mapping",
        "Division": {"Source": "CompanyName", "Fallback": "Department"},
    }), encoding="utf-8")
    return path


# --- department resolution (Bug A: GAL mail aliases + case-canonical names) --

def _make_alias_gal(path: Path) -> Path:
    path.write_text(
        "UserPrincipalName,Mail,Department\n"
        "alice@zava.example.com,Alice.A@mail.example.com,ZAVA\n"
        "bob@zava.example.com,Bob.B@mail.example.com,ZAVA\n"
        "carol@zava.example.com,,zava\n"
        ",dave@mail.example.com,Ops\n",
        encoding="utf-8",
    )
    return path


def test_department_mapping_registers_mail_aliases(tmp_path) -> None:
    mappings = load_department_mapping(_make_alias_gal(tmp_path / "GAL_Clean.csv"))
    # UPN keys are primary; the same row's Mail address is an alias entry
    assert "is_alias" not in mappings["ALICE@ZAVA.EXAMPLE.COM"]
    assert mappings["ALICE.A@MAIL.EXAMPLE.COM"]["is_alias"] is True
    assert mappings["ALICE.A@MAIL.EXAMPLE.COM"]["department"] == "ZAVA"
    # rows with only a Mail value keep it as the primary key
    assert "is_alias" not in mappings["DAVE@MAIL.EXAMPLE.COM"]
    assert mappings["DAVE@MAIL.EXAMPLE.COM"]["department"] == "Ops"
    # department casing canonicalized to the dominant variant in the file
    assert mappings["CAROL@ZAVA.EXAMPLE.COM"]["department"] == "ZAVA"


def test_activity_user_resolves_department_via_mail_and_case(tmp_path) -> None:
    mappings = load_department_mapping(_make_alias_gal(tmp_path / "GAL_Clean.csv"))
    registry = IdRegistry()
    # activity identifies the user by primary SMTP address, in different case;
    # the GAL row is keyed by UPN — both must land on the same mapped department
    _, dept_via_mail = registry.get_user("alice.a@MAIL.example.com", mappings)
    _, dept_via_upn = registry.get_user("ALICE@ZAVA.EXAMPLE.COM", mappings)
    assert dept_via_mail == dept_via_upn
    row = registry.department_rows[dept_via_mail]
    assert row["department"] == "ZAVA"
    assert row["is_mapped"] is True


def test_department_case_variants_share_one_dim_row() -> None:
    registry = IdRegistry()
    id_upper = registry.get_department({"department": "ZAVA", "is_mapped": True})
    id_lower = registry.get_department({"department": "zava", "is_mapped": True})
    assert id_upper == id_lower
    assert len(registry.department_rows) == 1
    # display casing = first seen (loader canonicalizes to dominant casing)
    assert registry.department_rows[id_upper]["department"] == "ZAVA"


def test_dim_user_seeding_skips_mail_alias_keys(tmp_path) -> None:
    mappings = load_department_mapping(_make_alias_gal(tmp_path / "GAL_Clean.csv"))
    registry = IdRegistry()
    write_dimensions(tmp_path / "out", registry, mappings, {}, RiskLookup(), set())
    upns = {row["user_upn"] for row in pq.read_table(
        tmp_path / "out" / "dim_user.parquet").to_pylist()}
    # one dim_user row per person (UPN or mail-only primary), none per alias
    assert upns == {
        "ALICE@ZAVA.EXAMPLE.COM", "BOB@ZAVA.EXAMPLE.COM",
        "CAROL@ZAVA.EXAMPLE.COM", "DAVE@MAIL.EXAMPLE.COM",
    }


# --- org enrichment fields (T6 polish 3, now config-driven) -------------------

def _make_org_gal(path: Path) -> Path:
    path.write_text(
        "UserPrincipalName,Department,CompanyName,JobTitle,OnPremisesDN\n"
        # CompanyName beats Department under the Zava-style config
        "fiona@zava.example.com,ZAVA,ZFR,Firefighter,"
        '"CN=Fiona,OU=Users,OU=South East,OU=Regions,OU=Win7MOE,OU=MOE,DC=corp,DC=internal"\n'
        # lower-case CompanyName variant: canonicalized to the dominant casing
        "gary@zava.example.com,ZAVA,zfr,,"
        '"CN=Gary,OU=Users,OU=Far Northern,OU=Regions,OU=MOE,DC=corp,DC=internal"\n'
        # leaver: Leavers OU, no Regions OU anywhere
        "wayne@zava.example.com,ZAVA,,Regional Director,"
        '"CN=Wayne,OU=ZAVA Leavers,OU=Leavers,OU=Org Users,DC=corp,DC=internal"\n'
        # generic-account pool directly under Regions
        "icc.info@zava.example.com,ZAVA,,,"
        '"CN=ICC Info,OU=Generic Accounts,OU=Regions,OU=Win7MOE,OU=MOE,DC=corp,DC=internal"\n'
        # SharedUsers pool also counts as generic
        "462.alpha@zava.example.com,ZAVA,,,"
        '"CN=462 Alpha,OU=SharedUsers,OU=Regions,OU=MOE,DC=corp,DC=internal"\n'
        # service OU with no Regions ancestor -> region Unknown
        "svc.sync@zava.example.com,ZAVA,,,"
        '"CN=Sync,OU=Office365Sync,OU=Win7MOE,OU=MOE,DC=corp,DC=internal"\n'
        # no Department, no CompanyName, no DN -> all Unknown
        "blank@zava.example.com,,,,\n",
        encoding="utf-8",
    )
    return path


def test_division_company_name_with_department_fallback(tmp_path) -> None:
    """Zava-style config: Division Source CompanyName, Fallback Department."""
    org = load_org_mapping(_make_org_mapping_config(tmp_path / "org-mapping.json"))
    mappings = load_department_mapping(_make_org_gal(tmp_path / "GAL_Clean.csv"), org)
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["user_division"] == "ZFR"
    # casing canonicalized to the dominant variant across the file
    assert mappings["GARY@ZAVA.EXAMPLE.COM"]["user_division"] == "ZFR"
    # no CompanyName -> Department fallback
    assert mappings["WAYNE@ZAVA.EXAMPLE.COM"]["user_division"] == "ZAVA"
    # neither -> Unknown
    assert mappings["BLANK@ZAVA.EXAMPLE.COM"]["user_division"] == "Unknown"


def test_default_mapping_division_mirrors_department(tmp_path) -> None:
    """No config: vanilla defaults — Division is just Department again."""
    mappings = load_department_mapping(_make_org_gal(tmp_path / "GAL_Clean.csv"))
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["user_division"] == "ZAVA"
    assert mappings["WAYNE@ZAVA.EXAMPLE.COM"]["user_division"] == "ZAVA"
    assert mappings["BLANK@ZAVA.EXAMPLE.COM"]["user_division"] == "Unknown"
    assert mappings["BLANK@ZAVA.EXAMPLE.COM"]["department"] == "Unmapped"


def test_region_parsed_from_dn(tmp_path) -> None:
    # built-in default rule: ou_under 'Regions' on OnPremisesDN
    mappings = load_department_mapping(_make_org_gal(tmp_path / "GAL_Clean.csv"))
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["user_region"] == "South East"
    assert mappings["GARY@ZAVA.EXAMPLE.COM"]["user_region"] == "Far Northern"
    # account-pool OUs under Regions are regions verbatim
    assert mappings["ICC.INFO@ZAVA.EXAMPLE.COM"]["user_region"] == "Generic Accounts"
    assert mappings["462.ALPHA@ZAVA.EXAMPLE.COM"]["user_region"] == "SharedUsers"
    # leaver / service OUs have no Regions ancestor
    assert mappings["WAYNE@ZAVA.EXAMPLE.COM"]["user_region"] == "Unknown"
    assert mappings["SVC.SYNC@ZAVA.EXAMPLE.COM"]["user_region"] == "Unknown"
    assert mappings["BLANK@ZAVA.EXAMPLE.COM"]["user_region"] == "Unknown"


def test_leaver_and_generic_account_flags(tmp_path) -> None:
    mappings = load_department_mapping(_make_org_gal(tmp_path / "GAL_Clean.csv"))
    assert mappings["WAYNE@ZAVA.EXAMPLE.COM"]["is_leaver"] is True
    assert mappings["WAYNE@ZAVA.EXAMPLE.COM"]["is_generic_account"] is False
    assert mappings["ICC.INFO@ZAVA.EXAMPLE.COM"]["is_generic_account"] is True
    assert mappings["462.ALPHA@ZAVA.EXAMPLE.COM"]["is_generic_account"] is True
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["is_leaver"] is False
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["is_generic_account"] is False
    assert mappings["FIONA@ZAVA.EXAMPLE.COM"]["job_title"] == "Firefighter"


def test_dn_parse_modes_edge_cases() -> None:
    # escaped comma inside an OU value survives the component split
    assert parse_dn_ous(
        r"CN=X,OU=Fire\, Rescue,OU=Regions,DC=corp"
    ) == ["Fire, Rescue", "Regions"]
    # ou_under: OU directly under the parent (DN is leaf-first)
    assert ou_under(r"CN=X,OU=Fire\, Rescue,OU=Regions,DC=corp", "Regions") == "Fire, Rescue"
    # parent as the leaf OU has no child -> None
    assert ou_under("CN=X,OU=Regions,OU=MOE,DC=corp", "Regions") is None
    # parent absent -> None
    assert ou_under("CN=X,OU=Office365Sync,OU=MOE,DC=corp", "Regions") is None
    # case-insensitive OU matching, on both sides
    assert ou_under("CN=X,OU=Users,OU=Central,OU=REGIONS,DC=corp", "regions") == "Central"
    # ou_contains: substring match against any OU
    assert ou_contains("CN=X,OU=zava leavers,OU=LEAVERS,DC=corp", "leaver") is True
    assert ou_contains("CN=X,OU=Users,OU=Central,OU=Regions,DC=corp", "leaver") is False
    # ou_name_in: exact OU-name membership, case-insensitive
    assert ou_name_in("CN=X,OU=GENERIC ACCOUNTS,OU=Regions,DC=corp",
                      ("Generic Accounts", "SharedUsers")) is True
    assert ou_name_in("CN=X,OU=Generic,OU=Regions,DC=corp",
                      ("Generic Accounts", "SharedUsers")) is False
    # empty / missing DNs
    assert ou_under(None, "Regions") is None
    assert ou_under("", "Regions") is None
    assert ou_contains(None, "leaver") is False
    assert ou_name_in("", ("Generic Accounts",)) is False


def test_activity_only_user_gets_unknown_org(tmp_path) -> None:
    mappings = load_department_mapping(_make_org_gal(tmp_path / "GAL_Clean.csv"))
    registry = IdRegistry()
    user_id, _ = registry.get_user("stranger@elsewhere.example.com", mappings)
    row = registry.user_rows[user_id]
    assert row["division"] == "Unknown"
    assert row["region"] == "Unknown"
    assert row["job_title"] is None
    assert row["is_leaver"] is False
    assert row["is_generic_account"] is False


# --- org-mapping config engine (T6 polish 4) ----------------------------------

def test_partial_config_keeps_defaults_for_omitted_fields(tmp_path) -> None:
    org = load_org_mapping(_make_org_mapping_config(tmp_path / "org-mapping.json"))
    assert org.fields["Division"].from_config is True
    defaults = default_org_mapping()
    for name in ("Department", "Region", "JobTitle", "IsLeaver", "IsGenericAccount"):
        assert org.fields[name] == defaults.fields[name]
        assert org.fields[name].from_config is False


def test_explicit_org_mapping_path_must_exist(tmp_path) -> None:
    with pytest.raises(EnrichmentError, match="--org-mapping does not exist"):
        resolve_org_mapping(tmp_path / "nope.json")


# --- resolution order at convert level -----------------------------------------

def test_defaults_only_convert_run(tmp_path, monkeypatch) -> None:
    """No config anywhere: manifest says 'defaults', division mirrors department."""
    monkeypatch.setattr(convert_module, "_DEFAULT_ORG_MAPPING", tmp_path / "absent.json")
    export = _make_export(tmp_path / "Export-20260504-150000")
    out = tmp_path / "out-defaults"
    manifest = convert(
        export, output_dir=out,
        risk_workbook=_make_workbook(tmp_path / "SIT-Risk-Analysis-test.xlsx"),
        department_csv=_make_gal(tmp_path / "GAL_Clean.csv"),
        sit_exclusions=None,
    )
    assert manifest["org_mapping"]["source"] == "defaults"
    assert manifest["org_mapping"]["fields"]["Division"] == {"Source": "Department"}
    users = {row["user_upn"]: row for row in pq.read_table(
        out / "dim_user.parquet").to_pylist()}
    alice = users["ALICE@CONTOSO.COM"]
    assert alice["division"] == "Dept A"  # mirrors department, NOT CompanyName
    assert alice["region"] == "Central"   # DN defaults still apply
    assert alice["job_title"] == "Data Scientist"
    assert users["BOB.GALONLY@CONTOSO.COM"]["division"] == "Dept B"


def test_local_config_auto_detected(tmp_path, monkeypatch) -> None:
    local = _make_org_mapping_config(tmp_path / "AEStarOrgMapping.local.json")
    monkeypatch.setattr(convert_module, "_DEFAULT_ORG_MAPPING", local)
    export = _make_export(tmp_path / "Export-20260504-160000")
    out = tmp_path / "out-local"
    manifest = convert(
        export, output_dir=out,
        risk_workbook=_make_workbook(tmp_path / "SIT-Risk-Analysis-test.xlsx"),
        department_csv=_make_gal(tmp_path / "GAL_Clean.csv"),
        sit_exclusions=None,
    )
    assert manifest["org_mapping"]["source"] == str(local)
    users = {row["user_upn"]: row for row in pq.read_table(
        out / "dim_user.parquet").to_pylist()}
    assert users["ALICE@CONTOSO.COM"]["division"] == "DIV-ONE"  # CompanyName via config


def test_cli_org_mapping_overrides_local_config(tmp_path, monkeypatch) -> None:
    local = tmp_path / "AEStarOrgMapping.local.json"
    local.write_text(json.dumps({"Division": {"Source": "Department"}}), encoding="utf-8")
    monkeypatch.setattr(convert_module, "_DEFAULT_ORG_MAPPING", local)
    explicit = _make_org_mapping_config(tmp_path / "explicit-org-mapping.json")
    export = _make_export(tmp_path / "Export-20260504-170000")
    out = tmp_path / "out-explicit"
    manifest = convert(
        export, output_dir=out, org_mapping=explicit,
        risk_workbook=_make_workbook(tmp_path / "SIT-Risk-Analysis-test.xlsx"),
        department_csv=_make_gal(tmp_path / "GAL_Clean.csv"),
        sit_exclusions=None,
    )
    assert manifest["org_mapping"]["source"] == str(explicit)
    users = {row["user_upn"]: row for row in pq.read_table(
        out / "dim_user.parquet").to_pylist()}
    assert users["ALICE@CONTOSO.COM"]["division"] == "DIV-ONE"


def test_flag_field_parse_none_reads_truthy_column(tmp_path) -> None:
    gal = tmp_path / "GAL_Clean.csv"
    gal.write_text(
        "UserPrincipalName,Department,HasLeft\n"
        "yes@x.example.com,Ops,TRUE\n"
        "no@x.example.com,Ops,false\n"
        "blank@x.example.com,Ops,\n",
        encoding="utf-8",
    )
    config = tmp_path / "org.json"
    config.write_text(json.dumps({"IsLeaver": {"Source": "HasLeft"}}), encoding="utf-8")
    mappings = load_department_mapping(gal, load_org_mapping(config))
    assert mappings["YES@X.EXAMPLE.COM"]["is_leaver"] is True
    assert mappings["NO@X.EXAMPLE.COM"]["is_leaver"] is False
    assert mappings["BLANK@X.EXAMPLE.COM"]["is_leaver"] is False


def test_value_field_ou_under_with_custom_arg(tmp_path) -> None:
    gal = tmp_path / "GAL_Clean.csv"
    gal.write_text(
        "UserPrincipalName,Department,OnPremisesDN\n"
        'jo@x.example.com,Ops,"CN=Jo,OU=Users,OU=Metro,OU=Branches,DC=corp"\n',
        encoding="utf-8",
    )
    config = tmp_path / "org.json"
    config.write_text(json.dumps({
        "Region": {"Source": "OnPremisesDN", "Parse": "ou_under", "Arg": "Branches"},
    }), encoding="utf-8")
    mappings = load_department_mapping(gal, load_org_mapping(config))
    assert mappings["JO@X.EXAMPLE.COM"]["user_region"] == "Metro"


@pytest.mark.parametrize("payload,match", [
    ("{not json", "not valid JSON"),
    ('["Division"]', "expected a JSON object"),
    ('{"Dept": {"Source": "Department"}}', "unknown field 'Dept'"),
    ('{"Division": "CompanyName"}', "expected an object"),
    ('{"Division": {"Source": "CompanyName", "Sauce": "x"}}', "unknown key"),
    ('{"Division": {"Fallback": "Department"}}', "'Source' is required"),
    ('{"Division": {"Source": ""}}', "'Source' is required"),
    ('{"Region": {"Source": "OnPremisesDN", "Parse": "ou_below", "Arg": "X"}}',
     "unknown Parse mode 'ou_below'"),
    ('{"Region": {"Source": "OnPremisesDN", "Parse": "ou_contains", "Arg": "X"}}',
     "unknown Parse mode 'ou_contains' for value field"),
    ('{"IsLeaver": {"Source": "OnPremisesDN", "Parse": "ou_under", "Arg": "X"}}',
     "unknown Parse mode 'ou_under' for flag field"),
    ('{"Region": {"Source": "OnPremisesDN", "Parse": "ou_under"}}',
     "requires 'Arg' to be a non-empty string"),
    ('{"IsGenericAccount": {"Source": "OnPremisesDN", "Parse": "ou_name_in", "Arg": "X"}}',
     "non-empty list"),
    ('{"IsGenericAccount": {"Source": "OnPremisesDN", "Parse": "ou_name_in", "Arg": []}}',
     "non-empty list"),
    ('{"Division": {"Source": "CompanyName", "Arg": "X"}}',
     "'Arg' is only valid with a Parse mode"),
    ('{"IsLeaver": {"Source": "HasLeft", "Fallback": "Department"}}',
     "'Fallback' is only valid on value fields"),
])
def test_malformed_org_mapping_hard_errors(tmp_path, payload, match) -> None:
    config = tmp_path / "org.json"
    config.write_text(payload, encoding="utf-8")
    with pytest.raises(EnrichmentError, match=match):
        load_org_mapping(config)


def test_config_referencing_unknown_gal_column_hard_errors(tmp_path) -> None:
    gal = _make_org_gal(tmp_path / "GAL_Clean.csv")
    config = tmp_path / "org.json"
    config.write_text(json.dumps({
        "Division": {"Source": "CompanyNameX", "Fallback": "Departmint"},
    }), encoding="utf-8")
    with pytest.raises(EnrichmentError) as excinfo:
        load_department_mapping(gal, load_org_mapping(config))
    message = str(excinfo.value)
    assert "Division.Source column 'CompanyNameX'" in message
    assert "Division.Fallback column 'Departmint'" in message
    assert "Available columns" in message
    assert "CompanyName" in message  # the real header is listed to help fix it


def test_default_rules_tolerate_missing_gal_columns(tmp_path) -> None:
    """An unconfigured tenant without DN/JobTitle/CompanyName columns still maps."""
    gal = tmp_path / "GAL_Clean.csv"
    gal.write_text(
        "UserPrincipalName,Department\n"
        "amy@x.example.com,Finance\n",
        encoding="utf-8",
    )
    mappings = load_department_mapping(gal)  # built-in defaults
    amy = mappings["AMY@X.EXAMPLE.COM"]
    assert amy["department"] == "Finance"
    assert amy["user_division"] == "Finance"
    assert amy["user_region"] == "Unknown"
    assert amy["job_title"] is None
    assert amy["is_leaver"] is False
    assert amy["is_generic_account"] is False
