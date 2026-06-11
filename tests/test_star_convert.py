"""End-to-end tests for the AE star-schema v6 converter on synthetic fixtures."""

from __future__ import annotations

import json
from pathlib import Path

import pyarrow.parquet as pq
import pytest

from parquet_builder.star import convert as convert_module
from parquet_builder.star import schema
from parquet_builder.star.convert import convert
from parquet_builder.star.enrich import EnrichmentError
from parquet_builder.star.keys import stable_int_id

SIT_GUID_OK = "11111111-1111-1111-1111-111111111111"
SIT_GUID_EXCLUDED = "22222222-2222-2222-2222-222222222222"
# Name-resolution chain fixtures (none of these have a workbook GUID row):
SIT_GUID_RAWNAME = "33333333-3333-3333-3333-333333333333"      # name only in raw payload
SIT_GUID_TENANT = "44444444-4444-4444-4444-444444444444"       # name only in tenant map
SIT_GUID_UNKNOWN = "55555555-5555-5555-5555-555555555555"      # named nowhere -> GUID label
SIT_GUID_BRIDGE = "66666666-6666-6666-6666-666666666666"       # tenant name matches workbook slug row
SIT_GUID_TENANT_EXCL = "77777777-7777-7777-7777-777777777777"  # tenant name on the exclusion list

RID_A = "aaaaaaaa-0000-0000-0000-000000000001"
RID_C = "cccccccc-0000-0000-0000-000000000003"
RID_D = "dddddddd-0000-0000-0000-000000000004"
RID_E = "eeeeeeee-0000-0000-0000-000000000005"
RID_F = "ffffffff-0000-0000-0000-000000000006"


def _record_a() -> dict:
    """DLP record: JSON-string-encoded SIT data, email, drift field."""
    return {
        "RecordIdentity": RID_A,
        "Activity": "DLP rule matched",
        "ActivityId": "DlpRuleMatch",
        "Happened": "2026-05-01T10:00:00Z",
        "User": "alice@contoso.com",
        "Workload": "Exchange",
        "DataPlatform": "Purview",
        "UserType": "Regular",
        "FilePath": "1779222579797",
        "ItemName": "Quarterly Report",  # no dot/URL: forces email-receiver domain fallback
        "FileSize": 1024,
        "FileExtension": "docx",
        # nested field arrives as a JSON-encoded STRING (robustness test)
        "SensitiveInfoTypeData": json.dumps([
            {"SensitiveInfoTypeId": SIT_GUID_OK, "Count": 3, "Confidence": 85,
             "ClassifierType": "Content", "UniqueCount": 2},
        ]),
        "SensitiveInfoTypeBucketsData": [
            {"Id": SIT_GUID_OK, "Low": 0, "Medium": 1, "High": 2, "ClassifierType": "Content"},
        ],
        "PolicyMatchInfo": {
            "PolicyId": "pol-1", "PolicyName": "Policy One", "PolicyMode": "Enable",
            "RuleId": "rule-1", "RuleName": "Rule One", "RuleActions": ["BlockAccess"],
        },
        "EmailInfo": {
            "Sender": "alice@contoso.com",
            "Receivers": ["bob@contoso.com", "ext@evil.com"],
            "Subject": "Q3 numbers",
            "MessageID": "<msg-1@contoso.com>",
        },
        "AttachmentDetails": [
            {"Name": "a.docx", "Size": 11, "Labels": None},
            {"Name": "b.xlsx", "Size": 22, "Labels": None},
        ],
        "FutureUnknownField": "drift-me",
    }


def _record_c() -> dict:
    """DLP record whose only SIT is on the exclusion list."""
    return {
        "RecordIdentity": RID_C,
        "Activity": "DLP rule matched",
        "ActivityId": "DlpRuleMatch",
        "Happened": "2026-05-01T11:00:00Z",
        "User": "alice@contoso.com",
        "Workload": "SharePoint",
        "DataPlatform": "Purview",
        "UserType": "Regular",
        "FilePath": "https://contoso.sharepoint.com/sites/HR/doc.pdf",
        "ItemName": "doc.pdf",
        "FileSize": 2048,
        "SensitiveInfoTypeData": [
            {"SensitiveInfoTypeId": SIT_GUID_EXCLUDED, "Count": 4, "Confidence": 75},
        ],
        "PolicyMatchInfo": {"PolicyId": "pol-1", "RuleId": "rule-1", "PolicyName": "Policy One"},
    }


def _record_d() -> dict:
    """Copilot interaction (lives on a single-dict Records page: F2 shape)."""
    return {
        "RecordIdentity": RID_D,
        "Activity": "Copilot Interaction",
        "ActivityId": "CopilotInteraction",
        "Happened": "2026-05-04T09:30:00Z",
        "User": "alice@contoso.com",
        "Workload": "Copilot",
        "DataPlatform": "PurviewForAI",
        "UserType": "Regular",
        "PurviewAIAppLocation": "Outlook",
        "AppIdentity": "Copilot.M365Copilot.Apps",
        "AppIdentityCategory": "Copilot",
        "AppIdentityGroup": "Copilot.M365Copilot",
        "PurviewAIAppName": "Copilot.M365Copilot.Apps",
        "EnrichedCopilotThreadOrCorrelationId": "19:thread",
        "EnrichedLLMMessageIds": ["1", "2"],
        "HasWebsearchQuery": True,
        "AreFilesReferenced": False,
        "AreSensitiveFilesReferenced": False,
        "CopilotEventData": {"AppHost": "Outlook", "ThreadId": "19:thread"},
        "AccessedResources": [{"SiteUrl": "https://outlook.office365.com/x"}],
    }


def _record_e() -> dict:
    return {
        "RecordIdentity": RID_E,
        "Activity": "DLP rule matched",
        "ActivityId": "DlpRuleMatch",
        "Happened": "2026-05-02T08:00:00Z",
        "User": "alice@contoso.com",
        "Workload": "Exchange",
        "UserType": "Regular",
        "PolicyMatchInfo": {"PolicyId": "pol-1", "RuleId": "rule-1"},
        "EmailInfo": {},  # empty payload: must NOT produce an email detail row
    }


def _record_f() -> dict:
    """SIT name-resolution chain record: every detection GUID is missing from
    the workbook; names come from the raw payload, the tenant map, or nowhere."""
    return {
        "RecordIdentity": RID_F,
        "Activity": "DLP rule matched",
        "ActivityId": "DlpRuleMatch",
        "Happened": "2026-05-02T09:00:00Z",
        "User": "alice@contoso.com",
        "Workload": "SharePoint",
        "UserType": "Regular",
        "FilePath": "https://contoso.sharepoint.com/sites/Ops/plan.xlsx",
        "ItemName": "plan.xlsx",
        "SensitiveInfoTypeData": [
            # raw payload carries a display name (and the tenant map carries a
            # DIFFERENT one: the raw payload must win the chain)
            {"SensitiveInfoTypeId": SIT_GUID_RAWNAME,
             "SensitiveInfoTypeName": "Raw Payload SIT",
             "Count": 2, "Confidence": 85, "ClassifierType": "Content"},
            {"SensitiveInfoTypeId": SIT_GUID_TENANT, "Count": 1, "Confidence": 75},
            {"SensitiveInfoTypeId": SIT_GUID_UNKNOWN, "Count": 1, "Confidence": 65},
            {"SensitiveInfoTypeId": SIT_GUID_BRIDGE, "Count": 5, "Confidence": 85},
            {"SensitiveInfoTypeId": SIT_GUID_TENANT_EXCL, "Count": 7, "Confidence": 85},
        ],
        "PolicyMatchInfo": {"PolicyId": "pol-1", "RuleId": "rule-1"},
    }


def _make_export(root: Path) -> Path:
    ae = root / "Data" / "ActivityExplorer"

    day1 = ae / "2026-05-01"
    day1.mkdir(parents=True)
    (day1 / "Page-001.json").write_text(json.dumps({
        "PageNumber": 1,
        "ExportTimestamp": "2026-05-01T23:00:00Z",
        "RecordCount": 3,
        "WaterMark": "wm-day1",
        "Records": [_record_a(), _record_a(), _record_c()],  # 2nd A = duplicate
    }), encoding="utf-8")

    day2 = ae / "2026-05-02"
    day2.mkdir(parents=True)
    (day2 / "Page-001.jsonl").write_text(
        # blank line in the middle: the loader must tolerate it
        json.dumps(_record_e()) + "\n\n" + json.dumps(_record_f()) + "\n",
        encoding="utf-8",
    )

    # day 2026-05-03 has no data: dim_date must still be continuous
    day4 = ae / "2026-05-04"
    day4.mkdir(parents=True)
    (day4 / "Page-001.json").write_text(json.dumps({
        "PageNumber": 1,
        "ExportTimestamp": "2026-05-04T23:00:00Z",
        "RecordCount": 1,
        "WaterMark": "wm-day4",
        "Records": _record_d(),  # single dict, NOT a list (F2 regression shape)
    }), encoding="utf-8")

    return root


def _make_workbook(path: Path) -> Path:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "SIT Risk Analysis"
    headers = [
        "SIT Name", "GUID / Slug", "Category", "Risk Description",
        "Risk Rating (1-10)", "Reference URL", "Australian PSPF Classification",
        "QGISCF", "QGISCF DLM", "Label Code", "Classifier Type", "Source",
        "Jurisdictions", "Scope", "Confidence", "Classification Tier",
        "Generic Classification", "Generic DLM",
    ]
    ws.append(headers)
    ws.append(["Test SIT One", SIT_GUID_OK, "Financial", "Money things", 8,
               "https://ref", "OFFICIAL", "QG-1", "DLM-1", "LC1", "Content",
               "Microsoft Built-in", "AU", "Tenant", "High", "Tier 1", "GC", "GD"])
    ws.append(["Excluded SIT", SIT_GUID_EXCLUDED, "Noise", "Excluded by report", 5,
               None, None, None, None, None, None, "Microsoft Built-in",
               None, None, None, None, None, None])
    ws.append(["Custom Only", "custom-only-slug", "Custom", "Never observed", 3,
               None, None, None, None, None, None, "Custom",
               None, None, None, None, None, None])
    ws.append(["Bridge Target", "bridge-target-slug", "Custom", "Bridged via tenant name", 7,
               None, None, None, None, None, None, "Custom",
               None, None, None, None, None, None])
    wb.save(path)
    return path


def _make_gal(path: Path) -> Path:
    path.write_text(
        "UserPrincipalName,Department,CompanyName,JobTitle,OnPremisesDN\n"
        "alice@contoso.com,Dept A,DIV-ONE,Data Scientist,"
        '"CN=Alice,OU=Users,OU=Central,OU=Regions,OU=MOE,DC=corp,DC=internal"\n'
        "bob.galonly@contoso.com,Dept B,,,\n",
        encoding="utf-8",
    )
    return path


def _make_exclusions(path: Path) -> Path:
    path.write_text(json.dumps({
        "_Description": "test exclusions",
        "ExcludedSITNames": ["Excluded SIT", "Tenant Excluded SIT"],
    }), encoding="utf-8")
    return path


def _make_sit_names(path: Path) -> Path:
    """Tenant GUID->name map in the CurrentTenantSITs.json shape."""
    path.write_text(json.dumps({
        "_Description": "test tenant SIT map",
        "_Count": 4,
        SIT_GUID_TENANT: "Tenant Map SIT",
        SIT_GUID_RAWNAME: "Tenant Shadow Name",  # raw-payload name must win
        SIT_GUID_BRIDGE: "Bridge Target",        # matches a workbook slug row
        SIT_GUID_TENANT_EXCL: "Tenant Excluded SIT",
        "not-a-guid": "ignored entry",
    }), encoding="utf-8")
    return path


def _make_org_mapping_config(path: Path) -> Path:
    """QFES-style override: Division from CompanyName, Department fallback.

    Deliberately partial — every other field must keep its built-in default."""
    path.write_text(json.dumps({
        "_Description": "test org mapping",
        "Division": {"Source": "CompanyName", "Fallback": "Department"},
    }), encoding="utf-8")
    return path


@pytest.fixture(scope="module")
def converted(tmp_path_factory) -> dict:
    root = tmp_path_factory.mktemp("star") / "Export-20260504-120000"
    export = _make_export(root)
    workbook = _make_workbook(root.parent / "SIT-Risk-Analysis-test.xlsx")
    gal = _make_gal(root.parent / "GAL_Clean.csv")
    exclusions = _make_exclusions(root.parent / "exclusions.json")
    org_mapping = _make_org_mapping_config(root.parent / "org-mapping.json")
    sit_names = _make_sit_names(root.parent / "CurrentTenantSITs-test.json")

    manifest = convert(
        export,
        risk_workbook=workbook,
        department_csv=gal,
        sit_exclusions=exclusions,
        org_mapping=org_mapping,  # explicit: hermetic against any repo-local config
        sit_names=sit_names,
        batch_size=2,  # force multiple sink flushes
    )
    return {"export": export, "output": export / "PowerBI-AE-Parquet-v6", "manifest": manifest}


def _read(converted: dict, table: str):
    return pq.read_table(converted["output"] / f"{table}.parquet")


def test_every_ssot_table_written_with_exact_schema(converted) -> None:
    for table_name in schema.TABLES:
        path = converted["output"] / f"{table_name}.parquet"
        assert path.exists(), f"{table_name}.parquet missing"
        actual = pq.read_schema(path)
        expected = schema.pyarrow_schema(table_name)
        assert actual.names == expected.names, table_name
        for name in expected.names:
            assert actual.field(name).type == expected.field(name).type, f"{table_name}.{name}"


def test_dedup_keeps_first_occurrence(converted) -> None:
    table = _read(converted, "fact_activity")
    assert table.num_rows == 5  # A, C, D(F2 dict page), E, F — duplicate A dropped
    manifest = converted["manifest"]
    assert manifest["raw_records_scanned"] == 6
    assert manifest["duplicates_skipped"] == 1


def test_f2_single_dict_records_page_kept(converted) -> None:
    index = _read(converted, "activity_record_index").to_pylist()
    by_rid = {row["record_identity"]: row for row in index}
    assert RID_D in by_rid
    assert by_rid[RID_D]["activity_id"] == stable_int_id("activity", RID_D)


def test_dim_date_continuous(converted) -> None:
    rows = _read(converted, "dim_date").to_pylist()
    assert [row["date_key"] for row in rows] == [20260501, 20260502, 20260503, 20260504]
    assert rows[0]["month_short"] == "May"
    assert rows[0]["week_of_year"] == 18


def test_extra_json_and_drift_report(converted) -> None:
    details = {row["activity_id"]: row for row in _read(converted, "fact_activity_detail").to_pylist()}
    row_a = details[stable_int_id("activity", RID_A)]
    extra = json.loads(row_a["extra_json"])
    assert extra["FutureUnknownField"] == "drift-me"
    drift_path = converted["output"] / "SchemaDrift.json"
    assert drift_path.exists()
    drift = json.loads(drift_path.read_text(encoding="utf-8"))
    assert "FutureUnknownField" in drift["unknown_fields_by_table"]["fact_activity_detail"]
    assert converted["manifest"]["schema_drift"]["total_unknown_fields"] >= 1


def test_sit_exclusion_filters_fact_and_aggs_only(converted) -> None:
    expected_keys = {
        SIT_GUID_OK, SIT_GUID_RAWNAME, SIT_GUID_TENANT, SIT_GUID_UNKNOWN,
        "bridge-target-slug",  # bridged onto the workbook slug row's key
    }
    sit_rows = _read(converted, "fact_activity_sit").to_pylist()
    assert {row["sit_key"] for row in sit_rows} == expected_keys
    manifest = converted["manifest"]
    assert manifest["sit_exclusions"]["sit_rows_before_exclusions"] == 7
    # workbook-named (record C) + tenant-named (record F) exclusions
    assert manifest["sit_exclusions"]["excluded_sit_match_rows"] == 2
    # the activity itself stays, with risk computed over ALL sits
    facts = {row["activity_id"]: row for row in _read(converted, "fact_activity").to_pylist()}
    row_c = facts[stable_int_id("activity", RID_C)]
    assert row_c["has_sit"] is True
    assert row_c["activity_risk_score"] == 5 * 4
    # aggregates exclude the excluded SIT too
    agg = _read(converted, "agg_department_sit_day").to_pylist()
    assert {row["sit_key"] for row in agg} == expected_keys


def test_fact_activity_sit_carries_v6_columns(converted) -> None:
    rows = _read(converted, "fact_activity_sit").to_pylist()
    row = rows[0]
    assert row["classifier_type"] == "Content"
    assert row["policy_rule_id"] is not None
    assert row["bucket_high"] == 2
    assert row["high_confidence_count"] == 2
    assert row["risk_weighted_count"] == 8 * 3


def test_email_detail_and_recipients(converted) -> None:
    details = _read(converted, "fact_email_detail").to_pylist()
    assert len(details) == 1  # record E's empty EmailInfo {} must not emit a row
    detail = details[0]
    assert detail["subject"] == "Q3 numbers"
    assert detail["message_id"] == "<msg-1@contoso.com>"
    assert detail["attachment_count"] == 2
    recipients = _read(converted, "fact_email_recipient").to_pylist()
    assert len(recipients) == 2
    # derived target domain: external receiver domain wins
    facts = {row["activity_id"]: row for row in _read(converted, "fact_activity").to_pylist()}
    row_a = facts[stable_int_id("activity", RID_A)]
    assert row_a["target_domain_id"] == stable_int_id("domain", "evil.com")


def test_copilot_interaction_row(converted) -> None:
    rows = _read(converted, "fact_copilot_interaction").to_pylist()
    assert len(rows) == 1
    row = rows[0]
    assert row["activity_id"] == stable_int_id("activity", RID_D)
    assert row["has_web_search_query"] is True
    assert row["are_files_referenced"] is False
    assert json.loads(row["copilot_event_data_json"])["AppHost"] == "Outlook"
    apps = _read(converted, "dim_app_identity").to_pylist()
    assert len(apps) == 1
    assert apps[0]["app_identity_group"] == "Copilot.M365Copilot"


def test_source_page_provenance(converted) -> None:
    pages = {row["page_id"]: row for row in _read(converted, "dim_source_page").to_pylist()}
    assert len(pages) == 3
    index = {row["record_identity"]: row for row in _read(converted, "activity_record_index").to_pylist()}
    page_row = pages[index[RID_D]["page_id"]]
    assert page_row["source_file"] == "Data/ActivityExplorer/2026-05-04/Page-001.json"
    assert page_row["watermark"] == "wm-day4"
    assert page_row["record_count"] == 1


def test_dim_user_unions_gal_with_has_activity(converted) -> None:
    users = {row["user_upn"]: row for row in _read(converted, "dim_user").to_pylist()}
    assert users["ALICE@CONTOSO.COM"]["has_activity"] is True
    assert users["BOB.GALONLY@CONTOSO.COM"]["has_activity"] is False
    departments = {row["department"] for row in _read(converted, "dim_department").to_pylist()}
    assert {"Dept A", "Dept B"} <= departments


def test_dim_user_org_enrichment_lands_in_parquet(converted) -> None:
    users = {row["user_upn"]: row for row in _read(converted, "dim_user").to_pylist()}
    alice = users["ALICE@CONTOSO.COM"]
    assert alice["division"] == "DIV-ONE"  # CompanyName wins over Department
    assert alice["region"] == "Central"  # OU directly under Regions
    assert alice["job_title"] == "Data Scientist"
    assert alice["is_leaver"] is False
    assert alice["is_generic_account"] is False
    bob = users["BOB.GALONLY@CONTOSO.COM"]
    assert bob["division"] == "Dept B"  # no CompanyName: Department fallback
    assert bob["region"] == "Unknown"  # no DN
    assert bob["job_title"] is None


def test_dim_sit_workbook_rows_with_observed_flag(converted) -> None:
    sits = {row["sit_name"]: row for row in _read(converted, "dim_sit").to_pylist()}
    assert sits["Test SIT One"]["observed"] is True
    assert sits["Test SIT One"]["risk_score"] == 8
    assert sits["Test SIT One"]["label_code"] == "LC1"
    assert sits["Excluded SIT"]["observed"] is True  # observed even though excluded
    assert sits["Custom Only"]["observed"] is False
    assert sits["Custom Only"]["sit_slug"] == "custom-only-slug"


def test_sit_name_resolution_chain(converted) -> None:
    """Display names: raw payload > tenant map > GUID fallback, with per-row
    provenance in source_sheet."""
    sits = {row["sit_key"]: row for row in _read(converted, "dim_sit").to_pylist()}
    raw_row = sits[SIT_GUID_RAWNAME]
    assert raw_row["sit_name"] == "Raw Payload SIT"  # beats "Tenant Shadow Name"
    assert raw_row["source_sheet"] == "Generated (name from AE raw payload)"
    assert raw_row["sit_id"] == SIT_GUID_RAWNAME
    assert raw_row["observed"] is True
    tenant_row = sits[SIT_GUID_TENANT]
    assert tenant_row["sit_name"] == "Tenant Map SIT"
    assert tenant_row["source_sheet"] == "Generated (name from tenant SIT map)"
    unknown_row = sits[SIT_GUID_UNKNOWN]
    assert unknown_row["sit_name"] == SIT_GUID_UNKNOWN  # GUID fallback
    assert unknown_row["source_sheet"] == "Generated from Activity Explorer export"
    # the tenant-named excluded SIT still gets a named, observed dim row
    assert sits[SIT_GUID_TENANT_EXCL]["sit_name"] == "Tenant Excluded SIT"
    assert sits[SIT_GUID_TENANT_EXCL]["observed"] is True


def test_sit_name_bridge_adopts_workbook_row(converted) -> None:
    """A tenant-map name matching a workbook slug row bridges the detection
    onto that row: its key, metadata and risk apply; no duplicate-name dim."""
    sit_rows = _read(converted, "fact_activity_sit").to_pylist()
    bridged = [row for row in sit_rows if row["sit_key"] == "bridge-target-slug"]
    assert len(bridged) == 1
    assert bridged[0]["risk_score"] == 7          # workbook metadata inherited
    assert bridged[0]["risk_weighted_count"] == 7 * 5
    dim = _read(converted, "dim_sit").to_pylist()
    assert SIT_GUID_BRIDGE not in {row["sit_key"] for row in dim}
    bridge_rows = [row for row in dim if row["sit_name"] == "Bridge Target"]
    assert len(bridge_rows) == 1                  # no duplicate-name dim row
    assert bridge_rows[0]["observed"] is True


def test_tenant_named_sit_matches_exclusion_list(converted) -> None:
    """Exclusion keys on the RESOLVED name: a SIT named only by the tenant
    map is excluded when that name is on the exclusion list."""
    sit_rows = _read(converted, "fact_activity_sit").to_pylist()
    assert SIT_GUID_TENANT_EXCL not in {row["sit_key"] for row in sit_rows}
    agg = _read(converted, "agg_department_sit_day").to_pylist()
    assert SIT_GUID_TENANT_EXCL not in {row["sit_key"] for row in agg}


def test_sit_name_resolution_manifest_provenance(converted) -> None:
    res = converted["manifest"]["sit_name_resolution"]
    assert res["tenant_sit_map"].endswith("CurrentTenantSITs-test.json")
    assert res["tenant_sit_map_entries"] == 4    # "_"-meta and non-GUID keys ignored
    assert res["observed_sits"] == 7
    assert res["resolved_by"] == {"workbook": 2, "raw_payload": 1, "tenant_map": 3}
    assert res["unresolved_guids"] == 1
    assert res["bridged_to_workbook"] == 1
    on_disk = json.loads(
        (converted["output"] / "manifest.json").read_text(encoding="utf-8"))
    assert on_disk["sit_name_resolution"] == res


def test_sit_name_map_loader_validation(tmp_path) -> None:
    from parquet_builder.star.enrich import load_sit_name_map, resolve_sit_names_path

    with pytest.raises(EnrichmentError):
        resolve_sit_names_path(tmp_path, tmp_path / "missing.json")
    assert resolve_sit_names_path(tmp_path, None, tmp_path / "absent-default.json") is None

    malformed = tmp_path / "malformed.json"
    malformed.write_text("{not json", encoding="utf-8")
    with pytest.raises(EnrichmentError):
        load_sit_name_map(malformed)
    not_object = tmp_path / "list.json"
    not_object.write_text("[1, 2]", encoding="utf-8")
    with pytest.raises(EnrichmentError):
        load_sit_name_map(not_object)


def test_archive_raw_written_by_default(converted) -> None:
    rows = {row["record_identity"]: row for row in _read(converted, "archive_raw").to_pylist()}
    assert len(rows) == 5
    sit_json = rows[RID_A]["sensitive_info_type_data"]
    assert SIT_GUID_OK in sit_json
    assert rows[RID_A]["original_activity_id"] == "DlpRuleMatch"


def test_schema_json_and_manifest_written(converted) -> None:
    schema_payload = json.loads((converted["output"] / "schema.json").read_text(encoding="utf-8"))
    assert schema_payload["version"] == 6
    manifest = json.loads((converted["output"] / "manifest.json").read_text(encoding="utf-8"))
    assert manifest["schema_version"] == 6
    assert manifest["profile"] == "powerbi_star"
    assert manifest["enrichment"]["sit_reference_rows"] == 4
    assert manifest["row_counts"]["fact_activity"] == 5


def test_unenriched_hard_fail_and_override(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(convert_module, "_DEFAULT_ORG_MAPPING", tmp_path / "absent.json")
    monkeypatch.setattr(convert_module, "_DEFAULT_SIT_NAMES", tmp_path / "absent-sits.json")
    export = _make_export(tmp_path / "Export-20260504-130000")
    with pytest.raises(EnrichmentError):
        convert(export, sit_exclusions=None)
    # with the override it must run and stamp enrichment: null
    manifest = convert(
        export,
        output_dir=tmp_path / "out-unenriched",
        allow_unenriched=True,
        sit_exclusions=None,
    )
    assert manifest["enrichment"] is None
    assert manifest["row_counts"]["fact_activity"] == 5
    # no workbook AND no tenant map: only the raw payload can name SITs
    res = manifest["sit_name_resolution"]
    assert res["tenant_sit_map"] is None
    assert res["resolved_by"] == {"workbook": 0, "raw_payload": 1, "tenant_map": 0}
    assert res["unresolved_guids"] == 6


# --- org-mapping resolution & provenance (engine unit tests live in
# --- test_star_org_mapping.py) -------------------------------------------------

def test_org_mapping_recorded_in_manifest(converted) -> None:
    org = converted["manifest"]["org_mapping"]
    assert org["source"].endswith("org-mapping.json")
    # config-supplied rule recorded as resolved
    assert org["fields"]["Division"] == {"Source": "CompanyName", "Fallback": "Department"}
    # omitted fields resolved to (and recorded as) the built-in defaults
    assert org["fields"]["Region"] == {
        "Source": "OnPremisesDN", "Parse": "ou_under", "Arg": "Regions"}
    assert org["fields"]["IsGenericAccount"] == {
        "Source": "OnPremisesDN", "Parse": "ou_name_in",
        "Arg": ["Generic Accounts", "SharedUsers"]}
    # the on-disk manifest carries the same provenance
    on_disk = json.loads(
        (converted["output"] / "manifest.json").read_text(encoding="utf-8"))
    assert on_disk["org_mapping"] == org


def test_no_archive_raw_flag(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(convert_module, "_DEFAULT_ORG_MAPPING", tmp_path / "absent.json")
    monkeypatch.setattr(convert_module, "_DEFAULT_SIT_NAMES", tmp_path / "absent-sits.json")
    export = _make_export(tmp_path / "Export-20260504-140000")
    workbook = _make_workbook(tmp_path / "SIT-Risk-Analysis-test.xlsx")
    gal = _make_gal(tmp_path / "GAL_Clean.csv")
    out = tmp_path / "out-noarchive"
    manifest = convert(
        export, output_dir=out, risk_workbook=workbook, department_csv=gal,
        archive_raw=False, sit_exclusions=None,
    )
    assert not (out / "archive_raw.parquet").exists()
    assert "archive_raw" not in manifest["row_counts"]
    assert manifest["options"]["archive_raw"] is False
