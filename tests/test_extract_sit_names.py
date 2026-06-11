"""Tests for parquet_builder.star.extract_sit_names (synthetic fixtures only)."""

from __future__ import annotations

import json
from pathlib import Path

import pyarrow as pa
import pyarrow.parquet as pq
import pytest

from parquet_builder.star.errors import EnrichmentError
from parquet_builder.star.extract_sit_names import (
    extract_pairs,
    main,
    merge_name_maps,
)
from parquet_builder.star.sit_reference import load_sit_name_map

GUID_A = "11111111-1111-1111-1111-111111111111"
GUID_B = "22222222-2222-2222-2222-222222222222"
GUID_C = "33333333-3333-3333-3333-333333333333"
GUID_D = "44444444-4444-4444-4444-444444444444"


def _write_json(path: Path, payload) -> Path:
    path.write_text(json.dumps(payload), encoding="utf-8")
    return path


def _rulepack_xml() -> str:
    """Synthetic ClassificationRuleCollection (RulePackage) XML.

    Mirrors the documented schema: Entity + Affinity rules, with the
    LocalizedStrings/Resource elements naming each rule GUID. Resource A
    lists a non-default localized name FIRST so the default="true" name
    must win; Resource B has a single name without a default attribute.
    """
    return f"""<?xml version="1.0" encoding="utf-16"?>
<RulePackage xmlns="http://schemas.microsoft.com/office/2011/mce">
  <RulePack id="{GUID_D}">
    <Version build="0" major="1" minor="0" revision="0"/>
    <Publisher id="{GUID_C}"/>
    <Details defaultLangCode="en-us">
      <LocalizedDetails langcode="en-us">
        <PublisherName>Synthetic</PublisherName>
        <Name>Synthetic Rule Pack</Name>
        <Description>Test fixture - no real tenant data.</Description>
      </LocalizedDetails>
    </Details>
  </RulePack>
  <Rules>
    <Entity id="{GUID_A}" patternsProximity="300" recommendedConfidence="75">
      <Pattern confidenceLevel="75"><IdMatch idRef="Regex_a"/></Pattern>
    </Entity>
    <Affinity id="{GUID_B}" evidencesProximity="300" thresholdConfidenceLevel="75">
      <Evidence confidenceLevel="75"><Match idRef="Regex_a"/></Evidence>
    </Affinity>
    <Regex id="Regex_a">(\\d{{9}})</Regex>
    <LocalizedStrings>
      <Resource idRef="{GUID_A}">
        <Name default="false" langcode="de-de">Mitarbeiter-ID</Name>
        <Name default="true" langcode="en-us">Employee ID</Name>
      </Resource>
      <Resource idRef="{GUID_B}">
        <Name langcode="en-us">All Credentials Bundle</Name>
      </Resource>
      <Resource idRef="not-a-guid">
        <Name default="true" langcode="en-us">Ignored Non-Guid</Name>
      </Resource>
    </LocalizedStrings>
  </Rules>
</RulePackage>
"""


def _write_rulepack_xml(path: Path) -> Path:
    # UTF-16 with BOM, matching how the rule-package cmdlet emits the XML.
    path.write_bytes(_rulepack_xml().encode("utf-16"))
    return path


# ---------------------------------------------------------------- formats


def test_flat_map_format(tmp_path) -> None:
    path = _write_json(tmp_path / "flat.json", {
        "_Description": "metadata ignored",
        "_Count": 2,
        GUID_A: "Credit Card Number",
        GUID_B.upper(): "Tax File Number",   # GUID case-normalised
        "not-a-guid": "ignored",
    })
    merged, summaries = merge_name_maps([path])
    assert merged == {GUID_A: "Credit Card Number", GUID_B: "Tax File Number"}
    assert summaries[0]["usable_pairs"] == 2


def test_portal_gettypes_tag_groups_format(tmp_path) -> None:
    path = _write_json(tmp_path / "ip-gettypes.json", [
        {"Type": "SensitiveInformationType", "DisplayName": "Sensitive info types",
         "TagRecords": [{"Name": "Pack - Alpha", "Id": GUID_A},
                        {"Name": "Beta", "Id": GUID_B}]},
        {"Type": "Retention", "TagRecords": []},
        {"Type": "TrainableClassifier",
         "TagRecords": [{"Name": "Source code", "Id": GUID_C}]},
    ])
    merged, _ = merge_name_maps([path])
    # all groups contribute, including trainable classifiers
    assert merged == {GUID_A: "Pack - Alpha", GUID_B: "Beta", GUID_C: "Source code"}


def test_preindex_records_format(tmp_path) -> None:
    path = _write_json(tmp_path / "_preindex.json", {
        "AggregateType": "Sit",
        "Records": [{"Name": "Alpha", "Id": GUID_A, "Count": 3},
                    {"Name": None, "Id": GUID_B},          # unnamed -> skipped
                    {"Name": GUID_C, "Id": GUID_C}],       # GUID label -> skipped
    })
    merged, summaries = merge_name_maps([path])
    assert merged == {GUID_A: "Alpha"}
    assert summaries[0]["usable_pairs"] == 1


def test_sit_folder_index_format(tmp_path) -> None:
    path = _write_json(tmp_path / "sit_folder_index.json", {
        "total_sits": 2,
        "sits": {
            GUID_A: {"count": 5, "folders": [], "name": "Alpha"},
            GUID_B: {"count": 1, "folders": []},  # no name -> skipped
        },
    })
    merged, _ = merge_name_maps([path])
    assert merged == {GUID_A: "Alpha"}


def test_sit_catalog_parquet_format(tmp_path) -> None:
    path = tmp_path / "sit_catalog.parquet"
    pq.write_table(pa.table({
        "sit_id": [GUID_A, GUID_B, GUID_C],
        "sit_name": ["Alpha", None, GUID_C],  # null and GUID labels skipped
        "name_source": ["envelope", "none", "envelope"],
    }), path)
    merged, _ = merge_name_maps([path])
    assert merged == {GUID_A: "Alpha"}


def test_parquet_without_expected_columns_is_an_error(tmp_path) -> None:
    path = tmp_path / "other.parquet"
    pq.write_table(pa.table({"foo": [1]}), path)
    with pytest.raises(EnrichmentError, match="sit_id/sit_name"):
        extract_pairs(path)


def test_csv_format_prefers_displayname(tmp_path) -> None:
    path = tmp_path / "trainable_classifiers.csv"
    path.write_text(
        "Id,Name,DisplayName\n"
        f"{GUID_A},{GUID_A},Source code\n"
        f"{GUID_B},Threat,\n",   # empty DisplayName -> skipped
        encoding="utf-8",
    )
    merged, _ = merge_name_maps([path])
    assert merged == {GUID_A: "Source code"}


def test_csv_without_id_name_columns_is_an_error(tmp_path) -> None:
    path = tmp_path / "bad.csv"
    path.write_text("Foo,Bar\n1,2\n", encoding="utf-8")
    with pytest.raises(EnrichmentError, match="Id/Name columns"):
        extract_pairs(path)


def test_unrecognised_json_format_is_an_error(tmp_path) -> None:
    path = _write_json(tmp_path / "weird.json", [1, 2, 3])
    with pytest.raises(EnrichmentError, match="Unrecognised"):
        extract_pairs(path)


def test_rulepack_xml_format(tmp_path) -> None:
    """ClassificationRuleCollection XML: Entity AND Affinity GUIDs named via
    LocalizedStrings, default-language name preferred, non-GUID ids dropped."""
    path = _write_rulepack_xml(tmp_path / "Microsoft Rule Package.xml")
    merged, summaries = merge_name_maps([path])
    assert merged == {
        GUID_A: "Employee ID",            # default="true" beats the first (de-de) name
        GUID_B: "All Credentials Bundle",  # affinity rule, single name without default
    }
    assert summaries[0]["usable_pairs"] == 2


def test_rulepack_xml_sniffed_without_extension(tmp_path) -> None:
    """A rule-pack artifact without a .xml extension is content-sniffed."""
    path = _write_rulepack_xml(tmp_path / "rulepack.dat")
    merged, _ = merge_name_maps([path])
    assert merged[GUID_A] == "Employee ID"
    assert merged[GUID_B] == "All Credentials Bundle"


def test_malformed_rulepack_xml_is_an_error(tmp_path) -> None:
    path = tmp_path / "bad.xml"
    path.write_text("<RulePackage><Rules>", encoding="utf-8")
    with pytest.raises(EnrichmentError, match="rule-pack XML"):
        extract_pairs(path)


def test_flat_map_wins_over_rulepack_xml(tmp_path) -> None:
    """Merge precedence: the tenant's flat snapshot (listed first) beats the
    rule-pack name for the same GUID; the pack still fills the gaps."""
    flat = _write_json(tmp_path / "CurrentTenantSITs.json", {GUID_A: "Tenant Alpha"})
    pack = _write_rulepack_xml(tmp_path / "pack.xml")
    merged, summaries = merge_name_maps([flat, pack])
    assert merged == {GUID_A: "Tenant Alpha", GUID_B: "All Credentials Bundle"}
    assert summaries[1]["usable_pairs"] == 2
    assert summaries[1]["added"] == 1


# ------------------------------------------------------- merge semantics


def test_first_input_wins_per_guid(tmp_path) -> None:
    first = _write_json(tmp_path / "first.json", {GUID_A: "Tenant Name"})
    second = _write_json(tmp_path / "second.json",
                         {GUID_A: "Portal Name", GUID_B: "Beta"})
    merged, summaries = merge_name_maps([first, second])
    assert merged == {GUID_A: "Tenant Name", GUID_B: "Beta"}
    assert summaries[1]["usable_pairs"] == 2
    assert summaries[1]["added"] == 1


def test_strip_prefix_applied_and_whitespace_collapsed(tmp_path) -> None:
    path = _write_json(tmp_path / "flat.json", {
        GUID_A: "PackA - Alpha  record type",
        GUID_B: "PackB - Beta record type",
        GUID_C: "Unprefixed name",
        GUID_D: "PackA - ",  # strips to empty -> original name kept
    })
    merged, _ = merge_name_maps(
        [path], strip_prefixes=("PackA - ", "PackB - "))
    assert merged[GUID_A] == "Alpha record type"
    assert merged[GUID_B] == "Beta record type"
    assert merged[GUID_C] == "Unprefixed name"
    assert merged[GUID_D] == "PackA -"  # whitespace-collapsed original

def test_missing_input_is_an_error(tmp_path) -> None:
    with pytest.raises(EnrichmentError, match="does not exist"):
        merge_name_maps([tmp_path / "nope.json"])


# ------------------------------------------------------------------- CLI


def test_main_output_round_trips_through_sit_names_loader(tmp_path) -> None:
    """The emitted file must be consumable by convert --sit-names."""
    source = _write_json(tmp_path / "ip-gettypes.json", [
        {"Type": "SensitiveInformationType",
         "TagRecords": [{"Name": "QGISCF - Alpha", "Id": GUID_A},
                        {"Name": "Beta", "Id": GUID_B}]},
    ])
    output = tmp_path / "SITNames.local.json"
    rc = main(["--input", str(source), "--output", str(output),
               "--strip-prefix", "QGISCF - "])
    assert rc == 0
    payload = json.loads(output.read_text(encoding="utf-8"))
    assert payload["_Count"] == 2
    assert payload["_StripPrefixes"] == ["QGISCF - "]
    # round-trip through the exact loader the converter uses
    assert load_sit_name_map(output) == {GUID_A: "Alpha", GUID_B: "Beta"}


def test_main_reports_error_and_exit_code(tmp_path, capsys) -> None:
    rc = main(["--input", str(tmp_path / "missing.json"),
               "--output", str(tmp_path / "out.json")])
    assert rc == 1
    assert "does not exist" in capsys.readouterr().err
    assert not (tmp_path / "out.json").exists()
