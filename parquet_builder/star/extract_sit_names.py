"""CLI: distil SIT-name artifacts into a SIT GUID->name map.

Why this exists: Activity Explorer detections mostly carry classification
sub-entity GUIDs that neither the flat ``Get-DlpSensitiveInformationType``
list nor the risk workbook know, so they fall back to GUID labels in
dim_sit. The tenant's DLP rule packages name every classification rule GUID
(the export tool saves them via ``Export-SitReferenceSnapshot`` and merges
the names into ``<ExportDir>/CurrentTenantSITs.json`` automatically), and
portal export artifacts carry the same pairs. This tool reads any of those
artifacts and emits one merged flat map in the CurrentTenantSITs.json shape
that ``py -m parquet_builder.star.convert --sit-names`` already accepts.

Supported input formats (sniffed by content/extension, not filename):

- flat map      ``{"<guid>": "<name>"}`` — CurrentTenantSITs.json or a
                previous output of this tool (``_``-prefixed keys ignored)
- rule-pack XML ``ClassificationRuleCollection`` / ``RulePackage`` XML as
                exported by ``Get-DlpSensitiveInformationTypeRulePackage``
                (saved under ``Data/Reference/RulePackages/*.xml``); Entity
                AND Affinity GUIDs are named via the LocalizedStrings
                resources, preferring the default-language name
- tag groups    ``[{"TagRecords": [{"Id", "Name"}], ...}]`` — portal
                type-catalog export (includes the TrainableClassifier
                group, so TC GUIDs resolve too)
- records       ``{"Records": [{"Id", "Name"}]}`` — portal aggregate-probe
                export
- folder index  ``{"sits": {"<guid>": {"name": ...}}}`` — portal export
                folder index
- sit catalog   parquet with ``sit_id``/``sit_name`` columns
- csv           any CSV with an Id column and a DisplayName/Name column
                (e.g. a trainable-classifier export)

Inputs are processed in CLI order and the FIRST source naming a GUID wins,
so list the tenant's own snapshot before cross-tenant artifacts.

``--strip-prefix`` (repeatable) removes a deployment-pack display prefix
(e.g. ``"QGISCF - "``) from incoming names: custom-pack SITs keep the same
GUIDs across tenants but each tenant prefixes the display names, while the
curated risk workbook lists them under the base names. With the prefix
stripped, the converter's name bridge attaches the workbook row (risk
metadata + exclusion behavior) instead of creating a bare named row.

Usage (the merged-map generation command):
    py -m parquet_builder.star.extract_sit_names
        --input <ExportDir>/CurrentTenantSITs.json
        --input <ExportDir>/Data/Reference/RulePackages/Microsoft_Rule_Package.xml
        --strip-prefix "QGISCF - "
        [--output ConfigFiles/SITNames.local.json]

The default output ConfigFiles/SITNames.local.json is gitignored
(``*.local.json``) — the emitted map is tenant data and must not be tracked.
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from .errors import EnrichmentError
from .sit_reference import GUID_RE

_CONFIG_DIR = Path(__file__).resolve().parents[2] / "ConfigFiles"
_DEFAULT_OUTPUT = _CONFIG_DIR / "SITNames.local.json"

# CSV column candidates, in preference order.
_CSV_ID_COLUMNS = ("Id", "ID", "id", "Guid", "GUID", "sit_id", "SitId")
_CSV_NAME_COLUMNS = ("DisplayName", "display_name", "Name", "name", "sit_name", "SitName")


def _clean_name(name: Any, strip_prefixes: tuple[str, ...]) -> str | None:
    """Normalise one display name; None when unusable (empty / GUID label)."""
    if name is None:
        return None
    text = " ".join(str(name).split())
    if not text or GUID_RE.match(text):
        return None
    for prefix in strip_prefixes:
        if text.startswith(prefix):
            stripped = text[len(prefix):].strip()
            if stripped:
                text = stripped
            break
    return text


def _pairs_from_flat_map(payload: dict[str, Any]) -> list[tuple[str, Any]]:
    return [
        (key, value) for key, value in payload.items()
        if not str(key).startswith("_")
    ]


def _pairs_from_id_name_records(records: list[Any]) -> list[tuple[str, Any]]:
    pairs: list[tuple[str, Any]] = []
    for record in records:
        if isinstance(record, dict):
            pairs.append((str(record.get("Id") or ""), record.get("Name")))
    return pairs


def _pairs_from_json(payload: Any, path: Path) -> list[tuple[str, Any]]:
    """Sniff the JSON artifact shape and extract raw (guid, name) pairs."""
    if isinstance(payload, dict) and isinstance(payload.get("sits"), dict):
        # Portal export folder index (sit_folder_index*.json)
        return [
            (str(guid), entry.get("name"))
            for guid, entry in payload["sits"].items()
            if isinstance(entry, dict)
        ]
    if isinstance(payload, dict) and isinstance(payload.get("Records"), list):
        # Portal aggregate-probe export ({"Records": [{Id, Name}]})
        return _pairs_from_id_name_records(payload["Records"])
    if isinstance(payload, list):
        # Portal type-catalog export: groups with TagRecords.
        pairs: list[tuple[str, Any]] = []
        matched = False
        for group in payload:
            if isinstance(group, dict) and isinstance(group.get("TagRecords"), list):
                matched = True
                pairs.extend(_pairs_from_id_name_records(group["TagRecords"]))
        if matched:
            return pairs
        raise EnrichmentError(f"Unrecognised JSON list format (no TagRecords groups): {path}")
    if isinstance(payload, dict):
        # Flat CurrentTenantSITs.json-shaped map (also this tool's own output).
        return _pairs_from_flat_map(payload)
    raise EnrichmentError(f"Unrecognised JSON artifact format: {path}")


def _xml_local_name(tag: Any) -> str:
    """Local element name without the XML namespace prefix."""
    return str(tag).rsplit("}", 1)[-1]


def _pairs_from_rulepack_xml(path: Path) -> list[tuple[str, Any]]:
    """Extract (guid, name) pairs from ClassificationRuleCollection XML.

    The rule-package schema requires a ``LocalizedStrings/Resource`` element
    (keyed by ``idRef`` GUID) for every Entity and Affinity rule, so the
    Resource elements name ALL classification rule GUIDs in the pack —
    including sub-entity GUIDs the flat SIT list does not surface. The
    default-language ``Name`` (``default="true"``) is preferred, falling
    back to the first non-empty ``Name``. Namespace-agnostic.
    """
    try:
        root = ET.parse(path).getroot()
    except (OSError, ET.ParseError) as exc:
        raise EnrichmentError(f"Cannot parse rule-pack XML {path}: {exc}") from exc

    pairs: list[tuple[str, Any]] = []
    for element in root.iter():
        if _xml_local_name(element.tag) != "Resource":
            continue
        guid = str(element.get("idRef") or "")
        default_name: str | None = None
        first_name: str | None = None
        for child in element:
            if _xml_local_name(child.tag) != "Name":
                continue
            text = (child.text or "").strip()
            if not text:
                continue
            if first_name is None:
                first_name = text
            if (child.get("default") or "").strip().lower() == "true":
                default_name = text
                break
        pairs.append((guid, default_name or first_name))
    return pairs


def _pairs_from_parquet(path: Path) -> list[tuple[str, Any]]:
    try:
        import pyarrow.parquet as pq
    except ModuleNotFoundError as exc:  # pragma: no cover - environment issue
        raise EnrichmentError(
            f"Missing Python dependency '{exc.name}'. "
            "Install runtime dependencies with `pip install -r requirements.txt`."
        ) from exc
    table = pq.read_table(path)
    if "sit_id" not in table.column_names or "sit_name" not in table.column_names:
        raise EnrichmentError(
            f"Parquet artifact has no sit_id/sit_name columns: {path} "
            f"(found: {', '.join(table.column_names)})"
        )
    ids = table.column("sit_id").to_pylist()
    names = table.column("sit_name").to_pylist()
    return [(str(guid or ""), name) for guid, name in zip(ids, names)]


def _pairs_from_csv(path: Path) -> list[tuple[str, Any]]:
    with open(path, "r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        fields = reader.fieldnames or []
        id_column = next((c for c in _CSV_ID_COLUMNS if c in fields), None)
        name_column = next((c for c in _CSV_NAME_COLUMNS if c in fields), None)
        if id_column is None or name_column is None:
            raise EnrichmentError(
                f"CSV artifact has no recognisable Id/Name columns: {path} "
                f"(found: {', '.join(fields)})"
            )
        return [
            (str(row.get(id_column) or ""), row.get(name_column))
            for row in reader
        ]


def extract_pairs(path: Path) -> list[tuple[str, Any]]:
    """Extract raw (guid, name) pairs from one artifact (format-sniffed)."""
    suffix = path.suffix.lower()
    if suffix == ".parquet":
        return _pairs_from_parquet(path)
    if suffix == ".csv":
        return _pairs_from_csv(path)
    if suffix == ".xml":
        return _pairs_from_rulepack_xml(path)
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        # Not JSON: rule-pack XML may arrive without a .xml extension. Sniff
        # for a leading '<' (errors="ignore" tolerates UTF-16 raw bytes) and
        # let ElementTree do the real, encoding-aware parse.
        try:
            head = path.read_text(encoding="utf-8", errors="ignore")[:256]
        except OSError:
            head = ""
        if head.lstrip("\ufeff\x00 \t\r\n").startswith("<"):
            return _pairs_from_rulepack_xml(path)
        raise EnrichmentError(f"Cannot read artifact {path}: {exc}") from exc
    return _pairs_from_json(payload, path)


def merge_name_maps(
    inputs: list[Path], strip_prefixes: tuple[str, ...] = (),
) -> tuple[dict[str, str], list[dict[str, Any]]]:
    """Merge artifacts into one GUID->name map (first source wins per GUID).

    Returns (merged map, per-source summaries).
    """
    merged: dict[str, str] = {}
    summaries: list[dict[str, Any]] = []
    for path in inputs:
        if not path.exists():
            raise EnrichmentError(f"--input does not exist: {path}")
        usable = 0
        added = 0
        for raw_guid, raw_name in extract_pairs(path):
            guid = raw_guid.strip().lower()
            if not GUID_RE.match(guid):
                continue
            name = _clean_name(raw_name, strip_prefixes)
            if name is None:
                continue
            usable += 1
            if guid not in merged:
                merged[guid] = name
                added += 1
        summaries.append({"source": str(path), "usable_pairs": usable, "added": added})
    return merged, summaries


def write_name_map(
    output: Path, merged: dict[str, str], summaries: list[dict[str, Any]],
    strip_prefixes: tuple[str, ...],
) -> None:
    payload: dict[str, Any] = {
        "_Description": (
            "Merged SIT/classifier GUID->name map for "
            "`py -m parquet_builder.star.convert --sit-names`. Generated by "
            "extract_sit_names; do not edit or track (tenant data)."
        ),
        "_GeneratedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "_Sources": [entry["source"] for entry in summaries],
        "_StripPrefixes": list(strip_prefixes),
        "_Count": len(merged),
    }
    payload.update({guid: merged[guid] for guid in sorted(merged)})
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Distil SIT-name artifacts (tenant snapshots, rule-pack XML, "
            "portal export artifacts) into one merged SIT GUID->name map "
            "(CurrentTenantSITs.json shape) for convert --sit-names."
        )
    )
    parser.add_argument(
        "--input", action="append", required=True, metavar="PATH",
        help="Artifact to read (repeatable; first source naming a GUID wins). "
             "Formats: flat map / type-catalog / Records / folder-index JSON, "
             "rule-pack XML, sit_catalog parquet, Id+Name CSV.")
    parser.add_argument(
        "--output", default=str(_DEFAULT_OUTPUT),
        help=f"Output JSON path (default: {_DEFAULT_OUTPUT}; gitignored).")
    parser.add_argument(
        "--strip-prefix", action="append", default=[], metavar="PREFIX",
        help="Display-name prefix to strip (repeatable), e.g. \"QGISCF - \". "
             "Aligns deployment-pack names with the risk workbook's base "
             "names so the converter's workbook bridge applies.")
    args = parser.parse_args(argv)

    inputs = [Path(item) for item in args.input]
    strip_prefixes = tuple(args.strip_prefix)
    output = Path(args.output)

    try:
        merged, summaries = merge_name_maps(inputs, strip_prefixes)
        write_name_map(output, merged, summaries, strip_prefixes)
    except EnrichmentError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    for entry in summaries:
        print(f"  {entry['source']}: {entry['usable_pairs']} usable pairs, "
              f"{entry['added']} new GUIDs")
    print(f"Wrote {len(merged)} GUID->name entries to {output}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
