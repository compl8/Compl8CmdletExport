"""CLI: distil portal-extract artifacts into a SIT GUID->name map.

Why this exists: Activity Explorer detections mostly carry classification
sub-entity GUIDs that neither ``Get-DlpSensitiveInformationType`` (the
CurrentTenantSITs.json snapshot) nor the risk workbook know, so they fall
back to GUID labels in dim_sit. The Purview portal names ALL of them, and
portal-API extract tools (Compl8Extractor_new, C8PortalPuller) persist the
GUID->name pairs in their outputs. This tool reads those artifacts and emits
one merged flat map in the CurrentTenantSITs.json shape that
``py -m parquet_builder.star.convert --sit-names`` already accepts.

Supported input formats (sniffed by content/extension, not filename):

- flat map      ``{"<guid>": "<name>"}`` — CurrentTenantSITs.json or a
                previous output of this tool (``_``-prefixed keys ignored)
- tag groups    ``[{"TagRecords": [{"Id", "Name"}], ...}]`` — portal
                getTypes dump (C8PortalPuller ``ip-gettypes.json``; includes
                the TrainableClassifier group, so TC GUIDs resolve too)
- records       ``{"Records": [{"Id", "Name"}]}`` — portal preindex
                aggregate probe (Compl8Extractor_new ``_preindex_v2_full.json``)
- folder index  ``{"sits": {"<guid>": {"name": ...}}}`` — Compl8Extractor_new
                ``sit_folder_index*.json``
- sit catalog   parquet with ``sit_id``/``sit_name`` columns —
                Compl8Extractor_new warehouse ``sit_catalog.parquet``
- csv           any CSV with an Id column and a DisplayName/Name column —
                GetTCs ``trainable_classifiers.csv``

Inputs are processed in CLI order and the FIRST source naming a GUID wins,
so list the tenant's own snapshot before cross-tenant portal artifacts.

``--strip-prefix`` (repeatable) removes a deployment-pack display prefix
(e.g. ``"QGISCF - "``) from incoming names: custom-pack SITs keep the same
GUIDs across tenants but each tenant prefixes the display names, while the
curated risk workbook lists them under the base names. With the prefix
stripped, the converter's name bridge attaches the workbook row (risk
metadata + exclusion behavior) instead of creating a bare named row.

Usage (the merged-map generation command):
    py -m parquet_builder.star.extract_sit_names
        --input ConfigFiles/CurrentTenantSITs.json
        --input <portal-extract>/ip-gettypes.json
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
        # Compl8Extractor_new sit_folder_index*.json
        return [
            (str(guid), entry.get("name"))
            for guid, entry in payload["sits"].items()
            if isinstance(entry, dict)
        ]
    if isinstance(payload, dict) and isinstance(payload.get("Records"), list):
        # Portal preindex aggregate probe (_preindex_v2_full.json)
        return _pairs_from_id_name_records(payload["Records"])
    if isinstance(payload, list):
        # Portal getTypes dump (ip-gettypes.json): groups with TagRecords.
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
    try:
        payload = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
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
            "Distil portal-extract artifacts into one merged SIT GUID->name "
            "map (CurrentTenantSITs.json shape) for convert --sit-names."
        )
    )
    parser.add_argument(
        "--input", action="append", required=True, metavar="PATH",
        help="Artifact to read (repeatable; first source naming a GUID wins). "
             "Formats: flat map / portal getTypes / preindex Records / "
             "sit_folder_index JSON, sit_catalog parquet, Id+Name CSV.")
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
