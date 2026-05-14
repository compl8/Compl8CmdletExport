"""CSV / JSON / Compl8 cache writers and the console summary table."""

from __future__ import annotations

import csv
import json
import time
from pathlib import Path

from .constants import TYPE_TC, log


def write_csv(results: list[dict], path: Path) -> None:
    keys = list(dict.fromkeys(k for r in results for k in r))
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys, extrasaction="ignore")
        w.writeheader()
        for r in results:
            row = {}
            for k, v in r.items():
                row[k] = json.dumps(v) if isinstance(v, (dict, list)) else v
            w.writerow(row)
    log.info("Saved %d entries (CSV) -> %s", len(results), path)


def write_json(results: list[dict], path: Path) -> None:
    path.write_text(json.dumps(results, indent=2), "utf-8")
    log.info("Saved %d entries (JSON) -> %s", len(results), path)


def write_compl8_format(results: list[dict], path: Path, tenant_id: str | None) -> int:
    """Write the Compl8 trainable-classifier cache file.

    Schema is intentionally small — the CE orchestrator's discovery code only
    needs Name to schedule aggregate/detail tasks. Extra metadata is included
    so downstream tooling can join against the parquet content_files table
    by classifier name without re-scraping.
    """
    tcs = [r for r in results if r.get("_Type") == TYPE_TC]
    classifiers = []
    for r in tcs:
        name = r.get("Name") or r.get("DisplayName")
        if not name:
            continue
        classifiers.append({
            "Id": r.get("Id") or r.get("ModelId"),
            "Name": name,
            "DisplayName": r.get("DisplayName") or name,
            "Type": r.get("type") or r.get("ModelType"),
            "SubType": r.get("subType"),
            "BusinessFunction": r.get("businessFunction"),
            "Applications": r.get("applications"),
            "Languages": (r.get("Languages") or "").split(", ") if r.get("Languages") else [],
            "ModelStatus": r.get("ModelStatus") or r.get("JobProcessedStatus"),
            "IsPublished": bool(r.get("IsPublished")),
            "IsDeprecated": bool(r.get("isDeprecated")),
        })
    payload = {
        "SchemaVersion": 1,
        "DiscoveredAt": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        "TenantId": tenant_id,
        "Source": "purview-portal",
        "ClassifierCount": len(classifiers),
        "Classifiers": sorted(classifiers, key=lambda c: str(c["Name"]).lower()),
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2), "utf-8")
    log.info("Saved %d trainable classifiers (Compl8 format) -> %s", len(classifiers), path)
    return len(classifiers)


def print_table(results: list[dict]) -> None:
    """Print a one-line-per-classifier summary table to console."""
    tcs = [r for r in results if r.get("_Type") == TYPE_TC]
    if not tcs:
        return

    print(f"\n{'Name':<45} {'Type':<12} {'Languages':>5}  {'Status':<12}  ID")
    print("-" * 120)
    for r in sorted(tcs, key=lambda x: str(x.get("Name", "")).lower()):
        name = str(r.get("Name") or r.get("DisplayName") or "?")[:43]
        ctype = str(r.get("type") or r.get("ModelType") or "")[:10]
        langs = r.get("Languages", "")
        lang_count = str(len(langs.split(", "))) if langs else "1"
        status = str(r.get("ModelStatus") or r.get("JobProcessedStatus") or "")[:10]
        rid = r.get("Id") or r.get("ModelId") or "?"
        print(f"{name:<45} {ctype:<12} {lang_count:>5}  {status:<12}  {rid}")
