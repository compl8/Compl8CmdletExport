"""Policy and RBAC config processing."""

from __future__ import annotations

from pathlib import Path

from .helpers import _now_iso, _safe_str
from .loaders import load_json_config


def process_dlp_policies(input_dir: Path) -> tuple[list[dict], list[dict]]:
    """Process DLP-Config.json -> (policies, rules)."""
    data = load_json_config(input_dir / "DLP-Config.json")
    if not data:
        return [], []

    ingested_at = _now_iso()
    policies = []
    rules = []

    for p in data.get("Policies", []):
        row = {k: _safe_str(v) for k, v in p.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        policies.append(row)

    for r in data.get("Rules", []):
        row = {k: _safe_str(v) for k, v in r.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        rules.append(row)

    print(f"  DLP: {len(policies)} policies, {len(rules)} rules")
    return policies, rules


def process_sensitivity_labels(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "SensitivityLabels-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    labels = []
    for lbl in data.get("Labels", []):
        row = {k: _safe_str(v) for k, v in lbl.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        labels.append(row)

    print(f"  Sensitivity labels: {len(labels)} records")
    return labels


def process_retention_labels(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "RetentionLabels-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    labels = []
    for lbl in data.get("Labels", []):
        row = {k: _safe_str(v) for k, v in lbl.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        labels.append(row)

    print(f"  Retention labels: {len(labels)} records")
    return labels


def process_rbac(input_dir: Path) -> list[dict]:
    data = load_json_config(input_dir / "RBAC-Config.json")
    if not data:
        return []

    ingested_at = _now_iso()
    groups = []
    for rg in data.get("RoleGroups", []):
        row = {k: _safe_str(v) for k, v in rg.items()}
        row["_source_tool"] = "cmdletexport"
        row["_ingested_at"] = ingested_at
        groups.append(row)

    print(f"  RBAC role groups: {len(groups)} records")
    return groups
