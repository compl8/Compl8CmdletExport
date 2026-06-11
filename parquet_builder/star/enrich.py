"""Enrichment input loaders: SIT risk workbook and GAL/department mapping.

F1 fix policy: enrichment inputs are resolved from explicit CLI paths first,
then the export root, then one directory level below it. If either input is
still missing the conversion FAILS unless --allow-unenriched was passed —
silent unenriched output (risk scores all zero, one department) poisoned the
v5 run and is no longer possible.

SIT reference loading (risk workbook, tenant GUID->name map, the
per-detection name-resolution chain) lives in sit_reference.py and is
re-exported here for backward compatibility.

Org-field sourcing for dim_user (division/region/job_title/is_leaver/
is_generic_account) is config-driven — see org_mapping.py (the engine) and
ConfigFiles/AEStarOrgMapping.example.json (the config contract).
"""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from .errors import EnrichmentError
from .org_mapping import (  # noqa: F401  (re-exported for backward compatibility)
    ORG_MAPPING_SOURCE_DEFAULTS,
    OrgFieldRule,
    OrgMapping,
    _column_value,
    _extract_org_flag,
    _extract_org_value,
    _normalized_row,
    _validate_org_columns,
    default_org_mapping,
    load_org_mapping,
    ou_contains,
    ou_name_in,
    ou_under,
    parse_dn_ous,
)
from .sit_reference import (  # noqa: F401  (re-exported for backward compatibility)
    GUID_RE,
    RiskLookup,
    _cell_bool,
    _cell_int,
    _cell_str,
    _norm_text,
    _read_worksheet_rows,
    _search_one,
    load_risk_workbook,
    load_sit_name_map,
    resolve_detected_sit,
    resolve_risk_workbook,
    resolve_sit_names_path,
    risk_band,
    sit_key_for,
    sit_key_for_detected_id,
)

_DEPARTMENT_FILE_CANDIDATES = (
    "user_department_mapping.csv",
    "user_department_mapping.xlsx",
    "User-Department-Mapping.csv",
    "User-Department-Mapping.xlsx",
    "GAL_Clean.csv",
)


def _mapping_value(normalized: dict[str, Any], names: tuple[str, ...]) -> str | None:
    for name in names:
        value = _column_value(normalized, name)
        if value is not None:
            return value
    return None


def _read_mapping_rows(path: Path) -> list[dict[str, Any]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as handle:
            return [dict(row) for row in csv.DictReader(handle)]
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        from openpyxl import load_workbook

        workbook = load_workbook(path, read_only=True, data_only=True)
        try:
            return list(_read_worksheet_rows(workbook.active))
        finally:
            workbook.close()
    return []


def load_department_mapping(
    path: Path, org_mapping: OrgMapping | None = None,
) -> dict[str, dict[str, Any]]:
    """User key (uppercased) -> department mapping rows from a GAL/department file.

    Each row is registered under BOTH its UserPrincipalName and its Mail
    address: Activity Explorer identifies users by primary SMTP address,
    which routinely lives on a different domain than the UPN (e.g. UPN
    user@qfes.qld.gov.au vs mail user@fire.qld.gov.au). Keying on UPN alone
    silently drops those users into 'Unmapped' at fact grain. UPN keys win
    collisions; mail-derived entries carry ``is_alias`` so dim_user seeding
    can skip them (one dim_user row per person, not per address).

    Department names are case-canonicalized across the file (most frequent
    casing wins) so variants like 'qfes'/'QFES' cannot split into two
    dim_department rows — Power BI's case-insensitive engine merges such
    labels and displays an arbitrary casing. Division values are
    canonicalized the same way.

    Org enrichment (user grain, surfaced on dim_user) is governed by
    ``org_mapping`` (built-in vanilla defaults when None — see
    org_mapping.default_org_mapping):
    - ``department`` / ``user_division`` / ``user_region`` / ``job_title``:
      value fields ('Unmapped'/'Unknown'/'Unknown'/None when nothing matches).
    - ``is_leaver`` / ``is_generic_account``: flag fields (default False).
    Config-supplied rules referencing columns missing from the file raise
    EnrichmentError (fail loudly); built-in default rules degrade silently.
    """
    mapping = org_mapping or default_org_mapping()
    rows = _read_mapping_rows(path)
    _validate_org_columns(mapping, rows, path)

    parsed: list[tuple[str | None, str | None, dict[str, Any]]] = []
    casing_counts: dict[str, dict[str, int]] = {}
    division_casing_counts: dict[str, dict[str, int]] = {}
    for row in rows:
        normalized = _normalized_row(row)
        upn = _mapping_value(normalized, ("userprincipalname", "user_upn", "upn", "user", "username"))
        mail = _mapping_value(normalized, ("mail", "email", "primarysmtpaddress"))
        if not upn and not mail:
            continue
        department = _extract_org_value(mapping.fields["Department"], normalized) or "Unmapped"
        counts = casing_counts.setdefault(department.casefold(), {})
        counts[department] = counts.get(department, 0) + 1
        user_division = _extract_org_value(mapping.fields["Division"], normalized) or "Unknown"
        division_counts = division_casing_counts.setdefault(user_division.casefold(), {})
        division_counts[user_division] = division_counts.get(user_division, 0) + 1
        parsed.append((upn, mail, {
            "department": department,
            "division": _mapping_value(normalized, ("division", "directorate", "group")),
            "business_unit": _mapping_value(
                normalized, ("business_unit", "businessunit", "unit", "section", "branch", "team")
            ),
            "user_division": user_division,
            "user_region": _extract_org_value(mapping.fields["Region"], normalized) or "Unknown",
            "job_title": _extract_org_value(mapping.fields["JobTitle"], normalized),
            "is_leaver": _extract_org_flag(mapping.fields["IsLeaver"], normalized),
            "is_generic_account": _extract_org_flag(mapping.fields["IsGenericAccount"], normalized),
            "mapping_source": path.name,
            "is_mapped": True,
        }))

    canonical_casing = {
        folded: max(counts.items(), key=lambda item: item[1])[0]
        for folded, counts in casing_counts.items()
    }
    canonical_division_casing = {
        folded: max(counts.items(), key=lambda item: item[1])[0]
        for folded, counts in division_casing_counts.items()
    }

    mappings: dict[str, dict[str, Any]] = {}
    for upn, mail, entry in parsed:
        entry["department"] = canonical_casing[entry["department"].casefold()]
        entry["user_division"] = canonical_division_casing[entry["user_division"].casefold()]
        primary = upn or mail
        mappings[primary.upper().strip()] = entry
    # Mail aliases never displace a primary (UPN) key.
    for upn, mail, entry in parsed:
        if not (upn and mail):
            continue
        alias_key = mail.upper().strip()
        if alias_key and alias_key not in mappings:
            mappings[alias_key] = {**entry, "is_alias": True}
    return mappings


def resolve_department_csv(input_dir: Path, explicit: Path | None) -> Path | None:
    if explicit is not None:
        if not explicit.exists():
            raise EnrichmentError(f"--department-csv does not exist: {explicit}")
        return explicit
    for name in _DEPARTMENT_FILE_CANDIDATES:
        found = _search_one(input_dir, name)
        if found is not None:
            return found
    return None


def resolve_enrichment_inputs(
    input_dir: Path,
    risk_workbook: Path | None,
    department_csv: Path | None,
    allow_unenriched: bool,
) -> tuple[Path | None, Path | None]:
    """Resolve both enrichment inputs, enforcing the F1 hard-fail policy."""
    risk_path = resolve_risk_workbook(input_dir, risk_workbook)
    dept_path = resolve_department_csv(input_dir, department_csv)

    missing = []
    if risk_path is None:
        missing.append("SIT risk workbook (SIT-Risk-Analysis*.xlsx)")
    if dept_path is None:
        missing.append(
            "department mapping (" + ", ".join(_DEPARTMENT_FILE_CANDIDATES) + ")"
        )
    if missing and not allow_unenriched:
        raise EnrichmentError(
            "Enrichment inputs not found: "
            + "; ".join(missing)
            + f". Searched {input_dir} and one level below. "
            "Pass --risk-workbook/--department-csv, or --allow-unenriched to "
            "deliberately produce an unenriched model (risk scores will be 0)."
        )
    return risk_path, dept_path
