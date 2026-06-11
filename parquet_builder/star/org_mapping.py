"""Config-driven org-field sourcing for dim_user (the GAL "department picker").

The dim_user SCHEMA is fixed (division/region/job_title/is_leaver/
is_generic_account always exist — the PBI model is generated from the SSOT
and is tenant-independent); only HOW those values are sourced from the GAL
is configurable, via ConfigFiles/AEStarOrgMapping.local.json or the
--org-mapping CLI argument (template: AEStarOrgMapping.example.json).

Built-in defaults are deliberately vanilla: Division mirrors Department, and
the DN-derived fields use generic AD patterns that degrade to 'Unknown'/False
on tenants without DNs, so an unconfigured tenant still renders.

Fail-loudly policy (consistent with the F1 enrichment fix): a malformed
config, an unknown field/Parse mode, or a config-referenced column missing
from the GAL file is a hard EnrichmentError — never a silent fallback.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

from .errors import EnrichmentError

# DN components are comma-separated; commas inside a value are escaped \,
_DN_COMPONENT_SPLIT = re.compile(r"(?<!\\),")

ORG_VALUE_FIELDS = ("Department", "Division", "Region", "JobTitle")
ORG_FLAG_FIELDS = ("IsLeaver", "IsGenericAccount")
ORG_FIELDS = ORG_VALUE_FIELDS + ORG_FLAG_FIELDS

_VALUE_PARSE_MODES = ("none", "ou_under")
_FLAG_PARSE_MODES = ("none", "ou_contains", "ou_name_in")

_TRUTHY_STRINGS = frozenset({"true", "yes", "y", "1"})

ORG_MAPPING_SOURCE_DEFAULTS = "defaults"


# --- DN parsing ---------------------------------------------------------------

def parse_dn_ous(dn: str | None) -> list[str]:
    """OU names from an AD distinguished name, leaf-first.

    'CN=Jo,OU=Users,OU=Central,OU=Regions,DC=x' -> ['Users', 'Central',
    'Regions']. Splits on unescaped commas only (values may contain '\\,').
    """
    if not dn or not str(dn).strip():
        return []
    ous: list[str] = []
    for component in _DN_COMPONENT_SPLIT.split(str(dn)):
        key, _, value = component.partition("=")
        if key.strip().upper() == "OU":
            ous.append(value.replace("\\,", ",").strip())
    return ous


def ou_under(dn: str | None, parent: str) -> str | None:
    """The OU directly under ``parent`` in a DN, or None.

    The DN is leaf-first, so that is the OU element immediately BEFORE the
    parent (e.g. OU=Users,OU=Central,OU=Regions,... with parent 'Regions' ->
    'Central'). None when the parent OU is absent or is itself the leaf OU.
    Matching is case-insensitive.
    """
    ous = parse_dn_ous(dn)
    lowered = [ou.casefold() for ou in ous]
    try:
        index = lowered.index(parent.casefold())
    except ValueError:
        return None
    if index == 0:
        return None
    return ous[index - 1] or None


def ou_contains(dn: str | None, needle: str) -> bool:
    """True when any OU in the DN contains ``needle`` (case-insensitive) —
    e.g. needle 'leaver' matches 'QFES Leavers', 'Leavers', ..."""
    needle_cf = needle.casefold()
    return any(needle_cf in ou.casefold() for ou in parse_dn_ous(dn))


def ou_name_in(dn: str | None, names: Iterable[str]) -> bool:
    """True when any OU in the DN exactly equals one of ``names``
    (case-insensitive) — name match, not position, so a re-parented OU does
    not silently drop the flag."""
    wanted = {name.casefold() for name in names}
    return any(ou.casefold() in wanted for ou in parse_dn_ous(dn))


# --- mapping model ------------------------------------------------------------

@dataclass(frozen=True)
class OrgFieldRule:
    """How one dim_user org field is sourced from the GAL file."""
    source: str                              # GAL column name
    parse: str = "none"                      # none | ou_under | ou_contains | ou_name_in
    arg: str | tuple[str, ...] | None = None  # parse-mode argument
    fallback: str | None = None              # raw column used when the primary yields nothing
    from_config: bool = False                # config-supplied rules get strict column validation

    def to_manifest(self) -> dict[str, Any]:
        payload: dict[str, Any] = {"Source": self.source}
        if self.parse != "none":
            payload["Parse"] = self.parse
        if self.arg is not None:
            payload["Arg"] = list(self.arg) if isinstance(self.arg, tuple) else self.arg
        if self.fallback is not None:
            payload["Fallback"] = self.fallback
        return payload


@dataclass(frozen=True)
class OrgMapping:
    """Resolved org-field rules plus provenance for the manifest."""
    fields: dict[str, OrgFieldRule]
    source_label: str = ORG_MAPPING_SOURCE_DEFAULTS

    def to_manifest(self) -> dict[str, Any]:
        return {
            "source": self.source_label,
            "fields": {name: self.fields[name].to_manifest() for name in ORG_FIELDS},
        }


def default_org_mapping() -> OrgMapping:
    """Built-in vanilla mapping used when no config file is present."""
    return OrgMapping(fields={
        "Department": OrgFieldRule(source="Department"),
        "Division": OrgFieldRule(source="Department"),  # mirrors Department
        "Region": OrgFieldRule(source="OnPremisesDN", parse="ou_under", arg="Regions"),
        "JobTitle": OrgFieldRule(source="JobTitle"),
        "IsLeaver": OrgFieldRule(source="OnPremisesDN", parse="ou_contains", arg="leaver"),
        "IsGenericAccount": OrgFieldRule(
            source="OnPremisesDN", parse="ou_name_in",
            arg=("Generic Accounts", "SharedUsers")),
    })


# --- config loading (fail loudly) ----------------------------------------------

def _org_rule_error(field: str, message: str, path: Path) -> EnrichmentError:
    return EnrichmentError(f"Org mapping config {path}: field '{field}': {message}")


def _parse_org_rule(field: str, raw: Any, path: Path) -> OrgFieldRule:
    if not isinstance(raw, dict):
        raise _org_rule_error(
            field, f"expected an object like {{\"Source\": \"<GAL column>\"}}, got {type(raw).__name__}", path)
    unknown = [key for key in raw if not key.startswith("_")
               and key not in ("Source", "Parse", "Arg", "Fallback")]
    if unknown:
        raise _org_rule_error(
            field, f"unknown key(s) {unknown}; supported: Source, Parse, Arg, Fallback", path)

    source = raw.get("Source")
    if not isinstance(source, str) or not source.strip():
        raise _org_rule_error(field, "'Source' is required and must name a GAL column", path)

    is_flag = field in ORG_FLAG_FIELDS
    valid_modes = _FLAG_PARSE_MODES if is_flag else _VALUE_PARSE_MODES
    parse = raw.get("Parse", "none")
    if parse not in valid_modes:
        kind = "flag" if is_flag else "value"
        raise _org_rule_error(
            field, f"unknown Parse mode '{parse}' for {kind} field "
                   f"(valid: {', '.join(valid_modes)})", path)

    arg = raw.get("Arg")
    if parse == "none":
        if arg is not None:
            raise _org_rule_error(field, "'Arg' is only valid with a Parse mode", path)
    elif parse == "ou_name_in":
        if (not isinstance(arg, list) or not arg
                or not all(isinstance(item, str) and item.strip() for item in arg)):
            raise _org_rule_error(
                field, "Parse 'ou_name_in' requires 'Arg' to be a non-empty list of OU names", path)
        arg = tuple(item.strip() for item in arg)
    else:  # ou_under / ou_contains
        if not isinstance(arg, str) or not arg.strip():
            raise _org_rule_error(
                field, f"Parse '{parse}' requires 'Arg' to be a non-empty string", path)
        arg = arg.strip()

    fallback = raw.get("Fallback")
    if fallback is not None:
        if is_flag:
            raise _org_rule_error(field, "'Fallback' is only valid on value fields", path)
        if not isinstance(fallback, str) or not fallback.strip():
            raise _org_rule_error(field, "'Fallback' must name a GAL column", path)
        fallback = fallback.strip()

    return OrgFieldRule(source=source.strip(), parse=parse, arg=arg,
                        fallback=fallback, from_config=True)


def load_org_mapping(path: Path) -> OrgMapping:
    """Load and validate an org-mapping config; malformed -> hard error.

    Fields omitted from the config keep their built-in default rule. Keys
    starting with '_' are metadata (repo config convention).
    """
    try:
        with path.open("r", encoding="utf-8-sig") as handle:
            payload = json.load(handle)
    except OSError as exc:
        raise EnrichmentError(f"Cannot read org mapping config {path}: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise EnrichmentError(f"Org mapping config {path} is not valid JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise EnrichmentError(
            f"Org mapping config {path}: expected a JSON object, got {type(payload).__name__}")

    fields = dict(default_org_mapping().fields)
    for key, raw in payload.items():
        if key.startswith("_"):
            continue
        if key not in ORG_FIELDS:
            raise EnrichmentError(
                f"Org mapping config {path}: unknown field '{key}' "
                f"(supported: {', '.join(ORG_FIELDS)})")
        fields[key] = _parse_org_rule(key, raw, path)
    return OrgMapping(fields=fields, source_label=str(path))


# --- extraction engine ----------------------------------------------------------

def _normalize_column(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", name.lower())


def _normalized_row(row: dict[str, Any]) -> dict[str, Any]:
    return {_normalize_column(str(key)): value for key, value in row.items() if key}


def _column_value(normalized: dict[str, Any], column: str) -> str | None:
    value = normalized.get(_normalize_column(column))
    if value is not None and str(value).strip():
        return str(value).strip()
    return None


def _extract_org_value(rule: OrgFieldRule, normalized: dict[str, Any]) -> str | None:
    """Value-field extraction: primary source (+parse), then raw fallback."""
    raw = _column_value(normalized, rule.source)
    value = ou_under(raw, rule.arg) if rule.parse == "ou_under" else raw
    if value is None and rule.fallback is not None:
        value = _column_value(normalized, rule.fallback)
    return value


def _extract_org_flag(rule: OrgFieldRule, normalized: dict[str, Any]) -> bool:
    raw = _column_value(normalized, rule.source)
    if rule.parse == "ou_contains":
        return ou_contains(raw, rule.arg)
    if rule.parse == "ou_name_in":
        return ou_name_in(raw, rule.arg)
    return raw is not None and raw.casefold() in _TRUTHY_STRINGS


def _validate_org_columns(mapping: OrgMapping, rows: list[dict[str, Any]], path: Path) -> None:
    """Config-supplied rules must reference real GAL columns (fail loudly).

    Built-in default rules are exempt: an unconfigured tenant whose GAL lacks
    e.g. OnPremisesDN still converts, with 'Unknown'/False org values.
    """
    if not rows:
        return
    available: set[str] = set()
    headers: dict[str, None] = {}
    for row in rows:
        for key in row:
            if key:
                available.add(_normalize_column(str(key)))
                headers.setdefault(str(key))
    problems = []
    for field in ORG_FIELDS:
        rule = mapping.fields[field]
        if not rule.from_config:
            continue
        for label, column in (("Source", rule.source), ("Fallback", rule.fallback)):
            if column and _normalize_column(column) not in available:
                problems.append(f"{field}.{label} column '{column}'")
    if problems:
        raise EnrichmentError(
            f"Org mapping config {mapping.source_label}: column(s) not present in "
            f"GAL file {path}: " + "; ".join(problems)
            + ". Available columns: " + ", ".join(headers))
