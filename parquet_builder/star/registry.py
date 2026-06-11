"""Dimension id registry: stable surrogate keys plus accumulated dim rows.

Ported from the v5 optimized converter's IdRegistry with v6 additions:
dim_app_identity, dim_source_page provenance, raw FileExtension/FileType on
dim_file, and has_activity tracking on users.
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from ..constants import EGRESS_ACTIVITIES, SERVICE_ACCOUNT_PATTERNS
from ..helpers import _safe_str
from .keys import stable_int_id
from .records import extract_domain_from_email


def extract_folder(path: str | None) -> str | None:
    if not path:
        return None
    text = str(path).strip()
    last_sep = max(text.rfind("\\"), text.rfind("/"))
    if last_sep <= 0:
        return None
    return text[:last_sep]


def extract_file_name(path: str | None) -> str | None:
    if not path:
        return None
    text = str(path).strip()
    last_sep = max(text.rfind("\\"), text.rfind("/"))
    if last_sep < 0:
        return text or None
    name = text[last_sep + 1:]
    return name or None


def file_extension_of(name: str | None) -> str | None:
    if not name or "." not in name:
        return None
    ext = name.rsplit(".", 1)[-1].strip().lower()
    return ext or None


def parse_domain_parts(domain: str | None) -> tuple[str | None, str | None]:
    if not domain:
        return None, None
    parts = domain.lower().split(".")
    tld = parts[-1] if parts else None
    parent = ".".join(parts[-3:]) if len(parts) >= 3 else domain.lower()
    return tld, parent


def activity_group_of(activity: str | None, workload: str | None,
                      has_policy: bool, has_email: bool) -> str:
    text = f"{activity or ''} {workload or ''}".lower()
    if "copilot" in text:
        return "Copilot"
    if has_policy:
        return "DLP"
    if has_email:
        return "Email"
    if activity in EGRESS_ACTIVITIES:
        return "Egress"
    if "file" in text:
        return "File"
    return "Other"


@dataclass
class IdRegistry:
    department_map: dict[tuple, int] = field(default_factory=dict)
    department_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    user_map: dict[str, int] = field(default_factory=dict)
    user_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    activity_type_map: dict[tuple, int] = field(default_factory=dict)
    activity_type_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    workload_map: dict[str, int] = field(default_factory=dict)
    workload_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    location_map: dict[str, int] = field(default_factory=dict)
    location_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    file_map: dict[str, int] = field(default_factory=dict)
    file_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    policy_map: dict[tuple, int] = field(default_factory=dict)
    policy_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    domain_map: dict[str, int] = field(default_factory=dict)
    domain_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    email_map: dict[str, int] = field(default_factory=dict)
    email_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    app_identity_map: dict[tuple, int] = field(default_factory=dict)
    app_identity_rows: dict[int, dict[str, Any]] = field(default_factory=dict)
    page_map: dict[str, int] = field(default_factory=dict)
    page_rows: dict[int, dict[str, Any]] = field(default_factory=dict)

    def get_department(self, mapping: dict[str, Any] | None) -> int:
        mapping = mapping or {}
        department = str(mapping.get("department") or "Unmapped").strip() or "Unmapped"
        division = _safe_str(mapping.get("division"))
        business_unit = _safe_str(mapping.get("business_unit"))
        # Case-insensitive identity: 'qfes' and 'QFES' must resolve to ONE
        # dim_department row (Power BI's engine is case-insensitive and would
        # merge the labels of two case-variant rows non-deterministically).
        # Display casing comes from the first mapping seen (the GAL loader
        # canonicalizes to the most frequent casing in the file).
        key = tuple(str(part or "").casefold() for part in (department, division, business_unit))
        if key not in self.department_map:
            department_id = stable_int_id("department", "|".join(key))
            self.department_map[key] = department_id
            self.department_rows[department_id] = {
                "department_id": department_id,
                "department": department,
                "division": division,
                "business_unit": business_unit,
                "mapping_source": _safe_str(mapping.get("mapping_source")) or "No mapping file",
                "is_mapped": bool(mapping.get("is_mapped")),
            }
        return self.department_map[key]

    def get_domain(self, domain: str | None) -> int | None:
        if not domain:
            return None
        key = str(domain).lower().strip()
        if not key:
            return None
        if key not in self.domain_map:
            domain_id = stable_int_id("domain", key)
            tld, parent = parse_domain_parts(key)
            self.domain_map[key] = domain_id
            self.domain_rows[domain_id] = {
                "domain_id": domain_id,
                "domain": key,
                "parent_domain": parent,
                "tld": tld,
            }
        return self.domain_map[key]

    def get_user(self, upn: str | None, department_mappings: dict[str, dict[str, Any]],
                 has_activity: bool = True) -> tuple[int | None, int | None]:
        if not upn:
            return None, self.get_department(None)
        key = str(upn).upper().strip()
        if not key:
            return None, self.get_department(None)
        mapping = department_mappings.get(key)
        department_id = self.get_department(mapping)
        if key not in self.user_map:
            user_id = stable_int_id("user", key)
            domain = extract_domain_from_email(key)
            self.user_map[key] = user_id
            self.user_rows[user_id] = {
                "user_id": user_id,
                "user_upn": key,
                "user_domain": domain,
                "department_id": department_id,
                "is_service_account": bool(SERVICE_ACCOUNT_PATTERNS.search(key)),
                "has_activity": has_activity,
            }
            self.get_domain(domain)
        elif has_activity:
            self.user_rows[self.user_map[key]]["has_activity"] = True
        return self.user_map[key], department_id

    def get_workload(self, workload: str | None) -> int | None:
        if not workload:
            return None
        key = str(workload).strip()
        if not key:
            return None
        if key not in self.workload_map:
            workload_id = stable_int_id("workload", key.lower())
            self.workload_map[key] = workload_id
            self.workload_rows[workload_id] = {"workload_id": workload_id, "workload": key}
        return self.workload_map[key]

    def get_activity_type(self, activity: str | None, group: str) -> int | None:
        if not activity:
            return None
        key = (str(activity).strip(), group)
        if not key[0]:
            return None
        if key not in self.activity_type_map:
            activity_type_id = stable_int_id("activity_type", "|".join(key).lower())
            self.activity_type_map[key] = activity_type_id
            self.activity_type_rows[activity_type_id] = {
                "activity_type_id": activity_type_id,
                "activity": key[0],
                "activity_group": group,
                "is_egress": key[0] in EGRESS_ACTIVITIES,
                "is_copilot": "copilot" in f"{key[0]} {group}".lower(),
            }
        return self.activity_type_map[key]

    def get_location(self, folder_path: str | None) -> int | None:
        if not folder_path:
            return None
        key = str(folder_path).strip()
        if not key:
            return None
        if key not in self.location_map:
            location_id = stable_int_id("location", key.lower())
            parts = [part for part in re.split(r"[\\/]+", key) if part]
            self.location_map[key] = location_id
            self.location_rows[location_id] = {
                "location_id": location_id,
                "folder_path": key,
                "folder_name": parts[-1] if parts else key,
                "path_depth": len(parts),
            }
        return self.location_map[key]

    def get_file(self, path: str | None, item_name: str | None,
                 file_type: str | None, raw_extension: str | None = None) -> int | None:
        file_path = _safe_str(path)
        file_name = extract_file_name(file_path) or _safe_str(item_name)
        if not file_path and not file_name:
            return None
        key = (file_path or file_name or "").lower()
        if key not in self.file_map:
            file_id = stable_int_id("file", key)
            self.file_map[key] = file_id
            self.file_rows[file_id] = {
                "file_id": file_id,
                "file_path": file_path,
                "file_name": file_name,
                "file_extension": (_safe_str(raw_extension) or "").lower() or file_extension_of(file_name),
                "file_type": _safe_str(file_type),
            }
        return self.file_map[key]

    def get_policy(self, policy: dict[str, Any] | None) -> int | None:
        if not policy:
            return None
        policy_id_raw = _safe_str(policy.get("PolicyId")) or ""
        rule_id_raw = _safe_str(policy.get("RuleId")) or ""
        if not policy_id_raw and not rule_id_raw:
            return None
        key = (policy_id_raw, rule_id_raw)
        if key not in self.policy_map:
            policy_rule_id = stable_int_id("policy_rule", "|".join(key))
            actions = policy.get("RuleActions")
            if isinstance(actions, (list, dict)):
                actions = json.dumps(actions, ensure_ascii=False, default=str)
            condition = policy.get("Condition")
            if condition is None and isinstance(policy.get("OtherConditions"), dict):
                condition = policy["OtherConditions"].get("Condition")
            if isinstance(condition, (list, dict)):
                condition = json.dumps(condition, ensure_ascii=False, default=str)
            self.policy_map[key] = policy_rule_id
            self.policy_rows[policy_rule_id] = {
                "policy_rule_id": policy_rule_id,
                "policy_id": policy_id_raw,
                "policy_name": _safe_str(policy.get("PolicyName")),
                "policy_mode": _safe_str(policy.get("PolicyMode")),
                "rule_id": rule_id_raw,
                "rule_name": _safe_str(policy.get("RuleName")),
                "rule_actions": _safe_str(actions),
                "condition": _safe_str(condition),
            }
        return self.policy_map[key]

    def get_email(self, email: str | None) -> int | None:
        if not email:
            return None
        key = str(email).lower().strip()
        if not key:
            return None
        if key not in self.email_map:
            email_address_id = stable_int_id("email", key)
            domain_id = self.get_domain(extract_domain_from_email(key))
            self.email_map[key] = email_address_id
            self.email_rows[email_address_id] = {
                "email_address_id": email_address_id,
                "email_address": key,
                "local_part": key.split("@", 1)[0] if "@" in key else key,
                "domain_id": domain_id,
            }
        return self.email_map[key]

    def get_app_identity(self, app_identity: str | None, category: str | None,
                         group: str | None, ai_app_name: str | None) -> int | None:
        values = tuple(_safe_str(v) for v in (app_identity, category, group, ai_app_name))
        if not any(values):
            return None
        if values not in self.app_identity_map:
            app_identity_id = stable_int_id(
                "app_identity", "|".join(v or "" for v in values).lower()
            )
            self.app_identity_map[values] = app_identity_id
            self.app_identity_rows[app_identity_id] = {
                "app_identity_id": app_identity_id,
                "app_identity": values[0],
                "app_identity_category": values[1],
                "app_identity_group": values[2],
                "purview_ai_app_name": values[3],
            }
        return self.app_identity_map[values]

    def get_page(self, source_file: str, meta: dict[str, Any],
                 parse_dt_fn) -> int:
        if source_file not in self.page_map:
            page_id = stable_int_id("page", source_file.lower())
            self.page_map[source_file] = page_id
            record_count = meta.get("record_count")
            page_number = meta.get("page_number")
            self.page_rows[page_id] = {
                "page_id": page_id,
                "source_file": source_file,
                "page_number": int(page_number) if isinstance(page_number, (int, float)) else None,
                "export_timestamp": parse_dt_fn(meta.get("export_timestamp")),
                "watermark": _safe_str(meta.get("watermark")),
                "record_count": int(record_count) if isinstance(record_count, (int, float)) else None,
            }
        return self.page_map[source_file]
