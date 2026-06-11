"""Raw Activity Explorer record parsing for the star converter.

Ports the robust JSON handling from the enhanced legacy fork
(build_activity_explorer_old_powerbi_data.py): page Records may arrive as a
list, a single dict (PowerShell ConvertTo-Json unwraps one-element arrays —
the F2 record-loss bug), or a JSON-encoded string; nested fields may be
JSON-encoded strings; EmailInfo/PolicyMatchInfo may be dicts or lists.
"""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from ..loaders import _load_jsonl_records

# Raw keys consumed by a typed column / dimension / nested explode. Anything
# else lands in fact_activity_detail.extra_json and the schema-drift report.
CONSUMED_KEYS = frozenset({
    "RecordIdentity", "Activity", "ActivityId", "Happened", "Timestamp",
    "User", "UserKey", "Workload", "Location", "DataPlatform", "UserType",
    "FilePath", "TargetFilePath", "ItemName", "FileName", "FileSize",
    "FileType", "FileExtension",
    "SensitiveInfoTypeData", "SensitiveInfoTypeBucketsData",
    "PolicyMatchInfo", "EmailInfo", "AttachmentDetails",
    "TargetDomain", "OriginatingDomain", "TargetUrl", "TargetURL",
    "DeviceName", "ClientIP", "Application", "ApplicationName", "Platform",
    "EnforcementMode", "RMSEncrypted", "RmsEncrypted", "PreviousFileName",
    "TargetPrinterName", "MDATPDeviceId", "JitTriggered", "JITTriggered",
    "EvidenceFile", "RemovableMediaDeviceAttributes", "EndpointOperation",
    "AuthorizedGroup", "MatchedPolicies", "DlpAuditEventMetadata",
    "SessionMetadata", "ItemMetadata", "ParentArchiveHash",
    "AgentId", "AgentName", "TargetAgentId", "TargetAgentName",
    "PlatformTargetAgentId",
    "FilePathUrl", "SourceLocationType", "DestinationLocationType",
    "Sha1", "Sha256", "ColdScanPolicyId", "PolicyVersion",
    "AssociatedAdminUnits", "SensitivityLabelIdsReferenced",
    "AppIdentity", "AppIdentityCategory", "AppIdentityGroup",
    "PurviewAIAppName", "PurviewAIAppLocation",
    "CopilotEventData", "AccessedResources",
    "EnrichedCopilotThreadOrCorrelationId", "EnrichedLLMMessageIds",
    "HasWebsearchQuery", "HasWebSearchQuery",
    "AreFilesReferenced", "AreSensitiveFilesReferenced",
})

# Any of these non-null marks a record as Copilot/AI-enriched.
COPILOT_KEYS = (
    "CopilotEventData", "AccessedResources",
    "EnrichedCopilotThreadOrCorrelationId", "EnrichedLLMMessageIds",
    "PurviewAIAppLocation", "PurviewAIAppName",
    "AppIdentity", "AppIdentityCategory", "AppIdentityGroup",
    "HasWebsearchQuery", "HasWebSearchQuery",
    "AreFilesReferenced", "AreSensitiveFilesReferenced",
)

APP_IDENTITY_KEYS = (
    "AppIdentity", "AppIdentityCategory", "AppIdentityGroup", "PurviewAIAppName",
)


def parse_json_string(raw: Any) -> Any:
    """Decode values that are JSON serialized inside a string; pass through otherwise."""
    if not isinstance(raw, str):
        return raw
    text = raw.strip()
    if not text or text[0] not in "[{":
        return raw
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return raw


def as_list(value: Any) -> list[dict[str, Any]]:
    """Normalize a nested field (None/dict/list/JSON string) to a list of dicts."""
    value = parse_json_string(value)
    if value is None or value == "":
        return []
    if isinstance(value, dict):
        return [value]
    if isinstance(value, list):
        rows = []
        for item in value:
            item = parse_json_string(item)
            if isinstance(item, dict):
                rows.append(item)
        return rows
    return []


def as_bool(value: Any) -> bool | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return value
    return str(value).strip().upper() in {"TRUE", "YES", "Y", "1"}


def json_str(value: Any) -> str | None:
    """Serialize a raw value for a string parquet column (JSON for complex)."""
    if value is None:
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (int, float)):
        return str(value)
    return json.dumps(value, ensure_ascii=False, default=str)


def parse_dt(value: Any) -> datetime | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        parsed = datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError:
        return None
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    return parsed


def date_key_of(value: datetime | None) -> int | None:
    if value is None:
        return None
    return int(value.strftime("%Y%m%d"))


def extract_domain_from_email(email: str | None) -> str | None:
    if not email or "@" not in email:
        return None
    return email.split("@", 1)[1].lower().strip() or None


def extract_domain_from_url(value: Any) -> str | None:
    """Extract a host/domain from URL-like Activity Explorer fields."""
    if not value:
        return None
    text = str(value).strip()
    if not text or (" " in text and "://" not in text):
        return None
    try:
        parsed = urlparse(text if "://" in text else f"//{text}")
    except ValueError:
        # e.g. bracketed-but-invalid IPv6 netloc; not a usable domain
        return None
    host = parsed.netloc or parsed.path.split("/", 1)[0]
    host = host.split("@")[-1].split(":")[0].strip().lower()
    if not host or "." not in host or "\\" in host:
        return None
    return host


def split_receivers(receivers: Any) -> list[str]:
    """Return receiver email addresses from list or delimited string shapes."""
    receivers = parse_json_string(receivers)
    if receivers is None:
        return []
    if isinstance(receivers, list):
        values = receivers
    else:
        values = re.split(r"[;,]", str(receivers))
    return [str(v).strip() for v in values if str(v).strip()]


def derive_target_domain(raw: dict[str, Any], email_row: dict[str, Any] | None,
                         user_domain: str | None) -> str | None:
    """Derive the target domain when the export has no TargetDomain field.

    Ported from the legacy fork; the old report's domain-flow pages rely on it.
    """
    explicit = raw.get("TargetDomain")
    if explicit and explicit != "None" and str(explicit).strip():
        return str(explicit).strip().lower()

    for key in ("TargetUrl", "TargetURL", "FilePathUrl", "ObjectId", "FilePath", "ItemName"):
        domain = extract_domain_from_url(raw.get(key))
        if domain:
            return domain

    if not email_row:
        return None

    sender_domain = extract_domain_from_email(str(email_row.get("Sender") or ""))
    internal_domains = {d for d in (user_domain, sender_domain) if d}
    receiver_domains = []
    for receiver in split_receivers(email_row.get("Receivers")):
        domain = extract_domain_from_email(receiver)
        if domain:
            receiver_domains.append(domain)

    for domain in receiver_domains:
        if domain not in internal_domains:
            return domain
    return receiver_domains[0] if receiver_domains else None


def load_page(path: Path) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Load one AE page: (records, wrapper metadata).

    Handles .jsonl (one record per line), JSON wrapper pages ({"Records": ...}
    where Records may be a list, a single dict, or a JSON string), and flat
    JSON lists. Returns ([], {}) for unreadable pages.
    """
    if path.suffix.lower() == ".jsonl":
        return _load_jsonl_records(path), {}

    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            data = json.load(handle)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        print(f"  WARNING: Skipping malformed JSON page: {path.name} ({exc})")
        return [], {}

    data = parse_json_string(data)
    meta: dict[str, Any] = {}
    if isinstance(data, dict):
        meta = {
            "page_number": data.get("PageNumber"),
            "export_timestamp": data.get("ExportTimestamp"),
            "watermark": data.get("WaterMark"),
            "record_count": data.get("RecordCount"),
        }
        records = parse_json_string(data.get("Records"))
    else:
        records = data

    if isinstance(records, dict):
        records = [records]
    if not isinstance(records, list):
        if records not in (None, ""):
            print(f"  WARNING: Unsupported Records shape in {path.name}: {type(records).__name__}")
        return [], meta

    rows = []
    for item in records:
        item = parse_json_string(item)
        if isinstance(item, dict):
            rows.append(item)
    return rows, meta
