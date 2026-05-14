"""Classifier dedup + fetch orchestration across the IPML endpoints."""

from __future__ import annotations

from .api import SessionExpired, api_call, extract_list, get_auth_tokens
from .constants import (
    TC_API_AGGREGATES,
    TC_API_GETALL,
    TC_API_METADATA,
    TYPE_SIT,
    TYPE_TC,
    log,
)


def dedupe_classifiers(metadata_items: list[dict], getall_items: list[dict]) -> list[dict]:
    """Merge and deduplicate classifiers from both endpoints.

    getAll returns one row per language variant (e.g. Targeted Harassment
    has 12 rows: en, fr, de, ...). We collapse language variants into a
    single row with a Languages list. ModelMetadata fields (type, subType,
    businessFunction, etc.) are merged in by ID.
    """
    meta_by_id = {}
    for item in metadata_items:
        mid = item.get("id") or item.get("Id") or item.get("ModelId")
        if mid:
            meta_by_id[mid] = item

    grouped: dict[str, dict] = {}
    for item in getall_items:
        cid = item.get("Id") or item.get("ModelId") or ""
        lang = item.get("Language", "")

        if cid not in grouped:
            row = dict(item)
            row["Languages"] = [lang] if lang else []
            grouped[cid] = row
        else:
            if lang and lang not in grouped[cid]["Languages"]:
                grouped[cid]["Languages"].append(lang)

    meta_fields = ("type", "subType", "businessFunction", "applications",
                   "allowedLanguages", "versions", "isDeprecated")
    for cid, row in grouped.items():
        meta = meta_by_id.get(cid)
        if meta:
            for field in meta_fields:
                if field in meta and field not in row:
                    row[field] = meta[field]

    results = list(grouped.values())

    for r in results:
        langs = r.get("Languages", [])
        r["Languages"] = ", ".join(sorted(langs)) if langs else ""

    # Strip noisy OData type annotation columns
    odata_keys = [k for k in (results[0] if results else {}) if "@odata" in k or "@is." in k]
    for r in results:
        for k in odata_keys:
            r.pop(k, None)

    return results


def fetch_classifiers(page, context, include_sits: bool = False) -> list[dict]:
    """Fetch trainable classifiers from the Purview IPML service.

    Uses the same endpoints the TC management page uses:
      ModelMetadata (GET)            -> classifier definitions with types
      CategoryTrainingModel/getAll   -> full classifier list with status

    Results are deduplicated and language variants collapsed.
    """
    xsrf, tid = get_auth_tokens(context)
    if not xsrf:
        raise RuntimeError("XSRF-TOKEN not found in cookies -- session invalid")

    metadata_items: list[dict] = []
    log.info("Fetching trainable classifier metadata...")
    metadata_url = f"{TC_API_METADATA}?type=GlobalOOB%2CCustomizedOOB%2CCategoryModel"
    try:
        metadata_items = extract_list(api_call(page, "GET", metadata_url, xsrf, tid))
        log.info("  ModelMetadata: %d entries", len(metadata_items))
    except RuntimeError as e:
        log.warning("ModelMetadata call failed: %s", e)

    getall_items: list[dict] = []
    log.info("Fetching full classifier list (getAll)...")
    try:
        getall_items = extract_list(api_call(page, "POST", TC_API_GETALL, xsrf, tid))
        log.info("  getAll: %d entries (pre-dedup)", len(getall_items))
    except RuntimeError as e:
        log.warning("getAll call failed: %s", e)

    results = dedupe_classifiers(metadata_items, getall_items)
    log.info("  Deduplicated: %d unique classifiers", len(results))

    for r in results:
        r["_Type"] = TYPE_TC

    if include_sits:
        log.info("Fetching SIT aggregates from Content Explorer...")
        sit_url = f"{TC_API_AGGREGATES}?type=SensitiveInformationType&fetchTagNames=true"
        try:
            sit_data = api_call(page, "GET", sit_url, xsrf, tid)
            sits = sit_data.get("Aggregates", [])
            log.info("  %d sensitive information types", len(sits))
            results.extend([{"_Type": TYPE_SIT, **s} for s in sits])
        except RuntimeError as e:
            log.warning("SIT aggregates call failed: %s", e)

    return results


# Re-export SessionExpired so callers can `from .classifiers import SessionExpired`
__all__ = ["dedupe_classifiers", "fetch_classifiers", "SessionExpired"]
