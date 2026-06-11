"""Deterministic integer surrogate keys for the star schema.

Ported from the v5 optimized converter (C8CmdLetExportReport
tools/convert_activity_explorer_optimized_to_parquet.py). SHA1 is used as a
stable content hash, not for any security purpose — the per-process
randomization of Python's built-in hash() (and .NET's String.GetHashCode)
would otherwise produce different ids between runs/processes.
"""

from __future__ import annotations

import hashlib


def stable_int_id(namespace: str, key: str) -> int:
    """Deterministic positive 63-bit integer id for a namespaced natural key.

    The same (namespace, key) pair always yields the same id, across runs and
    machines, so ids are safe to join across incremental conversions. Zero is
    reserved (never returned) so it can be used as a sentinel if needed.
    """
    digest = hashlib.sha1(
        f"{namespace}:{key}".encode("utf-8", errors="ignore"),
        usedforsecurity=False,
    ).digest()
    value = int.from_bytes(digest[:8], "big") & 0x7FFFFFFFFFFFFFFF
    return value or 1


# Backwards-compatible alias matching the v5 converter's private name.
_stable_int_id = stable_int_id
