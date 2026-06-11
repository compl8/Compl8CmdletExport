"""Shared error types for the star pipeline."""

from __future__ import annotations


class EnrichmentError(RuntimeError):
    """Raised when required enrichment inputs cannot be resolved or parsed."""
