"""Trainable-classifier scraper package (vendored from GetTCs).

Public surface mirrors the original single-file get_tcs.py so any external
callers that imported by name still work.
"""

from .api import SessionExpired, api_call, extract_list, get_auth_tokens
from .auth import (
    dismiss_popups,
    do_login,
    do_resume,
    load_state,
    save_state,
    wait_for_purview,
)
from .browser import check_deps, cleanup_playwright, launch_browser
from .classifiers import dedupe_classifiers, fetch_classifiers
from .cli import main
from .constants import (
    DEFAULT_COMPL8_OUT,
    PROJECT_ROOT,
    PURVIEW_URL,
    STATE_FILE,
    TC_API_AGGREGATES,
    TC_API_GETALL,
    TC_API_METADATA,
    TC_PAGE_URL,
    TYPE_SIT,
    TYPE_TC,
    USER_AGENT,
    log,
)
from .writers import print_table, write_compl8_format, write_csv, write_json

__all__ = [
    "DEFAULT_COMPL8_OUT",
    "PROJECT_ROOT",
    "PURVIEW_URL",
    "STATE_FILE",
    "SessionExpired",
    "TC_API_AGGREGATES",
    "TC_API_GETALL",
    "TC_API_METADATA",
    "TC_PAGE_URL",
    "TYPE_SIT",
    "TYPE_TC",
    "USER_AGENT",
    "api_call",
    "check_deps",
    "cleanup_playwright",
    "dedupe_classifiers",
    "dismiss_popups",
    "do_login",
    "do_resume",
    "extract_list",
    "fetch_classifiers",
    "get_auth_tokens",
    "launch_browser",
    "load_state",
    "log",
    "main",
    "print_table",
    "save_state",
    "wait_for_purview",
    "write_compl8_format",
    "write_csv",
    "write_json",
]
