"""Playwright dependency check, browser launch, and shutdown helpers."""

from __future__ import annotations

import sys
import threading

from .constants import USER_AGENT, log


def check_deps() -> None:
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
    except ImportError:
        print("playwright is not installed. Run:")
        print("  pip install playwright && playwright install chromium")
        sys.exit(1)


def cleanup_playwright(browser, pw) -> None:
    """Best-effort Playwright shutdown with timeout fallback.

    browser.close() / pw.stop() can hang on Windows; run them on a daemon
    thread and join with a 5-second deadline.
    """
    def _do_close():
        try:
            browser.close()
        except Exception:
            pass
        try:
            pw.stop()
        except Exception:
            pass

    t = threading.Thread(target=_do_close, daemon=True)
    t.start()
    t.join(timeout=5)
    if t.is_alive():
        log.info("Browser close timed out -- forcing exit.")


def launch_browser(pw, storage_state=None):
    """Create browser + context with standard config. Returns (browser, context)."""
    browser = pw.chromium.launch(headless=False)
    kwargs = {"viewport": {"width": 1400, "height": 900}, "user_agent": USER_AGENT}
    if storage_state:
        kwargs["storage_state"] = storage_state
    context = browser.new_context(**kwargs)
    return browser, context
