"""Auth state, login/resume flows, and popup dismissal for the Purview portal."""

from __future__ import annotations

import json
import sys
import time

from .browser import cleanup_playwright, launch_browser
from .constants import STATE_FILE, TC_PAGE_URL, log

POPUP_SELECTORS = [
    "button:has-text('Get started')",
    "button:has-text('Done')",
    "button[aria-label='Close']",
    "button[aria-label='Dismiss']",
    ".ms-Dialog-header button.ms-Dialog-button--close",
]


def dismiss_popups(page) -> None:
    """Click through Purview onboarding modals that block API access."""
    for sel in POPUP_SELECTORS:
        try:
            btn = page.query_selector(sel)
            if btn and btn.is_visible():
                btn.click()
                log.info("Dismissed popup: %s", sel[:60])
                page.wait_for_timeout(300)
        except Exception:
            pass


def wait_for_purview(page) -> bool:
    """Wait for login completion and TC page SPA to settle."""
    log.info("Waiting for login... (complete it in the browser window)")
    try:
        page.wait_for_url(
            "**/purview.microsoft.com/**",
            timeout=300_000,
            wait_until="domcontentloaded",
        )
    except Exception:
        log.error("Login timed out after 5 minutes.")
        return False

    log.info("Logged in. URL: %s", page.url[:120])

    if "trainableclassifiers" not in page.url.lower():
        log.info("Navigating to Trainable Classifiers page...")
        page.goto(TC_PAGE_URL, wait_until="domcontentloaded")

    log.info("Waiting for SPA to settle...")
    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        log.info("Network didn't fully idle -- proceeding.")

    time.sleep(1)
    dismiss_popups(page)
    return True


def load_state() -> dict | None:
    try:
        state = json.loads(STATE_FILE.read_text("utf-8"))
        if state.get("cookies"):
            return state
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        pass
    return None


def save_state(context) -> dict:
    """Save Playwright storage state to disk."""
    state = context.storage_state()
    state["_saved_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(json.dumps(state, indent=2), "utf-8")
    n = len(state.get("cookies", []))
    log.info("Session saved (%d cookies) -> %s", n, STATE_FILE.name)
    return state


def do_login(pw):
    """Fresh login: open Chromium, user authenticates, return (browser, context, page)."""
    browser, context = launch_browser(pw)
    page = context.new_page()

    log.info("Opening Purview login...")
    page.goto(TC_PAGE_URL, wait_until="domcontentloaded")

    if not wait_for_purview(page):
        cleanup_playwright(browser, pw)
        sys.exit(1)

    save_state(context)
    return browser, context, page


def do_resume(pw, state: dict):
    """Resume from saved state. Falls back to fresh login if session expired."""
    log.info("Resuming session from %s ...", STATE_FILE.name)
    browser, context = launch_browser(pw, storage_state=state)
    page = context.new_page()

    log.info("Navigating to Trainable Classifiers page...")
    page.goto(TC_PAGE_URL, wait_until="domcontentloaded")

    try:
        page.wait_for_load_state("networkidle", timeout=30_000)
    except Exception:
        log.info("Network didn't fully idle -- proceeding.")

    dismiss_popups(page)

    # Check if session expired (redirected to login page)
    url = page.url
    if "login.microsoftonline.com" in url or "login.live.com" in url:
        log.warning("Session expired -- redirected to login. Starting fresh...")
        cleanup_playwright(browser, pw)
        return do_login(pw)

    log.info("Session valid.")
    save_state(context)
    return browser, context, page
