"""CLI entry point for the trainable-classifier helper."""

from __future__ import annotations

import argparse
from pathlib import Path

from .api import get_auth_tokens
from .auth import do_login, do_resume, load_state
from .browser import check_deps, cleanup_playwright
from .classifiers import SessionExpired, fetch_classifiers
from .constants import DEFAULT_COMPL8_OUT, log
from .writers import print_table, write_compl8_format, write_csv, write_json


def main() -> int:
    check_deps()
    from playwright.sync_api import sync_playwright

    parser = argparse.ArgumentParser(
        description="Pull trainable classifiers from Microsoft Purview portal"
    )
    parser.add_argument("-o", "--output", default=None,
                        help="Diagnostic CSV/JSON output (default: none)")
    parser.add_argument("--json", action="store_true",
                        help="Diagnostic output is JSON instead of CSV")
    parser.add_argument("--compl8-out", default=str(DEFAULT_COMPL8_OUT),
                        help=f"Compl8 cache file path (default: {DEFAULT_COMPL8_OUT})")
    parser.add_argument("--no-compl8-out", action="store_true",
                        help="Skip writing the Compl8 cache file (diagnostic mode only)")
    parser.add_argument("--include-sits", action="store_true",
                        help="Also fetch sensitive information types (diagnostic only)")
    parser.add_argument("--force-login", action="store_true",
                        help="Force fresh login (ignore saved session)")
    args = parser.parse_args()

    state = None if args.force_login else load_state()

    pw = sync_playwright().start()
    browser = context = page = None

    try:
        if state:
            browser, context, page = do_resume(pw, state)
        else:
            browser, context, page = do_login(pw)

        try:
            results = fetch_classifiers(page, context, args.include_sits)
        except SessionExpired:
            log.warning("Session expired during API call. Re-authenticating...")
            cleanup_playwright(browser, pw)
            pw = sync_playwright().start()
            browser, context, page = do_login(pw)
            results = fetch_classifiers(page, context, args.include_sits)

        if not results:
            log.info("No classifiers found.")
            return 0

        _, tenant_id = get_auth_tokens(context)

        if not args.no_compl8_out:
            write_compl8_format(results, Path(args.compl8_out), tenant_id)

        if args.output:
            output_path = Path(args.output)
            if args.json:
                write_json(results, output_path)
            else:
                write_csv(results, output_path)

        print_table(results)
        return 0

    finally:
        if browser:
            cleanup_playwright(browser, pw)
