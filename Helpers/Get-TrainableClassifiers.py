#!/usr/bin/env python3
"""Get-TrainableClassifiers - Pull Purview trainable classifiers via the portal.

Vendored from C:\\claudecode\\GetTCs\\get_tcs.py and split into the
get_trainable_classifiers package for size/maintainability. Two changes
from the upstream script for the Compl8 export pipeline:

  1. Auth state lives in ConfigFiles/PurviewPortalAuth.local.json
     (gitignored by the *-local.json rule).
  2. --compl8-out <path> writes a Compl8-compatible classifier JSON
     (the format the CE orchestrator's TrainableClassifier discovery reads).
     The native CSV / --json outputs still work for standalone diagnostic use.

Why a portal scraper: Microsoft has not shipped a public PowerShell cmdlet
or Graph API for enumerating trainable classifiers. This script
authenticates through the Purview portal and calls the same internal
apiproxy endpoints the TC management page uses.

Usage:
    # Refresh the Compl8 cache (run before an export that needs TCs)
    python Helpers/Get-TrainableClassifiers.py

    # Diagnostic CSV (standalone)
    python Helpers/Get-TrainableClassifiers.py -o trainable_classifiers.csv

    # JSON dump including SITs (diagnostic)
    python Helpers/Get-TrainableClassifiers.py -o out.json --json --include-sits

    # Force fresh portal login
    python Helpers/Get-TrainableClassifiers.py --force-login

Requires Playwright. Install once with:
    pip install playwright
    playwright install chromium
"""

from __future__ import annotations

import sys

from get_trainable_classifiers import main

if __name__ == "__main__":
    sys.exit(main())
