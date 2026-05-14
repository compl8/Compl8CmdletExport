# Helpers

Project-adjacent utilities. Self-contained and not imported by the main export pipeline; you run them as separate steps.

## Get-TrainableClassifiers.py

Pulls the list of trainable classifiers from the Microsoft Purview portal and writes a cache that the Content Explorer orchestrator uses to auto-discover `TrainableClassifier` tag names.

Microsoft has not shipped a public cmdlet or Graph API for enumerating trainable classifiers. This helper authenticates through the Purview portal in a real browser and calls the same internal `apiproxy` endpoints the TC management page uses.

### One-time setup

```powershell
pip install playwright
playwright install chromium
```

### Refreshing the cache

```powershell
python Helpers/Get-TrainableClassifiers.py
```

What happens:

1. A Chromium window opens at the Purview "Trainable classifiers" page.
2. You sign in (MFA / conditional access flows are supported).
3. The session is saved to `ConfigFiles/PurviewPortalAuth.local.json` (gitignored).
4. The classifier list is written to `ConfigFiles/CurrentTenantTCs.local.json` (gitignored).
5. Subsequent runs reuse the saved session until it expires.

The cache file looks like:

```json
{
  "SchemaVersion": 1,
  "DiscoveredAt": "2026-05-14T07:00:00Z",
  "TenantId": "...",
  "Source": "purview-portal",
  "ClassifierCount": 47,
  "Classifiers": [
    {
      "Id": "8aef6743-...",
      "Name": "Source code",
      "DisplayName": "Source code",
      "Type": "GlobalOOB",
      "SubType": "Regular",
      "BusinessFunction": "IP & Trade Secrets",
      "Languages": ["en"],
      "ModelStatus": "Stable",
      "IsPublished": true,
      "IsDeprecated": false
    },
    ...
  ]
}
```

The CE orchestrator's `Get-TrainableClassifiersFromCache` reads this and feeds the names into the aggregate/detail task plan. If the cache file is missing or older than 30 days the orchestrator logs a warning and proceeds without trainable classifiers.

### From inside a Compl8 export run

```powershell
.\Export-Compl8Configuration.ps1 -ContentExplorer -RefreshTrainableClassifiers
```

Refreshes the cache once before the CE export starts, then proceeds normally. Useful when you want a single command instead of two steps.

### Diagnostic / standalone output

```powershell
# CSV
python Helpers/Get-TrainableClassifiers.py -o trainable_classifiers.csv

# JSON (including SITs)
python Helpers/Get-TrainableClassifiers.py -o tcs.json --json --include-sits

# Skip the Compl8 cache write
python Helpers/Get-TrainableClassifiers.py --no-compl8-out -o tcs.csv

# Force a fresh login
python Helpers/Get-TrainableClassifiers.py --force-login
```

### Code layout

Vendored from `C:\claudecode\GetTCs\get_tcs.py`. Split into a small package:

```
Helpers/
  Get-TrainableClassifiers.py       # thin CLI entry-point
  get_trainable_classifiers/
    __init__.py                     # re-exports
    constants.py                    # URLs, paths, logging
    browser.py                      # Playwright dep check, launch, cleanup
    auth.py                         # login / resume / session state
    api.py                          # portal API calls, token extraction
    classifiers.py                  # dedup + fetch orchestration
    writers.py                      # CSV / JSON / Compl8-cache writers
    cli.py                          # argparse entry point
```
