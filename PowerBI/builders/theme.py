"""Theme build/embedding, custom-visual packaging, and report chrome.

The Compl8 theme (PowerBI/themes/Compl8.Theme.json) is CY26SU02 (the base
theme shipped inside the reference report) deep-merged with the curated
ContentExplorerSITRisk palette. Generated projects embed BOTH: CY26SU02 as
the SharedResources base theme and Compl8 as a RegisteredResources custom
theme, referenced from Report/config.json themeCollection.
"""

from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

THEMES_DIR = Path(__file__).resolve().parents[1] / "themes"
CUSTOM_VISUALS_DIR = Path(__file__).resolve().parents[1] / "CustomVisuals"
BASE_THEME_NAME = "CY26SU02"
CUSTOM_THEME_FILE = "Compl8.Theme.json"
THEME_NAME = "Compl8"


def deep_merge(base: dict, overlay: dict) -> dict:
    """Recursive dict merge; overlay wins on conflicts."""
    merged = dict(base)
    for key, value in overlay.items():
        if key in merged and isinstance(merged[key], dict) and isinstance(value, dict):
            merged[key] = deep_merge(merged[key], value)
        else:
            merged[key] = value
    return merged


def build_merged_theme(base_path: Path, overlay_path: Path, out_path: Path) -> dict:
    """Merge a curated palette over a base theme and write the result."""
    base = json.loads(base_path.read_text(encoding="utf-8"))
    overlay = json.loads(overlay_path.read_text(encoding="utf-8"))
    merged = deep_merge(base, overlay)
    merged["name"] = THEME_NAME
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(json.dumps(merged, indent=2), encoding="utf-8", newline="\n")
    return merged


def theme_collection() -> dict[str, object]:
    """Report/config.json themeCollection: CY26SU02 base + Compl8 custom."""
    return {
        "baseTheme": {
            "name": BASE_THEME_NAME,
            "type": 2,
            "version": {"visual": "2.6.0", "report": "3.1.0", "page": "2.3.0"},
        },
        "customTheme": {"name": CUSTOM_THEME_FILE, "type": 1},
    }


def embed_theme(out_dir: Path) -> list[dict[str, object]]:
    """Copy theme files into StaticResources/; return resourcePackages entries."""
    base_src = THEMES_DIR / f"{BASE_THEME_NAME}.json"
    custom_src = THEMES_DIR / CUSTOM_THEME_FILE
    for source in (base_src, custom_src):
        if not source.exists():
            raise FileNotFoundError(f"Theme file not found: {source}")
    base_dest = out_dir / "StaticResources" / "SharedResources" / "BaseThemes" / f"{BASE_THEME_NAME}.json"
    custom_dest = out_dir / "StaticResources" / "RegisteredResources" / CUSTOM_THEME_FILE
    base_dest.parent.mkdir(parents=True, exist_ok=True)
    custom_dest.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(base_src, base_dest)
    shutil.copyfile(custom_src, custom_dest)
    return [
        {
            "resourcePackage": {
                "disabled": False,
                "items": [
                    {
                        "name": CUSTOM_THEME_FILE,
                        "path": CUSTOM_THEME_FILE,
                        "type": 202,
                    }
                ],
                "name": "RegisteredResources",
                "type": 1,
            }
        },
        {
            "resourcePackage": {
                "disabled": False,
                "items": [
                    {
                        "name": BASE_THEME_NAME,
                        "path": f"BaseThemes/{BASE_THEME_NAME}.json",
                        "type": 202,
                    }
                ],
                "name": "SharedResources",
                "type": 2,
            }
        },
    ]


def copy_custom_visuals(out_dir: Path, guids: list[str]) -> list[dict[str, object]]:
    """Package-on-demand: vendor only the custom visuals a project uses.

    Returns the report.json resourcePackages entries for the copied visuals.
    """
    entries: list[dict[str, object]] = []
    for guid in guids:
        source = CUSTOM_VISUALS_DIR / guid
        if not source.exists():
            raise FileNotFoundError(f"Custom visual resources not found: {source}")
        destination = out_dir / "CustomVisuals" / guid
        if destination.exists():
            shutil.rmtree(destination)
        shutil.copytree(source, destination)
        entries.append({
            "resourcePackage": {
                "disabled": False,
                "items": [
                    {
                        "name": f"{guid}.pbiviz.json",
                        "path": f"{guid}.pbiviz.json",
                        "type": 5,
                    }
                ],
                "name": guid,
                "type": 0,
            }
        })
    return entries


def linguistic_schema_xml() -> str:
    """Minimal valid linguistic schema (Q&A); entities are optional."""
    return (
        '<LinguisticSchema xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
        'Language="en-US" DynamicImprovement="Default" '
        'xmlns="http://schemas.microsoft.com/sqlserver/2016/01/linguisticschema" />'
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Rebuild PowerBI/themes/Compl8.Theme.json from a base theme and a palette overlay."
    )
    parser.add_argument("--base", default=str(THEMES_DIR / f"{BASE_THEME_NAME}.json"),
                        help="Base theme JSON (default: vendored CY26SU02).")
    parser.add_argument("--overlay", required=True,
                        help="Curated palette/style overlay JSON.")
    parser.add_argument("--output", default=str(THEMES_DIR / CUSTOM_THEME_FILE),
                        help="Merged theme output path.")
    args = parser.parse_args()
    build_merged_theme(Path(args.base), Path(args.overlay), Path(args.output))
    print(f"Wrote {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
