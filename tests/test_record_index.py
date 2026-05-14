"""Tests for the record_index emission in parquet_builder.content.process_content."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from parquet_builder.content import process_content


def _make_export(root: Path) -> Path:
    """Build a tiny synthetic export directory with mixed JSON and JSONL CE pages."""
    ce_root = root / "Data" / "ContentExplorer" / "SensitiveInformationType" / "CreditCard"
    ce_root.mkdir(parents=True)

    # Legacy JSON wrapper page (page_offset = 0-indexed position in Records)
    json_page = ce_root / "SharePoint-001.json"
    json_page.write_text(json.dumps({
        "PageNumber": 1,
        "TagType": "SensitiveInformationType",
        "TagName": "CreditCard",
        "Workload": "SharePoint",
        "Records": [
            {
                "FileUrl": "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/a.docx",
                "FileSourceUrl": "https://contoso.sharepoint.com/sites/HR",
                "Location": "SPO",
                "FileName": "a.docx",
                "SensitiveInfoTypes": "abc,def",
                "_ExportTagType": "SensitiveInformationType",
                "_ExportTagName": "CreditCard",
            },
            {
                "FileUrl": "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/b.docx",
                "FileSourceUrl": "https://contoso.sharepoint.com/sites/HR",
                "Location": "SPO",
                "FileName": "b.docx",
                "SensitiveInfoTypes": "abc",
                "_ExportTagType": "SensitiveInformationType",
                "_ExportTagName": "CreditCard",
            },
        ],
    }), encoding="utf-8")

    # JSONL page (page_offset = 1-indexed line number)
    jsonl_page = ce_root / "SharePoint-002.jsonl"
    jsonl_page.write_text(
        "\n".join([
            json.dumps({
                "FileUrl": "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/c.docx",
                "FileSourceUrl": "https://contoso.sharepoint.com/sites/HR",
                "Location": "SPO",
                "FileName": "c.docx",
                "SensitiveInfoTypes": "ghi",
                "_ExportTagType": "SensitiveInformationType",
                "_ExportTagName": "CreditCard",
            }),
            "",  # blank line that the loader must tolerate
            json.dumps({
                "FileUrl": "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/d.docx",
                "FileSourceUrl": "https://contoso.sharepoint.com/sites/HR",
                "Location": "SPO",
                "FileName": "d.docx",
                "SensitiveInfoTypes": "jkl",
                "_ExportTagType": "SensitiveInformationType",
                "_ExportTagName": "CreditCard",
            }),
        ]),
        encoding="utf-8",
    )

    return root


def test_record_index_captures_page_file_and_offset(tmp_path: Path) -> None:
    export_dir = _make_export(tmp_path / "Export-test")
    content_files, sit_detections, record_index = process_content(export_dir)

    assert len(content_files) == 4, "expected 4 content_files records"
    assert len(record_index) == 4, "record_index must have one row per content_files record"

    # Map by file_url for assertion ergonomics
    by_url = {r["file_url"]: r for r in record_index}
    assert set(by_url) == {
        "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/a.docx",
        "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/b.docx",
        "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/c.docx",
        "https://contoso.sharepoint.com/sites/HR/Shared%20Documents/d.docx",
    }

    # JSON wrapper page: offsets are 0-indexed array positions
    a = by_url["https://contoso.sharepoint.com/sites/HR/Shared%20Documents/a.docx"]
    b = by_url["https://contoso.sharepoint.com/sites/HR/Shared%20Documents/b.docx"]
    assert a["page_offset"] == 0
    assert b["page_offset"] == 1
    assert a["page_file"].replace("\\", "/").endswith("SharePoint-001.json")
    assert b["page_file"] == a["page_file"]

    # JSONL page: offsets are 1-indexed line numbers, blank line skipped
    c = by_url["https://contoso.sharepoint.com/sites/HR/Shared%20Documents/c.docx"]
    d = by_url["https://contoso.sharepoint.com/sites/HR/Shared%20Documents/d.docx"]
    assert c["page_offset"] == 1
    assert d["page_offset"] == 3  # line 1 = c, line 2 = blank, line 3 = d
    assert c["page_file"].replace("\\", "/").endswith("SharePoint-002.jsonl")

    # All rows share the same export root
    assert all(r["_source_export_dir"] == str(export_dir) for r in record_index)

    # Tag metadata propagated. CE records use `Location` ("SPO" etc.) as the workload
    # source via CONTENT_RENAMES, so the workload field reflects that.
    assert all(r["tag_name"] == "CreditCard" for r in record_index)
    assert all(r["workload"] == "SPO" for r in record_index)

    # Page file is RELATIVE to the export root (portable across machines)
    for r in record_index:
        assert not Path(r["page_file"]).is_absolute(), f"page_file should be relative: {r['page_file']}"


def test_record_index_round_trip_via_raw_lookup(tmp_path: Path) -> None:
    """Index page_file + page_offset must let us re-read the original record."""
    export_dir = _make_export(tmp_path / "Export-test")
    _, _, record_index = process_content(export_dir)
    b = next(r for r in record_index if r["file_url"].endswith("/b.docx"))
    # Re-read using the index's pointer
    page_path = export_dir / b["page_file"]
    with open(page_path, "r", encoding="utf-8-sig") as f:
        page_data = json.load(f)
    original = page_data["Records"][b["page_offset"]]
    assert original["FileUrl"] == b["file_url"]
    assert original["FileName"] == "b.docx"

    # Same round-trip but for the JSONL page
    d = next(r for r in record_index if r["file_url"].endswith("/d.docx"))
    page_path = export_dir / d["page_file"]
    with open(page_path, "r", encoding="utf-8-sig") as f:
        for i, line in enumerate(f, start=1):
            if i == d["page_offset"]:
                original = json.loads(line.strip())
                break
        else:
            pytest.fail("page_offset overran the file")
    assert original["FileName"] == "d.docx"
