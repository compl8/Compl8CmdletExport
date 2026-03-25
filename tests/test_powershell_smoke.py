from __future__ import annotations

import json
import subprocess
import textwrap
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = REPO_ROOT / "Modules" / "Compl8ExportFunctions.psm1"
CONFIG_PATH = REPO_ROOT / "ConfigFiles" / "ContentExplorerClassifiers.json"
MENU_PART_PATH = REPO_ROOT / "App" / "Host" / "Menu.ps1"
SCRIPT_PARTS_ROOT = REPO_ROOT / "App"
MODULE_PARTS_ROOT = REPO_ROOT / "Modules" / "Compl8ExportFunctions"


def run_pwsh(script: str) -> str:
    completed = subprocess.run(
        ["pwsh", "-NoProfile", "-Command", "-"],
        input=script,
        text=True,
        capture_output=True,
        cwd=REPO_ROOT,
        check=False,
    )
    if completed.returncode != 0:
        raise AssertionError(
            "PowerShell command failed\n"
            f"stdout:\n{completed.stdout}\n"
            f"stderr:\n{completed.stderr}"
        )
    return completed.stdout


def test_prerequisite_gate_rejects_outdated_exchange_module() -> None:
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        function Get-Module {{
            param(
                [switch]$ListAvailable,
                [string]$Name
            )
            if ($ListAvailable -and $Name -eq 'ExchangeOnlineManagement') {{
                return [pscustomobject]@{{ Version = [version]'3.1.0' }}
            }}
            return $null
        }}

        $result = Test-ExportPrerequisites
        Write-Output ('RESULT=' + $result)
        """
    )

    output = run_pwsh(script)
    assert "RESULT=False" in output


def test_resolve_ce_page_size_uses_settings_block(tmp_path: Path) -> None:
    export_dir = tmp_path / "Export-20260324-000000"
    export_dir.mkdir()

    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $result = Resolve-CEPageSize -ExportRunDirectory '{export_dir}' -ConfigPath '{CONFIG_PATH}' -FallbackPageSize 100
        Write-Output ('PAGESIZE=' + $result.PageSize)
        """
    )

    output = run_pwsh(script)
    assert "PAGESIZE=1000" in output


def test_safe_directory_names_disambiguate_sanitized_collisions() -> None:
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $first = ConvertTo-SafeDirectoryName -Name 'Policy/A'
        $second = ConvertTo-SafeDirectoryName -Name 'Policy:A'
        Write-Output ('FIRST=' + $first)
        Write-Output ('SECOND=' + $second)
        Write-Output ('EQUAL=' + ($first -eq $second))
        """
    )

    output = run_pwsh(script)
    assert "EQUAL=False" in output


def test_classifier_dir_prefers_existing_legacy_path_for_resume(tmp_path: Path) -> None:
    export_dir = tmp_path / "Export-20260324-010000"
    legacy_dir = export_dir / "Data" / "ContentExplorer" / "Sensitivity" / "Parent_Child"
    legacy_dir.mkdir(parents=True)

    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $resolved = Get-CEClassifierDir -ExportDir '{export_dir}' -TagType 'Sensitivity' -TagName 'Parent/Child'
        Write-Output ('RESOLVED=' + $resolved)
        """
    )

    output = run_pwsh(script)
    assert f"RESOLVED={legacy_dir}" in output


def test_split_powershell_sections_parse_cleanly() -> None:
    script = textwrap.dedent(
        f"""
        $files = Get-ChildItem -Path '{SCRIPT_PARTS_ROOT}', '{MODULE_PARTS_ROOT}' -Recurse -Filter *.ps1
        $failed = @()

        foreach ($file in $files) {{
            $tokens = $null
            $errors = $null
            [void][System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)

            if ($errors.Count -gt 0) {{
                $failed += ($file.FullName + ': ' + (($errors | ForEach-Object {{ $_.ToString() }}) -join ' | '))
            }}
        }}

        if ($failed.Count -gt 0) {{
            $failed | ForEach-Object {{ Write-Output $_ }}
            exit 1
        }}

        Write-Output 'PARSE=OK'
        """
    )

    output = run_pwsh(script)
    assert "PARSE=OK" in output


def test_build_auth_parameters_reads_root_level_auth_config(tmp_path: Path) -> None:
    script_root = tmp_path / "portable-root"
    config_dir = script_root / "ConfigFiles"
    config_dir.mkdir(parents=True)
    (config_dir / "AuthConfig.json").write_text(
        json.dumps(
            {
                "UseCertificateAuth": "True",
                "AppId": "test-app-id",
                "CertificateThumbprint": "ABC123",
                "Organization": "contoso.onmicrosoft.com",
            }
        ),
        encoding="utf-8",
    )

    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $scriptRoot = '{script_root}'
        $UserPrincipalName = $null
        . '{MENU_PART_PATH}'
        $result = Build-AuthParameters
        Write-Output ('APPID=' + $result.AppId)
        Write-Output ('ORG=' + $result.Organization)
        """
    )

    output = run_pwsh(script)
    assert "APPID=test-app-id" in output
    assert "ORG=contoso.onmicrosoft.com" in output
