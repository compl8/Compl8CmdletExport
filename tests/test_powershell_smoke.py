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


def test_content_explorer_settings_include_large_all_sit_fallbacks() -> None:
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $config = Read-JsonConfig -Path '{CONFIG_PATH}'
        $settings = Get-ContentExplorerSettings -ConfigObject $config -DefaultBatchSize 10 -DefaultWorkloads @('SharePoint','OneDrive') -DefaultPageSize 100
        Write-Output ('THRESHOLD=' + $settings.LargeAllSITDetailThreshold)
        Write-Output ('FALLBACKS=' + ($settings.LargeAllSITWorkloadFallbackWorkloads -join '|'))
        """
    )

    output = run_pwsh(script)
    threshold_lines = [line.strip() for line in output.splitlines() if line.startswith("THRESHOLD=")]
    assert threshold_lines == ["THRESHOLD=100"], f"expected THRESHOLD=100, got {threshold_lines}"
    assert "FALLBACKS=Exchange|Teams" in output


def test_large_all_sit_detail_tasks_fallback_only_for_exchange_and_teams() -> None:
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $workPlan = @(
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'Exchange'; TotalCount = 25;
                Locations = @(@{{ Name = 'user1@contoso.com'; ExpectedCount = 10 }}, @{{ Name = 'user2@contoso.com'; ExpectedCount = 15 }})
            }},
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'Teams'; TotalCount = 9;
                Locations = @(@{{ Name = 'user3@contoso.com'; ExpectedCount = 9 }})
            }},
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'SharePoint'; TotalCount = 30;
                Locations = @(@{{ Name = 'https://contoso.sharepoint.com/sites/a'; ExpectedCount = 12 }}, @{{ Name = 'https://contoso.sharepoint.com/sites/b'; ExpectedCount = 18 }})
            }},
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'OneDrive'; TotalCount = 7;
                Locations = @(@{{ Name = 'https://contoso-my.sharepoint.com/personal/user'; ExpectedCount = 7 }})
            }}
        )

        $detail = New-ContentExplorerDetailTasks -WorkPlanTasks $workPlan -DefaultPageSize 1000 -WorkloadFallbackWorkloads @('Exchange','Teams')
        $exchange = @($detail | Where-Object {{ $_.Workload -eq 'Exchange' }})
        $teams = @($detail | Where-Object {{ $_.Workload -eq 'Teams' }})
        $sharePoint = @($detail | Where-Object {{ $_.Workload -eq 'SharePoint' }})
        $oneDrive = @($detail | Where-Object {{ $_.Workload -eq 'OneDrive' }})

        Write-Output ('EXCHANGE_FALLBACK=' + (($exchange.Count -eq 1) -and $exchange[0].LocationType -eq 'WorkloadFallback' -and [string]::IsNullOrEmpty($exchange[0].Location)))
        Write-Output ('TEAMS_FALLBACK=' + (($teams.Count -eq 1) -and $teams[0].LocationType -eq 'WorkloadFallback' -and [string]::IsNullOrEmpty($teams[0].Location)))
        Write-Output ('SHAREPOINT_LOCATION=' + (($sharePoint.Count -eq 2) -and (@($sharePoint | Where-Object {{ $_.LocationType -eq 'SiteUrl' }}).Count -eq 2)))
        Write-Output ('ONEDRIVE_LOCATION=' + (($oneDrive.Count -eq 1) -and $oneDrive[0].LocationType -eq 'SiteUrl' -and -not [string]::IsNullOrEmpty($oneDrive[0].Location)))
        """
    )

    output = run_pwsh(script)
    assert "EXCHANGE_FALLBACK=True" in output
    assert "TEAMS_FALLBACK=True" in output
    assert "SHAREPOINT_LOCATION=True" in output
    assert "ONEDRIVE_LOCATION=True" in output


def test_content_explorer_export_assigns_large_all_sit_settings() -> None:
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.ps1").read_text(encoding="utf-8")
    assert "$largeAllSITDetailThreshold = $ceSettings.LargeAllSITDetailThreshold" in source
    assert "$largeAllSITFallbackCandidates = @($ceSettings.LargeAllSITWorkloadFallbackWorkloads)" in source


def test_retry_tasks_csv_round_trips_location_columns(tmp_path: Path) -> None:
    """Write-RetryTasksCsv must persist Location/LocationType so retries scope correctly."""
    csv_path = (tmp_path / "RetryTasks.csv").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $sp = [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='Credit Card'; Workload='SharePoint'; Location='https://contoso.sharepoint.com/sites/a'; LocationType='SiteUrl'; OriginalExpectedCount=1000; ActualCount=50; DiscrepancyPct=-95.0; PageSize=1000 }}
        $ex = [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='Credit Card'; Workload='Exchange'; Location=''; LocationType='WorkloadFallback'; OriginalExpectedCount=500; ActualCount=480; DiscrepancyPct=-4.0; PageSize=1000 }}
        Write-RetryTasksCsv -Path '{csv_path}' -RetryTasks @($sp, $ex)
        $rows = @(Import-Csv -Path '{csv_path}' -Encoding UTF8)
        Write-Output ('COUNT=' + $rows.Count)
        Write-Output ('HEADER=' + (($rows[0].PSObject.Properties.Name) -join ','))
        Write-Output ('SP_LOC=' + $rows[0].Location)
        Write-Output ('SP_TYPE=' + $rows[0].LocationType)
        Write-Output ('EX_LOC=[' + $rows[1].Location + ']')
        Write-Output ('EX_TYPE=' + $rows[1].LocationType)
        """
    )
    output = run_pwsh(script)
    assert "COUNT=2" in output, output
    assert "Location" in output and "LocationType" in output
    assert "SP_LOC=https://contoso.sharepoint.com/sites/a" in output
    assert "SP_TYPE=SiteUrl" in output
    assert "EX_LOC=[]" in output  # empty location for fallback row
    assert "EX_TYPE=WorkloadFallback" in output


def test_trainable_classifier_cache_round_trip(tmp_path: Path) -> None:
    """Get-TrainableClassifiersFromCache reads the JSON the helper produces."""
    cache_path = (tmp_path / "CurrentTenantTCs.local.json").as_posix()
    cache_payload = {
        "SchemaVersion": 1,
        "DiscoveredAt": "2026-05-14T07:00:00Z",
        "TenantId": "tenant-123",
        "Source": "purview-portal",
        "ClassifierCount": 2,
        "Classifiers": [
            {
                "Id": "8aef6743-61aa-44b9-9ae5-3bb3d77df535",
                "Name": "Source code",
                "DisplayName": "Source code",
                "Type": "GlobalOOB",
                "ModelStatus": "Stable",
                "IsDeprecated": False,
            },
            {
                "Id": "a02ddb8e-3c93-44ac-87c1-2f682b1cb78e",
                "Name": "Targeted Harassment",
                "DisplayName": "Targeted Harassment",
                "Type": "GlobalOOB",
                "ModelStatus": "Stable",
                "IsDeprecated": False,
            },
        ],
    }
    Path(cache_path).write_text(json.dumps(cache_payload), encoding="utf-8")
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $tcs = @(Get-TrainableClassifiersFromCache -ConfigPath '{cache_path}')
        Write-Output ('COUNT=' + $tcs.Count)
        Write-Output ('NAME0=' + $tcs[0].Name)
        Write-Output ('NAME1=' + $tcs[1].Name)
        Write-Output ('TYPE0=' + $tcs[0].Type)
        """
    )
    output = run_pwsh(script)
    assert "COUNT=2" in output, output
    assert "NAME0=Source code" in output, output
    assert "NAME1=Targeted Harassment" in output, output
    assert "TYPE0=GlobalOOB" in output, output


def test_trainable_classifier_cache_missing_returns_empty(tmp_path: Path) -> None:
    missing_path = (tmp_path / "does-not-exist.json").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $tcs = @(Get-TrainableClassifiersFromCache -ConfigPath '{missing_path}')
        Write-Output ('COUNT=' + $tcs.Count)
        """
    )
    output = run_pwsh(script)
    assert "COUNT=0" in output, output


def test_round_robin_dispatch_prioritizes_exchange_and_teams() -> None:
    """First N tasks must cover Exchange and Teams before piling on SharePoint."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $sp1 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteA'; LocationType='SiteUrl'; ExpectedCount=9000 }}
        $sp2 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteB'; LocationType='SiteUrl'; ExpectedCount=8000 }}
        $sp3 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteC'; LocationType='SiteUrl'; ExpectedCount=7000 }}
        $sp4 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteD'; LocationType='SiteUrl'; ExpectedCount=6000 }}
        $sp5 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteE'; LocationType='SiteUrl'; ExpectedCount=5000 }}
        $sp6 = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='SharePoint'; Location='siteF'; LocationType='SiteUrl'; ExpectedCount=4000 }}
        $ex  = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='Exchange';   Location='';     LocationType='WorkloadFallback'; ExpectedCount=500 }}
        $tm  = @{{ TagType='SensitiveInformationType'; TagName='CC'; Workload='Teams';      Location='';     LocationType='WorkloadFallback'; ExpectedCount=200 }}
        $tasks = @($sp1,$sp2,$sp3,$sp4,$sp5,$sp6,$ex,$tm)
        $ordered = @(Get-RoundRobinDetailTaskOrder -Tasks $tasks)
        Write-Output ('FIRST=' + $ordered[0].Workload)
        Write-Output ('SECOND=' + $ordered[1].Workload)
        Write-Output ('COUNT=' + $ordered.Count)
        $firstFour = @($ordered[0..3] | ForEach-Object {{ $_.Workload }})
        Write-Output ('EX_IN_4=' + ($firstFour -contains 'Exchange'))
        Write-Output ('TM_IN_4=' + ($firstFour -contains 'Teams'))
        """
    )
    output = run_pwsh(script)
    assert "COUNT=8" in output, output
    assert "FIRST=Exchange" in output, output
    assert "SECOND=Teams" in output, output
    assert "EX_IN_4=True" in output, output
    assert "TM_IN_4=True" in output, output


def test_worker_park_unpark_round_trip(tmp_path: Path) -> None:
    """Set-WorkerParked / Test-WorkerParked roundtrip via the parked marker file."""
    worker_dir = (tmp_path / "Worker-12345").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        New-Item -ItemType Directory -Force -Path '{worker_dir}' | Out-Null
        Write-Output ('PARKED_INITIAL=' + (Test-WorkerParked -WorkerDir '{worker_dir}'))
        Set-WorkerParked -WorkerDir '{worker_dir}' -Parked $true
        Write-Output ('PARKED_AFTER_PARK=' + (Test-WorkerParked -WorkerDir '{worker_dir}'))
        Set-WorkerParked -WorkerDir '{worker_dir}' -Parked $false
        Write-Output ('PARKED_AFTER_UNPARK=' + (Test-WorkerParked -WorkerDir '{worker_dir}'))
        """
    )
    output = run_pwsh(script)
    assert "PARKED_INITIAL=False" in output, output
    assert "PARKED_AFTER_PARK=True" in output, output
    assert "PARKED_AFTER_UNPARK=False" in output, output


def test_watermark_save_and_aggregate_delta(tmp_path: Path) -> None:
    """Watermarks round-trip and the aggregate-delta report classifies correctly."""
    script_root = (tmp_path / "fakeroot").as_posix()
    export_dir = (tmp_path / "Export-test").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        New-Item -ItemType Directory -Force -Path '{script_root}\\ConfigFiles' | Out-Null
        New-Item -ItemType Directory -Force -Path '{export_dir}\\_Coordination' | Out-Null

        $detailTasks = @(
            [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='CreditCard'; Workload='SharePoint'; Location=''; Status='Completed'; ExpectedCount=1000; OriginalExpectedCount=1000 }}
        )
        Save-WatermarksFromDetailTasks -ScriptRoot '{script_root}' -TenantPrefix 'zava' -DetailTasks $detailTasks -WasFullRun

        $marks = Read-Watermarks -ScriptRoot '{script_root}' -TenantPrefix 'zava'
        Write-Output ('MARKS_TASK_COUNT=' + $marks.Tasks.Count)
        Write-Output ('MARKS_TENANT=' + $marks.TenantPrefix)
        Write-Output ('MARKS_HAS_FULL=' + ($null -ne $marks.LastFullRunAt))

        $currentAggregates = @(
            [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='CreditCard'; Workload='SharePoint'; Location=''; ExpectedCount=1000 }},
            [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='NewSIT'; Workload='SharePoint'; Location=''; ExpectedCount=42 }},
            [PSCustomObject]@{{ TagType='SensitiveInformationType'; TagName='CreditCard'; Workload='OneDrive'; Location=''; ExpectedCount=2500 }}
        )
        Write-AggregateDeltaReport -ExportDir '{export_dir}' -Watermarks $marks -AggregateTasks $currentAggregates

        $report = Get-Content -Raw -Path '{export_dir}\\_Coordination\\AggregateDelta.json' | ConvertFrom-Json
        Write-Output ('UNCHANGED=' + $report.summary.unchanged)
        Write-Output ('NEW=' + $report.summary.new)
        """
    )
    output = run_pwsh(script)
    assert "MARKS_TASK_COUNT=1" in output, output
    assert "MARKS_TENANT=zava" in output, output
    assert "MARKS_HAS_FULL=True" in output, output
    assert "UNCHANGED=1" in output, output
    assert "NEW=2" in output, output


def test_unknown_workload_collapses_to_workload_fallback() -> None:
    """Get-ContentExplorerLocationType default + New-ContentExplorerDetailTasks collapse."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $type = Get-ContentExplorerLocationType -Workload 'Copilot'
        Write-Output ('UNKNOWN_TYPE=' + $type)
        $workPlan = ,@{{ TagType='SensitiveInformationType'; TagName='Credit Card'; Workload='Copilot'; TotalCount=30; Locations=@(@{{Name='loc1'; ExpectedCount=10}}, @{{Name='loc2'; ExpectedCount=20}}) }}
        $detail = @(New-ContentExplorerDetailTasks -WorkPlanTasks $workPlan -DefaultPageSize 1000)
        Write-Output ('DETAIL_COUNT=' + $detail.Count)
        Write-Output ('DETAIL_TYPE=' + $detail[0].LocationType)
        Write-Output ('DETAIL_LOC_EMPTY=' + [string]::IsNullOrEmpty($detail[0].Location))
        """
    )
    output = run_pwsh(script)
    assert "UNKNOWN_TYPE=WorkloadFallback" in output, output
    assert "DETAIL_COUNT=1" in output, output  # location tasks collapsed into one fallback task
    assert "DETAIL_TYPE=WorkloadFallback" in output
    assert "DETAIL_LOC_EMPTY=True" in output


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
