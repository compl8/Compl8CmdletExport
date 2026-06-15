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


def test_min_location_items_filters_small_locations() -> None:
    """MinLocationItems drops sub-threshold locations but never WorkloadFallback tasks."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $workPlan = @(
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'OneDrive'; TotalCount = 121;
                Locations = @(
                    @{{ Name = 'https://contoso-my.sharepoint.com/personal/u1'; ExpectedCount = 100 }},
                    @{{ Name = 'https://contoso-my.sharepoint.com/personal/u2'; ExpectedCount = 9 }},
                    @{{ Name = 'https://contoso-my.sharepoint.com/personal/u3'; ExpectedCount = 2 }},
                    @{{ Name = 'https://contoso-my.sharepoint.com/personal/u4'; ExpectedCount = 10 }}
                )
            }},
            @{{
                TagType = 'SensitiveInformationType'; TagName = 'Credit Card'; Workload = 'Exchange'; TotalCount = 5;
                Locations = @(@{{ Name = 'u5@contoso.com'; ExpectedCount = 5 }})
            }}
        )

        $filtered = @(New-ContentExplorerDetailTasks -WorkPlanTasks $workPlan -DefaultPageSize 1000 -WorkloadFallbackWorkloads @('Exchange') -MinLocationItems 10)
        $od = @($filtered | Where-Object {{ $_.Workload -eq 'OneDrive' }})
        $ex = @($filtered | Where-Object {{ $_.Workload -eq 'Exchange' }})
        Write-Output ('OD_COUNT=' + $od.Count)
        Write-Output ('OD_MIN=' + (($od | ForEach-Object {{ [int]$_.ExpectedCount }} | Measure-Object -Minimum).Minimum))
        Write-Output ('EX_FALLBACK_KEPT=' + (($ex.Count -eq 1) -and $ex[0].LocationType -eq 'WorkloadFallback'))

        $unfiltered = @(New-ContentExplorerDetailTasks -WorkPlanTasks $workPlan -DefaultPageSize 1000 -WorkloadFallbackWorkloads @('Exchange'))
        Write-Output ('DEFAULT_KEEPS_ALL=' + (@($unfiltered | Where-Object {{ $_.Workload -eq 'OneDrive' }}).Count -eq 4))
        """
    )

    output = run_pwsh(script)
    assert "OD_COUNT=2" in output, f"expected 2 OneDrive tasks (>=10 items), got: {output}"
    assert "OD_MIN=10" in output
    assert "EX_FALLBACK_KEPT=True" in output
    assert "DEFAULT_KEEPS_ALL=True" in output


def test_select_largest_pending_task_prioritizes_aggregates_then_largest() -> None:
    """Dispatch-order callback: pending aggregates win; otherwise largest pending ExpectedCount."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $tasks = [System.Collections.ArrayList]@(
            @{{ Phase = 'Detail'; Status = 'Completed'; ExpectedCount = 99999 }},
            @{{ Phase = 'Detail'; Status = 'Pending'; ExpectedCount = 50 }},
            @{{ Phase = 'Detail'; Status = 'Pending'; ExpectedCount = '12000' }},
            @{{ Phase = 'Detail'; Status = 'Pending'; ExpectedCount = 5000 }}
        )
        $pick = Select-LargestPendingTask -Tasks $tasks
        Write-Output ('LARGEST=' + $pick.ExpectedCount)

        [void]$tasks.Add(@{{ Phase = 'Aggregate'; Status = 'Pending'; ExpectedCount = 1 }})
        $pick2 = Select-LargestPendingTask -Tasks $tasks
        Write-Output ('AGG_FIRST=' + ($pick2.Phase -eq 'Aggregate'))

        $none = Select-LargestPendingTask -Tasks @(@{{ Phase = 'Detail'; Status = 'Completed'; ExpectedCount = 1 }})
        Write-Output ('NONE=' + ($null -eq $none))
        """
    )

    output = run_pwsh(script)
    assert "LARGEST=12000" in output, f"expected string-typed 12000 to win, got: {output}"
    assert "AGG_FIRST=True" in output
    assert "NONE=True" in output


def test_content_explorer_export_assigns_large_all_sit_settings() -> None:
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.Export.ps1").read_text(encoding="utf-8")
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
    """Get-TrainableClassifiersFromCache reads the externally-provided TC cache JSON.

    The cache file is produced by the external GetTCs tool (distributed
    separately) and dropped at ConfigFiles/CurrentTenantTCs.local.json; this
    fixture mirrors that contract."""
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


_RULEPACK_GUID_ENTITY = "11111111-1111-1111-1111-111111111111"
_RULEPACK_GUID_AFFINITY = "22222222-2222-2222-2222-222222222222"


def _write_synthetic_rulepack(path: Path) -> Path:
    """UTF-16 RulePackage XML fixture (synthetic; no tenant data)."""
    xml = f"""<?xml version="1.0" encoding="utf-16"?>
<RulePackage xmlns="http://schemas.microsoft.com/office/2011/mce">
  <RulePack id="44444444-4444-4444-4444-444444444444">
    <Version build="0" major="1" minor="0" revision="0"/>
    <Publisher id="33333333-3333-3333-3333-333333333333"/>
    <Details defaultLangCode="en-us">
      <LocalizedDetails langcode="en-us">
        <PublisherName>Synthetic</PublisherName>
        <Name>Synthetic Rule Pack</Name>
        <Description>Test fixture.</Description>
      </LocalizedDetails>
    </Details>
  </RulePack>
  <Rules>
    <Entity id="{_RULEPACK_GUID_ENTITY}" patternsProximity="300" recommendedConfidence="75">
      <Pattern confidenceLevel="75"><IdMatch idRef="Regex_a"/></Pattern>
    </Entity>
    <Affinity id="{_RULEPACK_GUID_AFFINITY}" evidencesProximity="300" thresholdConfidenceLevel="75">
      <Evidence confidenceLevel="75"><Match idRef="Regex_a"/></Evidence>
    </Affinity>
    <Regex id="Regex_a">(\\d{{9}})</Regex>
    <LocalizedStrings>
      <Resource idRef="{_RULEPACK_GUID_ENTITY}">
        <Name default="false" langcode="de-de">Mitarbeiter-ID</Name>
        <Name default="true" langcode="en-us">Employee ID</Name>
      </Resource>
      <Resource idRef="{_RULEPACK_GUID_AFFINITY}">
        <Name langcode="en-us">All Credentials Bundle</Name>
      </Resource>
      <Resource idRef="not-a-guid">
        <Name default="true" langcode="en-us">Ignored Non-Guid</Name>
      </Resource>
    </LocalizedStrings>
  </Rules>
</RulePackage>
"""
    path.write_bytes(xml.encode("utf-16"))
    return path


def test_get_sit_names_from_rulepack_xml_parses_localized_strings(tmp_path: Path) -> None:
    """Entity + Affinity GUIDs resolve; default-language name preferred; non-GUID ids dropped."""
    xml_path = _write_synthetic_rulepack(tmp_path / "pack.xml").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $names = Get-SitNamesFromRulePackXml -Path '{xml_path}'
        Write-Output ('COUNT=' + $names.Count)
        Write-Output ('ENTITY=' + $names['{_RULEPACK_GUID_ENTITY}'])
        Write-Output ('AFFINITY=' + $names['{_RULEPACK_GUID_AFFINITY}'])
        """
    )
    output = run_pwsh(script)
    assert "COUNT=2" in output, output
    assert "ENTITY=Employee ID" in output, output  # default="true" beats first (de-de) name
    assert "AFFINITY=All Credentials Bundle" in output, output


def test_export_sit_reference_snapshot_merges_flat_list_and_rule_packs(tmp_path: Path) -> None:
    """Flat-list names win, rule-pack XML fills the gaps; raw pack XML is saved;
    a second run without -Force skips (idempotent across CE+AE in a full export)."""
    xml_path = _write_synthetic_rulepack(tmp_path / "pack.xml").as_posix()
    export_dir = tmp_path / "Export-20260612-000000"
    export_dir.mkdir()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        function Get-DlpSensitiveInformationType {{
            [pscustomobject]@{{ Id = [guid]'{_RULEPACK_GUID_ENTITY}'; Name = 'Flat Alpha'; Publisher = 'Microsoft Corporation' }}
        }}

        function Get-DlpSensitiveInformationTypeRulePackage {{
            [pscustomobject]@{{
                RuleCollectionName = 'Synthetic Rule Package'
                SerializedClassificationRuleCollection = [System.IO.File]::ReadAllBytes('{xml_path}')
            }}
        }}

        $result = Export-SitReferenceSnapshot -ExportRunDirectory '{export_dir.as_posix()}'
        Write-Output ('TOTAL=' + $result.TotalNames)
        Write-Output ('PACKS=' + $result.RulePackCount)
        Write-Output ('PACKNAMES=' + $result.RulePackNameCount)
        $second = Export-SitReferenceSnapshot -ExportRunDirectory '{export_dir.as_posix()}'
        Write-Output ('SECOND_SKIPPED=' + ($null -eq $second))
        """
    )
    output = run_pwsh(script)
    assert "TOTAL=2" in output, output
    assert "PACKS=1" in output, output
    assert "PACKNAMES=2" in output, output
    assert "SECOND_SKIPPED=True" in output, output

    snapshot = json.loads((export_dir / "CurrentTenantSITs.json").read_text(encoding="utf-8-sig"))
    assert snapshot[_RULEPACK_GUID_ENTITY] == "Flat Alpha"  # flat-list name wins
    assert snapshot[_RULEPACK_GUID_AFFINITY] == "All Credentials Bundle"  # pack fills gap
    saved_packs = list((export_dir / "Data" / "Reference" / "RulePackages").glob("*.xml"))
    assert len(saved_packs) == 1, saved_packs


def test_export_sit_reference_snapshot_skips_without_cmdlets(tmp_path: Path) -> None:
    """No S&C session (cmdlet absent): warn + return $null, write nothing."""
    export_dir = tmp_path / "Export-20260612-000001"
    export_dir.mkdir()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $result = Export-SitReferenceSnapshot -ExportRunDirectory '{export_dir.as_posix()}'
        Write-Output ('NULL=' + ($null -eq $result))
        Write-Output ('FILE=' + (Test-Path '{(export_dir / "CurrentTenantSITs.json").as_posix()}'))
        """
    )
    output = run_pwsh(script)
    assert "NULL=True" in output, output
    assert "FILE=False" in output, output


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


def _parse_kv(output: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in output.splitlines():
        line = line.strip()
        if "=" in line:
            key, value = line.split("=", 1)
            result[key] = value
    return result


def test_progress_eta_tracks_recent_rate_after_speedup() -> None:
    """A long slow ramp then a sustained fast burst: the windowed ETA must reflect
    the recent (fast) rate, well above the lifetime cumulative average."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $state = @{{}}
        $base = Get-Date '2026-01-01T00:00:00Z'
        $completed = 0.0
        $r = $null
        # Slow: +1 unit every 5s for 600s (12 units/min). One-line loop: pwsh
        # -Command - over stdin silently drops output for multi-line loop blocks.
        for ($i = 1; $i -le 120; $i++) {{ $completed += 1; $r = Get-ProgressEta -State $state -Now $base.AddSeconds(5 * $i) -CompletedUnits $completed -RemainingUnits 1000 }}
        Write-Output ('SLOW=' + [math]::Round($r.RatePerMinute, 1))
        # Fast: +20 units every 5s for 240s (240 units/min); long enough to flush the 120s window
        for ($i = 1; $i -le 48; $i++) {{ $completed += 20; $r = Get-ProgressEta -State $state -Now $base.AddSeconds(600 + 5 * $i) -CompletedUnits $completed -RemainingUnits 1000 }}
        $elapsedMin = (600 + 240) / 60.0
        Write-Output ('FAST=' + [math]::Round($r.RatePerMinute, 1))
        Write-Output ('CUMULATIVE=' + [math]::Round($completed / $elapsedMin, 1))
        Write-Output ('SOURCE=' + $r.Source)
        Write-Output ('ETASEC=' + [math]::Round($r.EtaSeconds, 0))
        """
    )
    vals = _parse_kv(run_pwsh(script))
    assert vals["SOURCE"] == "window", vals
    slow = float(vals["SLOW"])
    fast = float(vals["FAST"])
    cumulative = float(vals["CUMULATIVE"])
    etasec = float(vals["ETASEC"])
    assert slow < 30, vals                       # recent rate during the slow ramp
    assert 220 <= fast <= 245, vals              # converged to the ~240/min recent burst
    assert cumulative < 110, vals                # lifetime average still dragged by the slow start
    assert fast > 2 * cumulative, vals           # recent rate dominates the average
    assert 220 <= etasec <= 300, vals            # ETA ~250s from recent rate, not ~780s cumulative


def test_progress_eta_reacts_to_slowdown() -> None:
    """A fast burst then a sustained slowdown: the windowed ETA must drop toward the
    recent (slow) rate, far below the lifetime average."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $state = @{{}}
        $base = Get-Date '2026-01-01T00:00:00Z'
        $completed = 0.0
        $r = $null
        for ($i = 1; $i -le 48; $i++) {{ $completed += 20; $r = Get-ProgressEta -State $state -Now $base.AddSeconds(5 * $i) -CompletedUnits $completed -RemainingUnits 5000 }}
        Write-Output ('PEAK=' + [math]::Round($r.RatePerMinute, 1))
        for ($i = 1; $i -le 48; $i++) {{ $completed += 1; $r = Get-ProgressEta -State $state -Now $base.AddSeconds(240 + 5 * $i) -CompletedUnits $completed -RemainingUnits 5000 }}
        $elapsedMin = (240 + 240) / 60.0
        Write-Output ('NOWRATE=' + [math]::Round($r.RatePerMinute, 1))
        Write-Output ('CUMULATIVE=' + [math]::Round($completed / $elapsedMin, 1))
        Write-Output ('SOURCE=' + $r.Source)
        """
    )
    vals = _parse_kv(run_pwsh(script))
    assert vals["SOURCE"] == "window", vals
    peak = float(vals["PEAK"])
    nowrate = float(vals["NOWRATE"])
    cumulative = float(vals["CUMULATIVE"])
    assert 220 <= peak <= 245, vals              # the earlier fast burst
    assert nowrate < 30, vals                    # converged down to the ~12/min recent rate
    assert nowrate < peak / 5, vals              # reacted to the slowdown
    assert nowrate < 0.5 * cumulative, vals      # recent rate far below the lifetime average


def test_progress_eta_warmup_then_window() -> None:
    """First frame is not ready; a sub-MinSpan frame is cumulative; once enough
    recent span accrues it switches to the windowed source."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $state = @{{}}
        $base = Get-Date '2026-01-01T00:00:00Z'
        $r1 = Get-ProgressEta -State $state -Now $base -CompletedUnits 0 -RemainingUnits 100
        $r2 = Get-ProgressEta -State $state -Now $base.AddSeconds(5) -CompletedUnits 3 -RemainingUnits 100
        $r3 = Get-ProgressEta -State $state -Now $base.AddSeconds(25) -CompletedUnits 12 -RemainingUnits 100
        Write-Output ('R1_READY=' + $r1.Ready)
        Write-Output ('R1_SOURCE=' + $r1.Source)
        Write-Output ('R2_READY=' + $r2.Ready)
        Write-Output ('R2_SOURCE=' + $r2.Source)
        Write-Output ('R3_SOURCE=' + $r3.Source)
        """
    )
    output = run_pwsh(script)
    assert "R1_READY=False" in output, output
    assert "R1_SOURCE=none" in output, output
    assert "R2_READY=True" in output, output
    assert "R2_SOURCE=cumulative" in output, output
    assert "R3_SOURCE=window" in output, output


def test_progress_eta_zero_remaining_is_done() -> None:
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $state = @{{}}
        $base = Get-Date '2026-01-01T00:00:00Z'
        $r = Get-ProgressEta -State $state -Now $base -CompletedUnits 100 -RemainingUnits 0
        Write-Output ('READY=' + $r.Ready)
        Write-Output ('ETASEC=' + $r.EtaSeconds)
        """
    )
    output = run_pwsh(script)
    assert "READY=True" in output, output
    assert "ETASEC=0" in output, output


def test_dashboard_eta_uses_windowed_estimator() -> None:
    """The dashboards must route ETA through Get-ProgressEta; the old
    cumulative-average ETA expressions must be gone."""
    source = (MODULE_PARTS_ROOT / "UI" / "02-Dashboards.ps1").read_text(encoding="utf-8")
    assert source.count("Get-ProgressEta -State") >= 3  # CE aggregate, CE detail, AE
    assert "s/task avg" not in source                   # old aggregate cumulative label
    assert "$pctPerSecond" not in source                # old AE cumulative-percent rate
    assert "$blendThreshold" not in source              # old detail seed/measured blend


def test_aggregate_dashboard_shows_per_phase_progress() -> None:
    """During the aggregate phase the orchestrator dashboard must report
    aggregate-task completion (aggDone/aggTotal, a stable denominator), not the
    conflated agg+detail total that balloons as detail tasks are generated."""
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.Export.ps1").read_text(encoding="utf-8")
    assert 'if ($displayPhase -eq "Aggregate") {' in source
    assert "$displayCompleted = $aggDone\n" in source
    assert "$displayTotal = $aggTotal\n" in source
    assert "$displayCompleted = $detDone" in source
    assert "$displayTotal = $detTotal" in source
    # The old conflated agg+detail total (made the aggregate % crawl) must be gone.
    assert "$displayCompleted = $aggDone + $detDone" not in source
    assert "$displayTotal = $aggTotal + $detTotal" not in source
