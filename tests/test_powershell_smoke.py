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


# ── B1: RunSummary.json + deterministic exit codes ────────────────────────────


def test_run_result_functions_are_exported() -> None:
    """Write-RunSummary and Get-ExportExitCode are visible after Import-Module."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        $wrs = Get-Command Write-RunSummary -ErrorAction SilentlyContinue
        $gee = Get-Command Get-ExportExitCode -ErrorAction SilentlyContinue
        Write-Output ('WRS=' + ($null -ne $wrs))
        Write-Output ('GEE=' + ($null -ne $gee))
        """
    )
    output = run_pwsh(script)
    assert "WRS=True" in output, output
    assert "GEE=True" in output, output


def test_get_export_exit_code_map() -> None:
    """Get-ExportExitCode returns the correct integer for every named status."""
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        Write-Output ('COMPLETED='   + (Get-ExportExitCode -Status 'Completed'))
        Write-Output ('FAILED='      + (Get-ExportExitCode -Status 'Failed'))
        Write-Output ('PARTIAL='     + (Get-ExportExitCode -Status 'Partial'))
        Write-Output ('AUTHFAILED='  + (Get-ExportExitCode -Status 'AuthFailed'))
        Write-Output ('CONFIGERROR=' + (Get-ExportExitCode -Status 'ConfigError'))
        Write-Output ('LOCKED='      + (Get-ExportExitCode -Status 'Locked'))
        """
    )
    output = run_pwsh(script)
    assert "COMPLETED=0" in output, output
    assert "FAILED=1" in output, output
    assert "PARTIAL=2" in output, output
    assert "AUTHFAILED=3" in output, output
    assert "CONFIGERROR=4" in output, output
    assert "LOCKED=5" in output, output


def test_write_run_summary_json_shape(tmp_path: Path) -> None:
    """Write-RunSummary writes a valid RunSummary.json with the required keys/values."""
    export_dir = tmp_path / "Export-20260615-120000"
    export_dir.mkdir()
    # Write a minimal section spec to a helper JSON so the PS script can load it
    # without multi-line hashtable literals (multi-line blocks are silent via stdin pipe).
    section_json = export_dir / "_test_section.json"
    section_json.write_text(
        json.dumps([{"Name": "SensitiveInformationType", "Status": "Completed", "RecordCount": 42, "ErrorCount": 0}]),
        encoding="utf-8",
    )
    export_dir_posix = export_dir.as_posix()
    section_path = section_json.as_posix()
    summary_path = (export_dir / "RunSummary.json").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        Initialize-ExportLog -LogDirectory '{export_dir_posix}' -Prefix 'Test' | Out-Null
        $secs = @(Get-Content -Raw '{section_path}' | ConvertFrom-Json)
        Write-RunSummary -ExportDir '{export_dir_posix}' -Result @{{ Mode='ContentExplorer'; Status='Completed'; StartedUtc=([datetime]'2026-06-15T12:00:00Z'); Sections=$secs; RemainingTasks=0; Errors=@() }}
        Write-Output ('WRITTEN=' + (Test-Path '{summary_path}'))
        """
    )
    output = run_pwsh(script)
    assert "WRITTEN=True" in output, output

    summary = json.loads((export_dir / "RunSummary.json").read_text(encoding="utf-8"))
    assert summary["schemaVersion"] == 1
    assert summary["mode"] == "ContentExplorer"
    assert summary["status"] == "Completed"
    assert summary["exitCode"] == 0
    assert summary["remainingTasks"] == 0
    assert isinstance(summary["sections"], list)
    assert len(summary["sections"]) == 1
    assert summary["sections"][0]["name"] == "SensitiveInformationType"
    assert summary["sections"][0]["recordCount"] == 42
    assert isinstance(summary["errors"], list)
    assert isinstance(summary["droppedErrors"], int)
    assert "startedUtc" in summary
    assert "endedUtc" in summary


def test_write_run_summary_partial_status(tmp_path: Path) -> None:
    """Write-RunSummary with Partial status emits exitCode 2."""
    export_dir = tmp_path / "Export-partial"
    export_dir.mkdir()
    export_dir_posix = export_dir.as_posix()
    summary_path = (export_dir / "RunSummary.json").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        Initialize-ExportLog -LogDirectory '{export_dir_posix}' -Prefix 'Test' | Out-Null
        Write-RunSummary -ExportDir '{export_dir_posix}' -Result @{{ Mode='ActivityExplorer'; Status='Partial'; RemainingTasks=3; Errors=@() }}
        Write-Output ('EXITCODE=' + (Get-ExportExitCode -Status 'Partial'))
        Write-Output ('WRITTEN=' + (Test-Path '{summary_path}'))
        """
    )
    output = run_pwsh(script)
    assert "EXITCODE=2" in output, output
    assert "WRITTEN=True" in output, output

    summary = json.loads((export_dir / "RunSummary.json").read_text(encoding="utf-8"))
    assert summary["status"] == "Partial"
    assert summary["exitCode"] == 2
    assert summary["remainingTasks"] == 3


def test_write_run_summary_failed_status(tmp_path: Path) -> None:
    """Write-RunSummary with Failed status emits exitCode 1 (the production-crash path)."""
    export_dir = tmp_path / "Export-failed"
    export_dir.mkdir()
    export_dir_posix = export_dir.as_posix()
    summary_path = (export_dir / "RunSummary.json").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        Initialize-ExportLog -LogDirectory '{export_dir_posix}' -Prefix 'Test' | Out-Null
        Write-RunSummary -ExportDir '{export_dir_posix}' -Result @{{ Mode='ContentExplorer'; Status='Failed'; RemainingTasks=0; Errors=@() }}
        Write-Output ('EXITCODE=' + (Get-ExportExitCode -Status 'Failed'))
        Write-Output ('WRITTEN=' + (Test-Path '{summary_path}'))
        """
    )
    output = run_pwsh(script)
    assert "EXITCODE=1" in output, output
    assert "WRITTEN=True" in output, output

    summary = json.loads((export_dir / "RunSummary.json").read_text(encoding="utf-8"))
    assert summary["status"] == "Failed"
    assert summary["exitCode"] == 1


def test_write_run_summary_errors_capped_to_20(tmp_path: Path) -> None:
    """Write-RunSummary caps the errors array at 20 and records droppedErrors count."""
    export_dir = tmp_path / "Export-caperrors"
    export_dir.mkdir()
    # Build 25 error entries in Python and hand them to PS via a JSON file — far
    # cleaner than emitting 25 inline hashtable literals into the here-string.
    # (Inline single-line hashtables work fine over the stdin pipe; only multi-line
    # for/foreach blocks are silently swallowed by `pwsh -Command -`.)
    errors_data = [{"Timestamp": f"2026-06-15T12:00:{i:02d}Z", "Message": f"error {i}"} for i in range(25)]
    errors_json = export_dir / "_test_errors.json"
    errors_json.write_text(json.dumps(errors_data), encoding="utf-8")
    export_dir_posix = export_dir.as_posix()
    errors_path = errors_json.as_posix()
    summary_path = (export_dir / "RunSummary.json").as_posix()
    script = textwrap.dedent(
        f"""
        Import-Module '{MODULE_PATH}' -Force
        Initialize-ExportLog -LogDirectory '{export_dir_posix}' -Prefix 'Test' | Out-Null
        $errs = @(Get-Content -Raw '{errors_path}' | ConvertFrom-Json)
        Write-RunSummary -ExportDir '{export_dir_posix}' -Result @{{ Mode='Full'; Status='Partial'; Errors=$errs }}
        Write-Output ('WRITTEN=' + (Test-Path '{summary_path}'))
        """
    )
    output = run_pwsh(script)
    assert "WRITTEN=True" in output, output

    summary = json.loads((export_dir / "RunSummary.json").read_text(encoding="utf-8"))
    assert len(summary["errors"]) == 20
    assert summary["droppedErrors"] == 5


# ── B2: -Unattended switch — prompt-gate source assertions ────────────────────


def test_unattended_gate_prompt_a_proceed_confirm() -> None:
    """Prompt A (Proceed?) in MainExecution.ps1 must be wrapped by $script:Unattended guard."""
    source = (SCRIPT_PARTS_ROOT / "MainExecution.ps1").read_text(encoding="utf-8")
    # The Read-Host must be inside an -not $script:Unattended block
    assert "if (-not $script:Unattended)" in source, "Unattended guard missing in MainExecution.ps1"
    assert 'Read-Host "Proceed with export? [Y]/N"' in source, "Prompt A text changed or missing"
    # The else branch must log the skip
    assert "prompt A skipped" in source, "Unattended else-branch log for prompt A missing"


def test_unattended_gate_prompt_b_aggregate_reuse() -> None:
    """Prompt B (aggregate reuse) in ContentExplorer.Export.ps1 must be guarded."""
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.Export.ps1").read_text(encoding="utf-8")
    assert "if (-not $script:Unattended)" in source, "Unattended guard missing in ContentExplorer.Export.ps1"
    assert 'Read-Host "Enter choice [N]"' in source, "Prompt B text changed or missing"
    assert "prompt B skipped" in source, "Unattended else-branch log for prompt B missing"
    # Unattended default must be N (generate fresh)
    assert '$choice = "N"' in source, "Unattended default for prompt B must be N"


def test_unattended_gate_prompt_c_resume_confirm() -> None:
    """Prompt C (resume confirm) in ContentExplorer.Resume.ps1 must be guarded."""
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.Resume.ps1").read_text(encoding="utf-8")
    assert "if (-not $script:Unattended)" in source, "Unattended guard missing in ContentExplorer.Resume.ps1"
    assert 'Read-Host "  Resume this export? [Y/n]"' in source, "Prompt C text changed or missing"
    assert "prompt C skipped" in source, "Unattended else-branch log for prompt C missing"


def test_unattended_gate_prompt_d_retry_confirm() -> None:
    """Prompt D (retry confirm) in ContentExplorer.Retry.ps1 must be guarded."""
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.Retry.ps1").read_text(encoding="utf-8")
    assert "if (-not $script:Unattended)" in source, "Unattended guard missing in ContentExplorer.Retry.ps1"
    assert 'Read-Host "  Retry these tasks? [Y/N]"' in source, "Prompt D text changed or missing"
    assert "prompt D skipped" in source, "Unattended else-branch log for prompt D missing"


def test_unattended_gate_prompt_e_tasks_csv_confirm() -> None:
    """Prompt E (tasks-CSV confirm) in ContentExplorer.TasksCsv.ps1 must be guarded."""
    source = (SCRIPT_PARTS_ROOT / "Orchestrator" / "ContentExplorer.TasksCsv.ps1").read_text(encoding="utf-8")
    assert "if (-not $script:Unattended)" in source, "Unattended guard missing in ContentExplorer.TasksCsv.ps1"
    assert 'Read-Host "  Run these tasks? [Y/N]"' in source, "Prompt E text changed or missing"
    assert "prompt E skipped" in source, "Unattended else-branch log for prompt E missing"


def test_unattended_gate_confirm_connected_tenant() -> None:
    """Confirm-ConnectedTenant's non-interactive guard must also honor $script:Unattended,
    so a scheduled run on an interactive console (IsInputRedirected = $false) does not hang."""
    source = (SCRIPT_PARTS_ROOT / "MainExecution.ps1").read_text(encoding="utf-8")
    assert "[Console]::IsInputRedirected -or $script:Unattended" in source, (
        "Confirm-ConnectedTenant guard must OR in $script:Unattended"
    )


def test_unattended_fallback_h_config_error_exit() -> None:
    """Fallback H in MainExecution.ps1: unattended + no-mode path must exit with ConfigError (4)."""
    source = (SCRIPT_PARTS_ROOT / "MainExecution.ps1").read_text(encoding="utf-8")
    # The positive-form guard `if ($script:Unattended)` is unique to fallback H — the five
    # prompt gates all use the negative `if (-not $script:Unattended)` — so it uniquely
    # pins the fallback-H block rather than matching any prompt gate.
    assert "if ($script:Unattended)" in source, "positive-form fallback-H guard missing in MainExecution.ps1"
    assert "Get-ExportExitCode -Status 'ConfigError'" in source, "ConfigError exit missing from MainExecution.ps1"
    assert "Write-RunSummary" in source, "Write-RunSummary not called in MainExecution.ps1 fallback H"


def test_unattended_switch_declared_in_entry_script() -> None:
    """Export-Compl8Configuration.ps1 must declare [switch]$Unattended and propagate it."""
    entry = (SCRIPT_PARTS_ROOT.parent / "Export-Compl8Configuration.ps1").read_text(encoding="utf-8")
    assert "[switch]$Unattended" in entry, "[switch]$Unattended not declared in entry script"
    assert "$script:Unattended = [bool]$Unattended" in entry, "Unattended not propagated to script scope"
