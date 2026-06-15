# Live Validation + Phase 1 (Unattended Foundation) — Plan

> **For agentic workers:** Part B (Phase 1 code) is executed via superpowers:subagent-driven-development, one commit per task, parse/import + smoke-assert verified. Part A is an **operational runbook** (owner-in-the-loop live runs), not TDD. Steps use `- [ ]` for tracking.

**Goal:** Prove the merged Phase-0 reliability/refactor work against a real tenant (Aairii first — small/fast — then a larger tenant), then build the unattended-run foundation (Phase 1) on top of validated code.

**Why this order:** Everything on `main` (22 commits) is parse/import/test-verified but has **never run against a live tenant**. A small interactive run on Aairii is the first real exercise of the orchestration changes (esp. H1, the CE refactors R1/R2/R5/R6, M5/M7) and the auth/connect path. Phase 1 (unattended/cert) is for *future scheduled* runs and is not needed for these interactive runs, so it follows validation.

**Tech stack:** PowerShell 7.4+, ExchangeOnlineManagement ≥ 3.2.0 (Connect-IPPSSession), Python analytics (`parquet_builder`, pytest), pbi-tools.core.

**Verification reality:** the PS code has no unit-test runtime harness — verification = `Parser::ParseFile` parse-check + `Import-Module` + in-process mocks + `tests/test_powershell_smoke.py` source assertions + (ultimately) the live run. Python side has 177 pytest. Plan steps reflect this, not pytest-TDD for PS.

---

## Part 0 — Housekeeping (optional, ~2 min, owner OK needed)

- [ ] **Push `main` to origin** (still pending owner go-ahead): `git -C C:\claudecode\Compl8CmdletExport push origin main`. Not required for the runs.
- [ ] **Delete merged branch:** `git -C C:\claudecode\Compl8CmdletExport branch -d backlog/reliability-and-refactor` (fully merged into main).

---

## Part A — Live validation runbook (Aairii → larger)

Owner-in-the-loop: interactive (WAM/browser) auth — **owner logs in**. No cert/Phase-1 needed here.

### A1 — Pre-flight (before any connect)

- [ ] **Confirm environment.** In repo root: `pwsh -v` ≥ 7.4; `Get-Module ExchangeOnlineManagement -ListAvailable | Select Version` ≥ 3.2.0; `Import-Module .\Modules\Compl8ExportFunctions.psm1 -Force` → OK; `Test-ExportPrerequisites` (module fn) passes.
- [ ] **Resolve run mechanics (KEY).** The interactive Connect-IPPSSession opens a browser; the Claude PowerShell tool runs `-NonInteractive` and cannot complete WAM auth. Two supported modes:
  - **(preferred) Owner-run:** owner runs the exact command below in their own interactive `pwsh` window and completes login; Claude monitors by reading `Output\Export-*\_Logs\*` and `Output\Export-*\Data\*` (Claude can read those files) and analyzes.
  - **(fallback) Claude background-launch:** Claude starts the command via `run_in_background`; if the browser/WAM prompt surfaces for the owner to complete, proceed; if it can't (NonInteractive blocks it), fall back to owner-run.
  Decide which at run time; either way Claude does the output analysis.
- [ ] **Pick Aairii smoke scope.** Aairii is small (it's the export's named feed), so a near-full export is still quick. Validate single-terminal first (lowest risk), then multi-terminal, then scale to a larger tenant. Keep CE scope modest on the first pass (e.g. default `ConfigFiles\ContentExplorerClassifiers.json`, and `-CEMinLocationItems 1` to skip empty locations) so aggregate+detail completes fast.
- [ ] **Snapshot a clean baseline:** note the current `Output\` contents so new `Export-*` dirs are easy to find.

### A2 — Aairii single-terminal Content Explorer (validates H1, R1/R2/R5/R6, M5, M7, L4/L5, SIT snapshot)

This is the **highest-value smoke**: H1 was a crash at the *first detail task of fresh single-terminal CE* — exactly this path — and R1/R2/R5/R6 all restructured CE.

- [ ] **Run (owner logs in):**
  ```powershell
  .\Export-Compl8Configuration.ps1 -ContentExplorer -CEMinLocationItems 1 -UserPrincipalName <aairii-admin-upn>
  ```
  (or menu option `[1]`). Single-terminal, fresh.
- [ ] **Watch for (Claude tails `Output\Export-*\_Logs\`):**
  - Connect succeeds; aggregate phase discovers tags; **detail phase starts without the old `$exportDir` null crash** (H1 confirmed);
  - `Invoke-CEAggregatePaging` paginates aggregates cleanly (R5); detail tasks build via `Build-CEDetailExportParams` (R2); no signal re-warn loops (M5);
  - `_manifest.json` written; per-page JSON files present under `Data\ContentExplorer\...`.
- [ ] **Success criteria:** run reaches `Completed` phase; `Data\ContentExplorer\_manifest.json` exists with non-zero TotalRecords (or a clean zero if the tenant truly has none); no unhandled exception in `_Logs\ExportProject-Errors.log`; `RemainingTasks.csv` empty/absent.
- [ ] **If it fails:** capture the exception + stack from `_Logs\`, identify which committed change (or a pre-existing latent issue) is implicated, fix on a new branch with the same review cadence, re-run. (First-live-run surprises are expected — this is the point of the smoke.)

### A3 — Aairii single-terminal Activity Explorer (validates H2, M4, M5-AE, L3)

- [ ] **Run (owner logs in):**
  ```powershell
  .\Export-Compl8Configuration.ps1 -ActivityExplorer -PastDays 2
  ```
- [ ] **Watch for:** initial-day query survives a transient blip if any (M4); per-day pages written under `Data\ActivityExplorer\YYYY-MM-DD\`; on any partial day the status is recorded as PartialFailure/Failed (H2) — not silently Completed; `_manifest.json` written with DaysExported.
- [ ] **Success criteria:** AE phase `AECompleted`; manifest day count == requested; no false "complete" on a partial day.

### A4 — Aairii multi-terminal CE (validates dispatch loop, M1 worker reclaim, worker spawn/health)

Small/quick on Aairii, but exercises the worker + dispatch surface the single-terminal runs don't.

- [ ] **Run (owner logs in; spawns worker windows):**
  ```powershell
  .\Export-Compl8Configuration.ps1 -ContentExplorer -WorkerCount 2 -CEMinLocationItems 1
  ```
- [ ] **Watch for:** 2 worker windows spawn and auth (interactive → each window logs in, expected); orchestrator dispatches via file-drop; dashboard `TaskTime` column populates (L4); workers complete tasks, emit `detail-done-*` signals; orchestrator reaches `Completed`; no leaked `pwsh` after completion.
- [ ] **Success criteria:** all detail tasks `Completed`; manifest counts match single-terminal run for the same scope (cross-check A2); workers exit cleanly.

### A5 — Verify outputs + analytics conversion + SIT snapshot

- [ ] **SIT snapshot:** confirm `Output\Export-*\CurrentTenantSITs.json` written; then run the converter and check resolution:
  ```powershell
  py -m parquet_builder.star.convert --input-dir "Output\Export-<stamp>"  # AE star (if AE run)
  py .\build_unified_parquet.py --input-dir "Output\Export-<stamp>"        # or via -UnifiedParquet
  ```
  Inspect `manifest.json` → `sit_name_resolution`; ideally `unresolved_guids == 0` (proves the rule-pack XML covers sub-entity GUIDs). Records the live SIT-snapshot verification owner item.
- [ ] **Analytics end-to-end:** re-run the CE export (or A2) with `-UnifiedParquet`, and the AE export with `-PowerBIParquet`, confirming `C8TuningInput\` and `PowerBI-AE-Parquet-v6\` are produced with sane row counts + `manifest.json`.
- [ ] **D6 (RecordIdentity stability):** if two overlapping AE pulls were done, confirm dedup on `RecordIdentity` collapses the boundary-day overlap (informs the central-store track).
- [ ] **Success criteria:** parquet outputs build; SIT names resolve; counts reconcile against the export manifests.

### A6 — Larger export (scale validation)

- [ ] **Run** a bigger tenant, multi-terminal (`-WorkerCount 4`+), full scope, ideally cert auth if configured (else interactive). Monitor via the tail commands in CLAUDE.md. This stresses throughput, adaptive paging, retry-bucket, and worker reclaim at volume.
- [ ] **Success criteria:** completes without data loss; `RetryTasks.csv` (if any) is small and re-runs clean; analytics conversion succeeds at scale.

**Exit of Part A:** a signed-off statement of what ran live, what passed, and any fixes made. This unblocks confidence for Phase 1 and the central-store track.

---

## Part B — Phase 1: unattended foundation (implementation)

Build after (or alongside) Part A. Source: REVIEW-20260610 §C.4 ranks 1,2,3,5,6 (rank 9 already done). These enable scheduled, unattended (cert-auth) monthly runs. **Mostly decision-independent**; cert-store depth (rank 4) is Phase 3, central-store cadence (D1) is Phase 4.

### Task B1 — RunSummary.json + deterministic exit codes (H3; ranks 6+3) — do first, others build on it
**Files:** Create `Modules/Compl8ExportFunctions/Core/<NN>-RunResult.ps1` (e.g. `11-RunResult.ps1` under Core; register in `.psm1 $sectionFiles` + `Export-ModuleMember` + `.psd1`); Modify `App/MainExecution.ps1` (final exit path), and the AE/CE orchestrators where terminal phase is written.
- [ ] **Step 1 — Define the contract.** A function `Write-RunSummary -ExportDir <dir> -Result <hashtable>` that writes `<ExportDir>\RunSummary.json`: `{ schemaVersion, startedUtc, endedUtc, mode, exitCode, status (Completed|Partial|Failed|AuthFailed|ConfigError|Locked), sections:[{name,status,recordCount,errorCount}], remainingTasks:int, errors:[...capped] }`. And exit-code constants: `0` ok, `2` partial, `3` auth, `4` config, `5` lock.
- [ ] **Step 2 — Smoke-assert.** Add to `tests/test_powershell_smoke.py`: assert `Write-RunSummary` exists and the exit-code map is present; in-process, call `Write-RunSummary` with a sample result to a temp dir and assert the JSON shape/keys.
- [ ] **Step 3 — Wire MainExecution.** Compute the final status from section outcomes (AE/CE phase, `Write-RemainingTasksCsv` count, error log) and `exit $code`. Today `MainExecution.ps1` always exits 0 (REVIEW §A H3). Replace with the contract. Keep interactive behavior unchanged when not `-Unattended` (still write RunSummary; exit code still meaningful but won't surprise interactive users).
- [ ] **Step 4 — Verify:** parse-check + import; run the smoke asserts; manual `pwsh` dry of `Write-RunSummary`. **Commit.**

### Task B2 — `-Unattended` switch (rank 1)
**Files:** Modify `Export-Compl8Configuration.ps1` (param + propagate to `$script:Unattended`), `App/MainExecution.ps1` (prompt A + fallback H), `App/Orchestrator/ContentExplorer.Export/Resume/Retry/TasksCsv.ps1` (prompts B–E).
- [ ] **Step 1 — Add `[switch]$Unattended`** to the entry script; stash to a script-scope var visible to the dot-sourced parts (same mechanism as `$OutputDirectory`/`$scriptRoot`).
- [ ] **Step 2 — Gate each blocking prompt** (REVIEW §C.1 A–E) behind `if (-not $script:Unattended) { <prompt> } else { <deterministic default> }`: A "Proceed?" → proceed; B aggregate-reuse → **generate fresh** (deterministic at monthly cadence; do NOT reuse <30-day aggregates); C resume-confirm → proceed; D retry-confirm and E tasks-csv-confirm → proceed (NOT the silent exit-0 cancel). Find current lines by content (REVIEW cites are pre-refactor).
- [ ] **Step 3 — Fallback H:** in `MainExecution.ps1`, when no recognized mode and `-Unattended`, **fail-fast** with exit `4` (ConfigError) + RunSummary, instead of dropping to the interactive menu / quitting exit 0.
- [ ] **Step 4 — Smoke-assert + verify:** add smoke asserts that each prompt site is guarded by `$script:Unattended`; parse-check + import; in-process run a mode with redirected stdin + `-Unattended` and confirm no `Read-Host` is hit (mock `Read-Host` to throw, assert it isn't called). **Commit.**

### Task B3 — AuthConfig validation + cert pre-flight (rank 2)
**Files:** Modify `App/Host/Menu.ps1` `Build-AuthParameters` (or extract a `Test-AuthConfig` into `Core/03-Connection.ps1`), and the unattended entry path.
- [ ] **Step 1 — `Test-AuthConfig`**: validate `AuthConfig.json` shape (AppId, Organization, CertificateThumbprint OR cert path) and, when thumbprint, confirm the cert exists in `Cert:\CurrentUser\My` (or `LocalMachine\My`) and `NotAfter > now (+ buffer)`. Returns structured result.
- [ ] **Step 2 — Hard-fail in unattended:** if `-Unattended` and AuthConfig is missing/malformed/cert-not-found/expired, write RunSummary (status `AuthFailed`/`ConfigError`) and exit `3`/`4` — do NOT silently fall back to interactive (today `Build-AuthParameters` silently falls back, REVIEW §C.2). Interactive mode keeps the fallback.
- [ ] **Step 3 — Smoke-assert + verify:** asserts that the unattended path calls `Test-AuthConfig` and exits non-zero on failure; in-process test with a bogus AuthConfig → expect the failure result (no connect attempt). **Commit.**

### Task B4 — Worker hidden-window + shutdown on all paths (rank 5)
**Files:** Modify `App/Host/Menu.ps1` (`Start-WorkerTerminals` / the `Start-Process pwsh` spawn) and the worker-shutdown sites: CE fresh-multi already stops workers; add shutdown to **AE multi/resume** (`ActivityExplorer.ps1`) and **CE resume-multi** (`ContentExplorer.Resume.ps1`) (REVIEW §C.1 notes these leak N pwsh per run).
- [ ] **Step 1 — Hidden window in unattended:** when `-Unattended`, spawn workers with `-WindowStyle Hidden` and drop `-NoExit` (coordinate with the 15-min staleness reclaim — a worker that finishes and exits is fine; ensure the orchestrator's completion detection doesn't depend on the window staying open). Interactive keeps `-NoExit` for visibility.
- [ ] **Step 2 — Guaranteed shutdown:** add a `finally`/cleanup that stops spawned worker PIDs on the AE-multi, AE-resume, and CE-resume-multi completion/abort paths (mirror the CE fresh-multi stop logic).
- [ ] **Step 3 — Smoke-assert + verify:** asserts the AE/CE-resume paths call the worker-stop helper; parse-check + import. (Window-style + spawn behavior is validated in the next multi-terminal live run.) **Commit.**

### Task B5 — (optional, lower priority — schedule into Phase 1 only if time) 
Concurrency lockfile under `Output\` → exit `5` (rank 8); retention/log rotation (rank 10); keepalive for CE resume-multi / tasks-csv / AE-multi loops (rank 11). Each is small; same task shape (file, behavior, smoke-assert, commit). Cert-store options (rank 4) and auto-resume/retry passes (rank 7) are **Phase 3** — out of Phase 1 scope.

**Phase 1 exit:** a tenant can be exported unattended via a scheduled task: cert auth (pre-flighted), no prompts, deterministic exit code + `RunSummary.json`, hidden/clean workers. The scheduler wrapper itself (Phase 5) and central store (Phase 4) follow.

---

## Decisions still needed (gate later phases, NOT Part A or Phase-1 core)
- **D1** AE cadence vs 30-day API horizon → recommend weekly/fortnightly AE, monthly CE (Phase 4 / scheduler).
- **D2** keep a slim policy snapshot to feed `parquet_builder/policy.py`? → gates the descope (~1,540 lines, Phase 2).
- **D3** CE deletion semantics (live-set vs tombstones) → central store.
- **D4** raw-page retention after parquet build → janitor.
- **D5** central store location (metadata embeds personal-site URLs) → local/controlled disk.
- **D6** RecordIdentity stability across overlapping AE pulls → **verify during Part A** (A5).

## Spec coverage / self-review notes
- Part A covers every merged change category (H1/CE-refactors in A2/A4, AE fixes in A3, M5/M7 in A2, analytics+SIT in A5, scale in A6).
- Phase 1 covers REVIEW §C.4 ranks 1,2,3,5,6 (rank 9 done); ranks 4,7 deferred to Phase 3, ranks 8,10,11 optional B5, rank 12 (sovereign cloud) out of scope.
- Open confirm-points for the implementer (flagged, not placeholders): exact current line numbers of prompts A–E (shifted by the refactor — find by content), and the precise worker-spawn/stop call sites in the post-R6 split files.
