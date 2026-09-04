param(
    [Parameter(Mandatory = $true)][string]$InputMpp,
    [Parameter(Mandatory = $true)][string]$OutputMpp,
    [Parameter(Mandatory = $true)][string]$ReportJson,
    [int]$LeadDays = 15
)

$ErrorActionPreference = "Stop"
$inputPath = [System.IO.Path]::GetFullPath($InputMpp)
$outputPath = [System.IO.Path]::GetFullPath($OutputMpp)
$reportPath = [System.IO.Path]::GetFullPath($ReportJson)

if (-not [System.IO.File]::Exists($inputPath)) {
    throw "Input MPP not found: $inputPath"
}
foreach ($newPath in @($outputPath, $reportPath)) {
    if ([System.IO.File]::Exists($newPath)) {
        throw "Refusing to overwrite existing output: $newPath"
    }
}
if ($LeadDays -lt 0) { throw "LeadDays must be non-negative" }

function Convert-ProjectDate($value) {
    try { return ([datetime]$value).ToString("yyyy-MM-ddTHH:mm:ss") }
    catch { return $null }
}

function Get-Utf8Text([string]$base64) {
    return [System.Text.Encoding]::UTF8.GetString(
        [System.Convert]::FromBase64String($base64)
    )
}

function Get-TaskArray($project) {
    $items = @()
    foreach ($task in $project.Tasks) {
        if ($null -ne $task) { $items += $task }
    }
    return @($items)
}

function Find-ScopeEndId($tasks, $rootId, $rootLevel) {
    $next = $tasks |
        Where-Object { [int]$_.ID -gt $rootId -and [int]$_.OutlineLevel -le $rootLevel } |
        Sort-Object { [int]$_.ID } |
        Select-Object -First 1
    if ($null -eq $next) { return [int]::MaxValue }
    return [int]$next.ID
}

function Has-DirectPredecessor($task, $predecessorId) {
    $pattern = "(^|;)${predecessorId}(?!\d)"
    return ([string]$task.Predecessors) -match $pattern
}

$projectApp = $null
$completed = $false
try {
    [System.IO.File]::Copy($inputPath, $outputPath, $false)

    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }
    $missing = [Type]::Missing
    $projectApp.FileOpenEx(
        $outputPath, $false, $missing, $missing, $missing, $missing, $true
    ) | Out-Null
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project did not open the output copy" }

    # Apply the batch in manual mode so Project does not recalculate the full
    # network after every individual constraint assignment.
    try { $projectApp.Calculation = 0 } catch { }

    # Decode localized labels so Windows PowerShell 5.1 does not depend on the
    # script file's BOM/legacy code-page detection.
    $tenderRootName = Get-Utf8Text("0KLQtdC90LTQtdGA0Ysg0KHQnNCg")
    $smrRootName = Get-Utf8Text("0KHQnNCg")
    $preparationPattern = Get-Utf8Text("XtCf0L7QtNCz0L7RgtC+0LLQutCwINCi0Jc=")
    $contractPattern = Get-Utf8Text("XtCX0LDQutC70Y7Rh9C10L3QuNC1INC00L7Qs9C+0LLQvtGA0LA=")
    $constraintName = Get-Utf8Text("0J3QsNGH0LDQu9C+INC90LUg0YDQsNC90LXQtQ==")

    $tasks = @(Get-TaskArray $project)
    $tenderRoots = @($tasks | Where-Object {
        [string]$_.Name -eq $tenderRootName -and [int]$_.OutlineLevel -eq 2
    })
    if ($tenderRoots.Count -ne 1) {
        throw "Expected exactly one level-2 tender summary; found $($tenderRoots.Count)"
    }
    $smrRoots = @($tasks | Where-Object {
        [string]$_.Name -eq $smrRootName -and [int]$_.OutlineLevel -eq 1
    })
    if ($smrRoots.Count -ne 1) {
        throw "Expected exactly one level-1 SMR summary; found $($smrRoots.Count)"
    }

    $tenderRoot = $tenderRoots[0]
    $smrRoot = $smrRoots[0]
    $tenderEndId = Find-ScopeEndId $tasks ([int]$tenderRoot.ID) ([int]$tenderRoot.OutlineLevel)
    $smrEndId = Find-ScopeEndId $tasks ([int]$smrRoot.ID) ([int]$smrRoot.OutlineLevel)

    $groups = @($tasks | Where-Object {
        [int]$_.ID -gt [int]$tenderRoot.ID -and
        [int]$_.ID -lt $tenderEndId -and
        [int]$_.OutlineLevel -eq 3 -and [bool]$_.Summary
    } | Sort-Object { [int]$_.ID })

    $plans = @()
    $skipped = @()
    foreach ($group in $groups) {
        $groupEnd = $tasks |
            Where-Object {
                [int]$_.ID -gt [int]$group.ID -and
                [int]$_.ID -lt $tenderEndId -and
                [int]$_.OutlineLevel -le 3
            } |
            Sort-Object { [int]$_.ID } |
            Select-Object -First 1
        $groupEndId = if ($null -eq $groupEnd) { $tenderEndId } else { [int]$groupEnd.ID }
        $children = @($tasks | Where-Object {
            [int]$_.ID -gt [int]$group.ID -and [int]$_.ID -lt $groupEndId -and -not [bool]$_.Summary
        })
        $preparations = @($children | Where-Object { [string]$_.Name -match $preparationPattern })
        $contracts = @($children | Where-Object { [string]$_.Name -match $contractPattern })
        if ($preparations.Count -ne 1 -or $contracts.Count -ne 1) {
            $skipped += [ordered]@{
                group_id = [int]$group.ID
                group_name = [string]$group.Name
                reason = "requires exactly one preparation and one contract task"
                preparation_count = $preparations.Count
                contract_count = $contracts.Count
            }
            continue
        }

        $preparation = $preparations[0]
        $contract = $contracts[0]
        $smrSuccessors = @($tasks | Where-Object {
            [int]$_.ID -gt [int]$smrRoot.ID -and
            [int]$_.ID -lt $smrEndId -and
            -not [bool]$_.Summary -and [double]$_.Duration -gt 0 -and
            (Has-DirectPredecessor $_ ([int]$contract.ID))
        } | Sort-Object { [datetime]$_.Start }, { [int]$_.ID })
        if ($smrSuccessors.Count -eq 0) {
            $skipped += [ordered]@{
                group_id = [int]$group.ID
                group_name = [string]$group.Name
                reason = "contract has no direct non-summary, non-milestone successor in the SMR block"
                preparation_id = [int]$preparation.ID
                contract_id = [int]$contract.ID
            }
            continue
        }

        $earliestStart = [datetime]$smrSuccessors[0].Start
        $sameEarliest = @($smrSuccessors | Where-Object { [datetime]$_.Start -eq $earliestStart })
        $chainSpan = ([datetime]$contract.Finish) - ([datetime]$preparation.Start)
        if ($chainSpan.TotalDays -lt 0) {
            throw "Negative tender chain span in group $([int]$group.ID)"
        }
        $targetContractFinish = $earliestStart.AddDays(-$LeadDays)
        $constraintDate = $targetContractFinish.Subtract($chainSpan)
        $plans += [ordered]@{
            group = $group
            preparation = $preparation
            contract = $contract
            successors = $smrSuccessors
            same_earliest = $sameEarliest
            earliest_start_before = $earliestStart
            preparation_start_before = [datetime]$preparation.Start
            chain_span = $chainSpan
            target_contract_finish = $targetContractFinish
            constraint_date = $constraintDate
            old_constraint_type = [int]$preparation.ConstraintType
            old_constraint_date = $preparation.ConstraintDate
        }
    }

    if ($plans.Count -eq 0) { throw "No qualifying tender chains were found" }

    foreach ($plan in $plans) {
        # Microsoft PjConstraint: pjSNET (Start No Earlier Than) = 4.
        $plan.preparation.ConstraintType = 4
        $plan.preparation.ConstraintDate = [datetime]$plan.constraint_date
    }
    # DEC-39: perform one full calculation and save only in automatic mode.
    try { $projectApp.Calculation = -1 } catch { }
    if ([int]$projectApp.Calculation -ne -1) {
        throw "Microsoft Project automatic calculation could not be enabled"
    }
    $projectApp.CalculateProject() | Out-Null

    # Calendar exceptions and external predecessors inside a tender group can
    # make a raw calendar-span subtraction differ from Project's calculated
    # chain. Calibrate only through the requested preparation constraint. If a
    # negative correction would put SNET before the original preparation start,
    # existing predecessors prevent that constraint type from pulling the chain
    # earlier; retain the formula date and report the remaining shortfall.
    $calibrationRounds = 0
    for ($round = 1; $round -le 4; $round++) {
        $corrections = @()
        foreach ($plan in $plans) {
            $finishDeltaDays = ($plan.target_contract_finish - [datetime]$plan.contract.Finish).TotalDays
            if ([Math]::Abs($finishDeltaDays) -le (1.0 / 1440.0)) { continue }
            $currentConstraint = [datetime]$plan.preparation.ConstraintDate
            if ($finishDeltaDays -lt 0 -and $currentConstraint -le $plan.preparation_start_before) {
                continue
            }
            $corrections += [ordered]@{
                plan = $plan
                date = $currentConstraint.AddDays($finishDeltaDays)
            }
        }
        if ($corrections.Count -eq 0) { break }
        try { $projectApp.Calculation = 0 } catch { }
        foreach ($correction in $corrections) {
            $correction.plan.preparation.ConstraintDate = [datetime]$correction.date
            $correction.plan.constraint_date = [datetime]$correction.date
        }
        try { $projectApp.Calculation = -1 } catch { }
        if ([int]$projectApp.Calculation -ne -1) {
            throw "Microsoft Project automatic calculation could not be re-enabled"
        }
        $projectApp.CalculateProject() | Out-Null
        $calibrationRounds = $round
    }

    $adjustments = @()
    foreach ($plan in $plans) {
        $successorsAfter = @($plan.successors | Sort-Object { [datetime]$_.Start }, { [int]$_.ID })
        $earliestAfter = [datetime]$successorsAfter[0].Start
        $contractFinishAfter = [datetime]$plan.contract.Finish
        $actualLeadDays = ($earliestAfter - $contractFinishAfter).TotalDays
        if ([int]$plan.preparation.ConstraintType -ne 4) {
            throw "Constraint type verification failed for task $([int]$plan.preparation.ID)"
        }
        if ([Math]::Abs((([datetime]$plan.preparation.ConstraintDate) - ([datetime]$plan.constraint_date)).TotalMinutes) -gt 1) {
            throw "Constraint date verification failed for task $([int]$plan.preparation.ID)"
        }
        $adjustments += [ordered]@{
            group_id = [int]$plan.group.ID
            group_name = [string]$plan.group.Name
            preparation_id = [int]$plan.preparation.ID
            preparation_uid = [int]$plan.preparation.UniqueID
            preparation_name = [string]$plan.preparation.Name
            old_constraint_type = [int]$plan.old_constraint_type
            old_constraint_date = Convert-ProjectDate $plan.old_constraint_date
            new_constraint_type = [int]$plan.preparation.ConstraintType
            new_constraint_name = $constraintName
            new_constraint_date = Convert-ProjectDate $plan.preparation.ConstraintDate
            chain_span_calendar_days = [double]$plan.chain_span.TotalDays
            contract_id = [int]$plan.contract.ID
            contract_name = [string]$plan.contract.Name
            contract_finish_after = Convert-ProjectDate $plan.contract.Finish
            target_contract_finish = Convert-ProjectDate $plan.target_contract_finish
            earliest_smr_start_before = Convert-ProjectDate $plan.earliest_start_before
            earliest_smr_start_after = Convert-ProjectDate $earliestAfter
            actual_lead_calendar_days = [double]$actualLeadDays
            meets_requested_lead_after_recalculation = ($actualLeadDays -ge $LeadDays)
            earliest_smr_successors = @($plan.same_earliest | ForEach-Object {
                [ordered]@{
                    id = [int]$_.ID
                    uid = [int]$_.UniqueID
                    name = [string]$_.Name
                    start = Convert-ProjectDate $_.Start
                    predecessors = [string]$_.Predecessors
                }
            })
        }
    }

    $projectApp.FileSave() | Out-Null
    if ([int]$projectApp.Calculation -ne -1) {
        throw "Output MPP was not saved with automatic calculation enabled"
    }
    if (-not [System.IO.File]::Exists($outputPath)) {
        throw "MS Project did not save the output MPP"
    }

    $report = [ordered]@{
        source_mpp = $inputPath
        output_mpp = $outputPath
        project_name = [string]$project.Name
        calculation_mode = [int]$projectApp.Calculation
        tender_root_id = [int]$tenderRoot.ID
        smr_root_id = [int]$smrRoot.ID
        lead_calendar_days = $LeadDays
        calibration_rounds = $calibrationRounds
        adjusted_count = $adjustments.Count
        recalculation_warning_count = @($adjustments | Where-Object {
            -not $_.meets_requested_lead_after_recalculation
        }).Count
        skipped_count = $skipped.Count
        adjustments = $adjustments
        skipped = $skipped
    }
    $report | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $completed = $true
}
catch {
    $line = $_.InvocationInfo.ScriptLineNumber
    $source = $_.InvocationInfo.Line
    throw "$($_.Exception.Message) (line ${line}: $source)"
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.FileCloseEx(0, $true) | Out-Null } catch { }
        try { $projectApp.Quit(0) | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (-not $completed) {
        if ([System.IO.File]::Exists($outputPath)) { [System.IO.File]::Delete($outputPath) }
        if ([System.IO.File]::Exists($reportPath)) { [System.IO.File]::Delete($reportPath) }
    }
}
