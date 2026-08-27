param(
    [Parameter(Mandatory = $true)][string]$InputXml,
    [Parameter(Mandatory = $true)][string]$OutputMpp,
    [Parameter(Mandatory = $true)][string]$ReportJson,
    [Parameter(Mandatory = $true)][string]$DurationsJson,
    [string]$SampleTaskIds = ""
)

$ErrorActionPreference = "Stop"
$xmlPath = [System.IO.Path]::GetFullPath($InputXml)
$mppPath = [System.IO.Path]::GetFullPath($OutputMpp)
$reportPath = [System.IO.Path]::GetFullPath($ReportJson)
$durationsPath = [System.IO.Path]::GetFullPath($DurationsJson)

if (-not [System.IO.File]::Exists($xmlPath)) {
    throw "Input XML not found: $xmlPath"
}
if ([System.IO.File]::Exists($mppPath)) {
    throw "Refusing to overwrite existing MPP: $mppPath"
}
if (-not [System.IO.File]::Exists($durationsPath)) {
    throw "Duration overrides not found: $durationsPath"
}

$projectApp = $null
try {
    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }

    $xmlText = [System.IO.File]::ReadAllText($xmlPath, [System.Text.Encoding]::UTF8)
    $openResult = $projectApp.OpenXML($xmlText)
    if ($openResult -ne 0) {
        throw "MS Project OpenXML returned $openResult"
    }

    $project = $projectApp.ActiveProject
    if ($null -eq $project) {
        throw "MS Project did not create an active project"
    }

    # Project 2021 preserves the imported hierarchy and links but can replace
    # MSPDI durations with zero. Assigning the same values through the native
    # COM property makes Project retain and calculate them reliably.
    $durationOverrides = Get-Content -LiteralPath $durationsPath -Raw -Encoding UTF8 |
        ConvertFrom-Json
    $durationOverrideCount = 0
    $durationMismatchCount = 0
    foreach ($entry in $durationOverrides) {
        $task = $project.Tasks.Item([int]$entry.id)
        if ($null -eq $task) {
            throw "Task for duration override not found: $($entry.id)"
        }
        $expectedDuration = [double]$entry.duration_minutes
        $task.Duration = $expectedDuration
        $durationOverrideCount += 1
        if ([Math]::Abs(([double]$task.Duration) - $expectedDuration) -gt 0.01) {
            $durationMismatchCount += 1
        }
    }
    $projectApp.CalculateProject() | Out-Null

    $taskCount = [int]$project.Tasks.Count
    $sampleIds = @{}
    foreach ($sampleId in $SampleTaskIds.Split(',')) {
        if ($sampleId) { $sampleIds[$sampleId] = $true }
    }
    $samples = @()
    foreach ($sampleId in $sampleIds.Keys) {
        $task = $project.Tasks.Item([int]$sampleId)
        if ($null -ne $task) {
            $taskFinish = [datetime]$task.Finish
            $samples += [ordered]@{
                id = [int]$task.ID
                name = [string]$task.Name
                outline_level = [int]$task.OutlineLevel
                start = ([datetime]$task.Start).ToString("yyyy-MM-dd")
                finish = $taskFinish.ToString("yyyy-MM-dd")
                duration = [string]$task.DurationText
                duration_minutes = [double]$task.Duration
                predecessors = [string]$task.Predecessors
                constraint_type = [int]$task.ConstraintType
            }
        }
    }

    $saveResult = $projectApp.FileSaveAs($mppPath)
    if (-not [System.IO.File]::Exists($mppPath)) {
        throw "MS Project did not save MPP (FileSaveAs=$saveResult)"
    }

    $report = [ordered]@{
        task_count = $taskCount
        duration_override_count = $durationOverrideCount
        duration_mismatch_count = $durationMismatchCount
        project_name = [string]$project.Name
        project_start = ([datetime]$project.ProjectStart).ToString("yyyy-MM-dd")
        project_finish = ([datetime]$project.ProjectFinish).ToString("yyyy-MM-dd")
        samples = $samples
        output_mpp = $mppPath
    }
    $report | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $reportPath -Encoding UTF8
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.Quit() | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
