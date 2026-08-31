param(
    [Parameter(Mandatory = $true)][string]$InputMpp,
    [Parameter(Mandatory = $true)][string]$OutputJson
)

$ErrorActionPreference = "Stop"
$mppPath = [System.IO.Path]::GetFullPath($InputMpp)
$jsonPath = [System.IO.Path]::GetFullPath($OutputJson)

if (-not [System.IO.File]::Exists($mppPath)) {
    throw "Input MPP not found: $mppPath"
}
if ([System.IO.File]::Exists($jsonPath)) {
    throw "Refusing to overwrite existing JSON: $jsonPath"
}

function Convert-ProjectDate($value) {
    try { return ([datetime]$value).ToString("yyyy-MM-dd") }
    catch { return $null }
}

function Convert-ProjectNumber($value) {
    try { return [double]$value }
    catch { return $null }
}

$projectApp = $null
try {
    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }
    $missing = [Type]::Missing
    $projectApp.FileOpenEx(
        $mppPath, $true, $missing, $missing, $missing, $missing, $true
    ) | Out-Null
    $project = $projectApp.ActiveProject
    if ($null -eq $project) {
        throw "MS Project did not open the input MPP"
    }
    $calculationMode = [int]$projectApp.Calculation
    $tasks = [System.Collections.Generic.List[object]]::new()
    foreach ($task in $project.Tasks) {
        if ($null -eq $task) { continue }
        $tasks.Add([ordered]@{
            uid = [int]$task.UniqueID
            id = [int]$task.ID
            name = [string]$task.Name
            outline_level = [int]$task.OutlineLevel
            summary = [bool]$task.Summary
            milestone = [bool]$task.Milestone
            start = Convert-ProjectDate $task.Start
            finish = Convert-ProjectDate $task.Finish
            duration_minutes = Convert-ProjectNumber $task.Duration
            duration_text = [string]$task.DurationText
            duration_calendar_days = [string]$task.DurationText
            percent_complete = [int]$task.PercentComplete
            critical = [bool]$task.Critical
            total_slack_minutes = Convert-ProjectNumber $task.TotalSlack
            constraint_type = [int]$task.ConstraintType
            constraint_date = Convert-ProjectDate $task.ConstraintDate
            deadline = Convert-ProjectDate $task.Deadline
            predecessors = [string]$task.Predecessors
        })
    }

    $snapshot = [ordered]@{
        name = [string]$project.Name
        start = Convert-ProjectDate $project.ProjectStart
        finish = Convert-ProjectDate $project.ProjectFinish
        calculation_mode = $calculationMode
        tasks = $tasks
    }
    $snapshot | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath $jsonPath -Encoding UTF8
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.FileCloseEx(0, $true) | Out-Null } catch { }
        # pjDoNotSave = 0. Passing it explicitly prevents a hidden save prompt.
        try { $projectApp.Quit(0) | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
