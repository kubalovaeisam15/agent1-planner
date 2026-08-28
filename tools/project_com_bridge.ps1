param(
    [Parameter(Mandatory = $true)][string]$InputXml,
    [Parameter(Mandatory = $true)][string]$TemplateMpp,
    [Parameter(Mandatory = $true)][string]$OutputMpp,
    [Parameter(Mandatory = $true)][string]$ReportJson,
    [Parameter(Mandatory = $true)][string]$DurationsJson,
    [string]$SampleTaskIds = ""
)

$ErrorActionPreference = "Stop"
$xmlPath = [System.IO.Path]::GetFullPath($InputXml)
$templatePath = [System.IO.Path]::GetFullPath($TemplateMpp)
$mppPath = [System.IO.Path]::GetFullPath($OutputMpp)
$reportPath = [System.IO.Path]::GetFullPath($ReportJson)
$durationsPath = [System.IO.Path]::GetFullPath($DurationsJson)

if (-not [System.IO.File]::Exists($xmlPath)) {
    throw "Input XML not found: $xmlPath"
}
if (-not [System.IO.File]::Exists($templatePath)) {
    throw "Template MPP not found: $templatePath"
}
if ([System.IO.File]::Exists($mppPath)) {
    throw "Refusing to overwrite existing MPP: $mppPath"
}
if (-not [System.IO.File]::Exists($durationsPath)) {
    throw "Duration overrides not found: $durationsPath"
}

function Get-CustomFields($application) {
    $families = @(
        @{ prefix = "Text"; count = 30 }, @{ prefix = "Number"; count = 20 },
        @{ prefix = "Cost"; count = 10 }, @{ prefix = "Duration"; count = 10 },
        @{ prefix = "Start"; count = 10 }, @{ prefix = "Finish"; count = 10 },
        @{ prefix = "Flag"; count = 20 }, @{ prefix = "Outline Code"; count = 10 }
    )
    $items = @()
    foreach ($taskScope in @($true, $false)) {
        $fieldType = if ($taskScope) { 0 } else { 1 }
        foreach ($family in $families) {
            for ($index = 1; $index -le $family.count; $index++) {
                $internalName = "$($family.prefix)$index"
                try {
                    $fieldId = $application.FieldNameToFieldConstant($internalName, $fieldType)
                    $alias = [string]$application.CustomFieldGetName($fieldId)
                    $formula = [string]$application.CustomFieldGetFormula($fieldId)
                    if ($alias -or $formula) {
                        $items += [ordered]@{
                            task_scope = $taskScope
                            internal_name = $internalName
                            alias = $alias
                            formula = $formula
                        }
                    }
                } catch { }
            }
        }
    }
    return $items
}

$projectApp = $null
$exportSucceeded = $false
$templateWorkPath = Join-Path ([System.IO.Path]::GetDirectoryName($mppPath)) `
    ("agent1-template-" + [guid]::NewGuid().ToString("N") + ".mpp")
try {
    # OrganizerMoveItem is applied to a disposable copy, so the corporate
    # source file cannot be changed even if a Project version implements
    # "move" literally rather than as the Organizer's usual copy operation.
    [System.IO.File]::Copy($templatePath, $templateWorkPath, $false)

    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }
    $missing = [Type]::Missing

    # Inventory corporate calendars before building the calculated project.
    $projectApp.FileOpenEx(
        $templateWorkPath, $true, $missing, $missing, $missing, $missing, $true
    ) | Out-Null
    $templateProject = $projectApp.ActiveProject
    if ($null -eq $templateProject) { throw "MS Project did not open the template" }
    $templateTaskCount = [int]$templateProject.Tasks.Count
    $templateCalendars = @()
    foreach ($calendar in $templateProject.BaseCalendars) {
        if ($null -ne $calendar) { $templateCalendars += [string]$calendar.Name }
    }
    $templateCustomFields = @(Get-CustomFields $projectApp)

    # MSPDI remains the schedule transport because it preserves the generated
    # hierarchy, links and elapsed-day lags exactly.
    $xmlText = [System.IO.File]::ReadAllText($xmlPath, [System.Text.Encoding]::UTF8)
    $openResult = $projectApp.OpenXML($xmlText)
    if ($openResult -ne 0) { throw "MS Project OpenXML returned $openResult" }
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project did not create an active project" }

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

    # Project 2021 rejects the documented "copy all" mode, so corporate
    # calendars and named custom fields are copied one by one.
    $organizerResults = @()
    foreach ($calendarName in $templateCalendars) {
        try {
            $moved = $projectApp.OrganizerMoveItem(
                5, $templateWorkPath, $mppPath, $calendarName, $true
            )
        } catch {
            throw "Organizer calendar copy failed [$calendarName]: $($_.Exception.Message)"
        }
        $organizerResults += [ordered]@{
            type = 5; task_scope = $true; name = $calendarName; copied = [bool]$moved
        }
    }
    foreach ($field in $templateCustomFields) {
        $fieldType = if ($field.task_scope) { 0 } else { 1 }
        $fieldId = $projectApp.FieldNameToFieldConstant($field.internal_name, $fieldType)
        try {
            $renamed = $true
            if ($field.alias) {
                $renamed = [bool]$projectApp.CustomFieldRename($fieldId, $field.alias)
            }
            $formulaSet = $true
            if ($field.formula) {
                $formulaSet = [bool]$projectApp.CustomFieldSetFormula($fieldId, $field.formula)
            }
            $moved = $renamed -and $formulaSet
        } catch {
            throw "Custom field copy failed [$($field.internal_name)]: $($_.Exception.Message)"
        }
        $organizerResults += [ordered]@{
            type = 9
            task_scope = [bool]$field.task_scope
            name = $field.alias
            internal_name = $field.internal_name
            copied = [bool]$moved
        }
    }
    $mandatoryCopies = @($organizerResults | Where-Object { -not $_.copied })
    if ($mandatoryCopies.Count -gt 0) {
        throw "Microsoft Project did not copy all corporate fields/calendars"
    }

    $projectApp.FileSave() | Out-Null

    # Verify in the active output that the task transport and all template
    # calendars survived the Organizer operation.
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project lost the active output MPP" }
    $outputCalendars = @()
    foreach ($calendar in $project.BaseCalendars) {
        if ($null -ne $calendar) { $outputCalendars += [string]$calendar.Name }
    }
    $outputCustomFields = @(Get-CustomFields $projectApp)
    $missingCalendars = @($templateCalendars | Where-Object { $_ -notin $outputCalendars })
    if ($missingCalendars.Count -gt 0) {
        throw "Corporate calendars lost during copy: $($missingCalendars -join ', ')"
    }
    if ([int]$project.Tasks.Count -ne $taskCount) {
        throw "Organizer changed the task count: $($project.Tasks.Count) instead of $taskCount"
    }
    $missingFields = @($templateCustomFields | Where-Object {
        $expected = $_
        -not ($outputCustomFields | Where-Object {
            $_.task_scope -eq $expected.task_scope -and
            $_.internal_name -eq $expected.internal_name -and
            $_.alias -eq $expected.alias -and $_.formula -eq $expected.formula
        })
    })
    if ($missingFields.Count -gt 0) {
        throw "Corporate custom fields lost during copy: $($missingFields.Count)"
    }

    $report = [ordered]@{
        task_count = $taskCount
        duration_override_count = $durationOverrideCount
        duration_mismatch_count = $durationMismatchCount
        project_name = [string]$project.Name
        project_start = ([datetime]$project.ProjectStart).ToString("yyyy-MM-dd")
        project_finish = ([datetime]$project.ProjectFinish).ToString("yyyy-MM-dd")
        template_mpp = $templatePath
        template_task_count = $templateTaskCount
        template_calendars = $templateCalendars
        template_custom_fields = $templateCustomFields
        output_calendars = $outputCalendars
        output_custom_fields = $outputCustomFields
        missing_template_calendars = $missingCalendars
        missing_template_custom_fields = $missingFields
        organizer_copies = $organizerResults
        samples = $samples
        output_mpp = $mppPath
    }
    $report | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    $exportSucceeded = $true
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.FileCloseEx(0, $true) | Out-Null } catch { }
        try { $projectApp.Quit(0) | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ([System.IO.File]::Exists($templateWorkPath)) {
        [System.IO.File]::Delete($templateWorkPath)
    }
    if (-not $exportSucceeded -and [System.IO.File]::Exists($mppPath)) {
        [System.IO.File]::Delete($mppPath)
    }
}
