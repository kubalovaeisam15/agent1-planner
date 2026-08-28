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

foreach ($requiredFile in @($xmlPath, $templatePath, $durationsPath)) {
    if (-not [System.IO.File]::Exists($requiredFile)) { throw "Required file not found: $requiredFile" }
}
if ([System.IO.File]::Exists($mppPath)) { throw "Refusing to overwrite existing MPP: $mppPath" }

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

function Get-CollectionNames($collection) {
    $names = @()
    if ($null -eq $collection) { return $names }
    foreach ($item in $collection) {
        if ($null -eq $item) { continue }
        try { $names += [string]$item.Name } catch { $names += [string]$item }
    }
    return @($names | Sort-Object -Unique)
}

function Get-CorporateObjects($project) {
    $reports = @()
    try { $reports = @(Get-CollectionNames $project.Reports) } catch { }
    return [ordered]@{
        views = @(Get-CollectionNames $project.Views)
        task_tables = @(Get-CollectionNames $project.TaskTables)
        resource_tables = @(Get-CollectionNames $project.ResourceTables)
        task_filters = @(Get-CollectionNames $project.TaskFilters)
        resource_filters = @(Get-CollectionNames $project.ResourceFilters)
        task_groups = @(Get-CollectionNames $project.TaskGroups)
        resource_groups = @(Get-CollectionNames $project.ResourceGroups)
        reports = $reports
    }
}

$projectApp = $null
$exportSucceeded = $false
try {
    # The output starts as an exact copy of the corporate MPP. This preserves
    # all local views, tables, filters, groups, reports, fields and calendars.
    [System.IO.File]::Copy($templatePath, $mppPath, $false)

    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }
    $missing = [Type]::Missing

    $projectApp.FileOpenEx(
        $mppPath, $false, $missing, $missing, $missing, $missing, $true
    ) | Out-Null
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project did not open the corporate template copy" }

    $templateTaskCount = [int]$project.Tasks.Count
    $templateCalendars = @()
    foreach ($calendar in $project.BaseCalendars) {
        if ($null -ne $calendar) { $templateCalendars += [string]$calendar.Name }
    }
    $templateCustomFields = @(Get-CustomFields $projectApp)
    $templateCorporateObjects = Get-CorporateObjects $project
    $templateWindowName = [string]$project.Name

    # Remove only task rows. Project-level corporate objects stay in the MPP.
    while ($project.Tasks.Count -gt 0) {
        $firstTask = $project.Tasks.Item(1)
        if ($null -eq $firstTask) { throw "Unable to delete the first template task" }
        $firstTask.Delete()
    }

    # This Project build rejects XMLDOM merge through FileOpenEx. Import the
    # calculated schedule in a second Project window and transfer all task
    # rows into the still-open corporate template copy instead.
    $placeholder = $project.Tasks.Add("__AGENT1_IMPORT_TARGET__")
    $xmlText = [System.IO.File]::ReadAllText($xmlPath, [System.Text.Encoding]::UTF8)
    $openResult = $projectApp.OpenXML($xmlText)
    if ($openResult -ne 0) { throw "MS Project OpenXML returned $openResult" }
    $importProject = $projectApp.ActiveProject
    if ($null -eq $importProject) { throw "MS Project did not open the MSPDI schedule" }

    $projectApp.SelectAll() | Out-Null
    $projectApp.EditCopy() | Out-Null
    $projectApp.WindowActivate($templateWindowName) | Out-Null
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project could not reactivate the corporate template copy" }
    $projectApp.SelectRow(1, $false) | Out-Null
    $projectApp.EditPaste() | Out-Null

    foreach ($candidate in @($project.Tasks)) {
        if ($null -ne $candidate -and [string]$candidate.Name -eq "__AGENT1_IMPORT_TARGET__") {
            $candidate.Delete()
            break
        }
    }

    # Project 2021 can import valid MSPDI durations as zero. Reapply the same
    # values natively, then let Project calculate the resulting schedule.
    $durationOverrides = Get-Content -LiteralPath $durationsPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $durationOverrideCount = 0
    $durationMismatchCount = 0
    foreach ($entry in $durationOverrides) {
        $task = $project.Tasks.Item([int]$entry.id)
        if ($null -eq $task) { throw "Task for duration override not found: $($entry.id)" }
        $expectedDuration = [double]$entry.duration_minutes
        $task.Duration = $expectedDuration
        $durationOverrideCount += 1
        if ([Math]::Abs(([double]$task.Duration) - $expectedDuration) -gt 0.01) { $durationMismatchCount += 1 }
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
            $samples += [ordered]@{
                id = [int]$task.ID
                name = [string]$task.Name
                outline_level = [int]$task.OutlineLevel
                start = ([datetime]$task.Start).ToString("yyyy-MM-dd")
                finish = ([datetime]$task.Finish).ToString("yyyy-MM-dd")
                duration = [string]$task.DurationText
                duration_minutes = [double]$task.Duration
                predecessors = [string]$task.Predecessors
                constraint_type = [int]$task.ConstraintType
            }
        }
    }

    $outputCalendars = @()
    foreach ($calendar in $project.BaseCalendars) {
        if ($null -ne $calendar) { $outputCalendars += [string]$calendar.Name }
    }
    $outputCustomFields = @(Get-CustomFields $projectApp)
    $outputCorporateObjects = Get-CorporateObjects $project
    $missingCalendars = @($templateCalendars | Where-Object { $_ -notin $outputCalendars })
    $missingFields = @($templateCustomFields | Where-Object {
        $expected = $_
        -not ($outputCustomFields | Where-Object {
            $_.task_scope -eq $expected.task_scope -and
            $_.internal_name -eq $expected.internal_name -and
            $_.alias -eq $expected.alias -and $_.formula -eq $expected.formula
        })
    })
    if ($missingCalendars.Count -gt 0) {
        throw "Corporate calendars lost during task replacement: $($missingCalendars -join ', ')"
    }
    if ($missingFields.Count -gt 0) {
        throw "Corporate custom fields lost during task replacement: $($missingFields.Count)"
    }
    $missingCorporateObjects = [ordered]@{}
    foreach ($category in $templateCorporateObjects.Keys) {
        $expectedItems = @($templateCorporateObjects[$category])
        $actualItems = @($outputCorporateObjects[$category])
        $missingItems = @($expectedItems | Where-Object { $_ -notin $actualItems })
        if ($missingItems.Count -gt 0) { $missingCorporateObjects[$category] = $missingItems }
    }
    if ($missingCorporateObjects.Count -gt 0) {
        throw "Corporate views/tables/filters/groups lost: $($missingCorporateObjects.Keys -join ', ')"
    }

    $projectApp.FileSave() | Out-Null
    if (-not [System.IO.File]::Exists($mppPath)) { throw "MS Project did not save the output MPP" }

    $report = [ordered]@{
        export_mode = "corporate-template-base-task-transfer"
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
        template_corporate_objects = $templateCorporateObjects
        output_calendars = $outputCalendars
        output_custom_fields = $outputCustomFields
        output_corporate_objects = $outputCorporateObjects
        missing_template_calendars = $missingCalendars
        missing_template_custom_fields = $missingFields
        missing_template_corporate_objects = $missingCorporateObjects
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
    if (-not $exportSucceeded -and [System.IO.File]::Exists($mppPath)) {
        [System.IO.File]::Delete($mppPath)
    }
}
