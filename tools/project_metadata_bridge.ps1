param(
    [Parameter(Mandatory = $true)][string]$InputMpp,
    [Parameter(Mandatory = $true)][string]$OutputJson
)

$ErrorActionPreference = "Stop"
$mppPath = [System.IO.Path]::GetFullPath($InputMpp)
$jsonPath = [System.IO.Path]::GetFullPath($OutputJson)
if (-not [System.IO.File]::Exists($mppPath)) { throw "Input MPP not found: $mppPath" }
if ([System.IO.File]::Exists($jsonPath)) { throw "Refusing to overwrite JSON: $jsonPath" }

function Get-CustomFields($application) {
    $families = @(
        @{ prefix = "Text"; count = 30 }, @{ prefix = "Number"; count = 20 },
        @{ prefix = "Cost"; count = 10 }, @{ prefix = "Duration"; count = 10 },
        @{ prefix = "Start"; count = 10 }, @{ prefix = "Finish"; count = 10 },
        @{ prefix = "Flag"; count = 20 }, @{ prefix = "Outline Code"; count = 10 }
    )
    $items = @()
    foreach ($taskScope in @($true, $false)) {
        $fieldType = if ($taskScope) { 0 } else { 1 } # pjTask / pjResource
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
try {
    $projectApp = New-Object -ComObject MSProject.Application
    try { $projectApp.Visible = $false } catch { }
    try { $projectApp.DisplayAlerts = $false } catch { }
    $missing = [Type]::Missing
    $projectApp.FileOpenEx(
        $mppPath, $true, $missing, $missing, $missing, $missing, $true
    ) | Out-Null
    $project = $projectApp.ActiveProject
    if ($null -eq $project) { throw "MS Project did not open MPP" }
    $calendars = @()
    foreach ($calendar in $project.BaseCalendars) {
        if ($null -ne $calendar) { $calendars += [string]$calendar.Name }
    }
    [ordered]@{
        mpp = $mppPath
        task_count = [int]$project.Tasks.Count
        calendars = $calendars
        custom_fields = @(Get-CustomFields $projectApp)
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.FileCloseEx(0, $true) | Out-Null } catch { }
        try { $projectApp.Quit(0) | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
