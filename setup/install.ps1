param([switch]$CheckOnly)

$ErrorActionPreference = "Stop"
$projectRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))

function Convert-ToTomlBasicString([string]$value) {
    return $value.Replace('\', '\\').Replace('"', '\"')
}

$requiredFiles = @(
    "AGENTS.md",
    "CLAUDE.md",
    "instructions/typGRP.md",
    "instructions/bindings.md",
    "instructions/standards.md",
    "tools/mcp_server.py",
    "tools/mspdi_adapter.py",
    "tools/mpp_validator.py"
)

$missingFiles = @($requiredFiles | Where-Object {
    -not [System.IO.File]::Exists((Join-Path $projectRoot $_))
})
if ($missingFiles.Count -gt 0) {
    throw "Package files are missing: $($missingFiles -join ', ')"
}
$xlsxTemplates = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot "instructions") -Filter "*.xlsx" -File)
if ($xlsxTemplates.Count -ne 1) {
    throw "Expected exactly one XLSX template in instructions; found $($xlsxTemplates.Count)"
}
$mppTemplates = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot "data") -Filter "*.mpp" -File)
if ($mppTemplates.Count -ne 1) {
    throw "Expected exactly one corporate MPP template in data; found $($mppTemplates.Count)"
}

$pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
if ($null -eq $pythonCommand) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
}
if ($null -eq $pythonCommand) {
    throw "Python was not found in PATH. Install Python 3.12 and try again."
}
$pythonPath = [System.IO.Path]::GetFullPath($pythonCommand.Source)
$pythonVersion = (& $pythonPath --version 2>&1 | Out-String).Trim()
if ($LASTEXITCODE -ne 0) { throw "Unable to run Python: $pythonPath" }

$projectApp = $null
try {
    $projectApp = New-Object -ComObject MSProject.Application
    $projectVersion = [string]$projectApp.Version
}
catch {
    throw "Microsoft Project COM is unavailable. Install Microsoft Project Desktop. $($_.Exception.Message)"
}
finally {
    if ($null -ne $projectApp) {
        try { $projectApp.Quit(0) | Out-Null } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($projectApp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if (-not $CheckOnly) {
    $codexDirectory = Join-Path $projectRoot ".codex"
    [System.IO.Directory]::CreateDirectory($codexDirectory) | Out-Null
    $pythonToml = Convert-ToTomlBasicString $pythonPath
    $serverToml = Convert-ToTomlBasicString (Join-Path $projectRoot "tools/mcp_server.py")
    $codexConfig = @"
[mcp_servers.agent1-ms-project]
command = "$pythonToml"
args = ["$serverToml"]
startup_timeout_sec = 10
tool_timeout_sec = 1800
enabled = true
"@
    [System.IO.File]::WriteAllText(
        (Join-Path $codexDirectory "config.toml"),
        $codexConfig,
        [System.Text.UTF8Encoding]::new($false)
    )

    $mcpConfig = [ordered]@{
        mcpServers = [ordered]@{
            "agent1-ms-project" = [ordered]@{
                type = "stdio"
                command = $pythonPath
                args = @((Join-Path $projectRoot "tools/mcp_server.py"))
                env = [ordered]@{}
            }
        }
    }
    [System.IO.File]::WriteAllText(
        (Join-Path $projectRoot ".mcp.json"),
        ($mcpConfig | ConvertTo-Json -Depth 6),
        [System.Text.UTF8Encoding]::new($false)
    )
}

Write-Host "MS Project Agent: environment check passed" -ForegroundColor Green
Write-Host "Project root: $projectRoot"
Write-Host "Python: $pythonVersion ($pythonPath)"
Write-Host "Microsoft Project: $projectVersion"
if ($CheckOnly) {
    Write-Host "CheckOnly mode: configuration was not changed"
} else {
    Write-Host "Created .codex/config.toml and .mcp.json"
    Write-Host "Restart Codex or launch Claude Code from the project root"
}
