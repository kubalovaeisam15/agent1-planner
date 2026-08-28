param(
    [string]$Version = "0.6.0",
    [string]$OutputDirectory = "packages"
)

$ErrorActionPreference = "Stop"
$projectRoot = [System.IO.Path]::GetFullPath((Split-Path -Parent $PSScriptRoot))
$packageDirectory = [System.IO.Path]::GetFullPath((Join-Path $projectRoot $OutputDirectory))
[System.IO.Directory]::CreateDirectory($packageDirectory) | Out-Null

$archiveName = "MS-Project-Agent-$Version.zip"
$archivePath = Join-Path $packageDirectory $archiveName
$checksumPath = "$archivePath.sha256.txt"
if ([System.IO.File]::Exists($archivePath) -or [System.IO.File]::Exists($checksumPath)) {
    throw "A package with this version already exists: $archivePath"
}

$stagePath = [System.IO.Path]::GetFullPath((Join-Path $packageDirectory (".stage-" + [guid]::NewGuid().ToString("N"))))
if (-not ([System.IO.Path]::GetDirectoryName($stagePath)).Equals(
    $packageDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
    throw "Unsafe staging path: $stagePath"
}

try {
    [System.IO.Directory]::CreateDirectory($stagePath) | Out-Null
    foreach ($directoryName in @("instructions", "data", "tools", "tests", "docs", "setup", ".agents", ".codex")) {
        [System.IO.Directory]::CreateDirectory((Join-Path $stagePath $directoryName)) | Out-Null
    }

    Get-ChildItem -LiteralPath $projectRoot -Filter "*.md" -File | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $stagePath
    }
    Copy-Item -Path (Join-Path $projectRoot "instructions\*") `
        -Destination (Join-Path $stagePath "instructions") -Recurse

    $mppTemplates = @(Get-ChildItem -LiteralPath (Join-Path $projectRoot "data") -Filter "*.mpp" -File)
    if ($mppTemplates.Count -ne 1) {
        throw "Expected exactly one corporate MPP template; found $($mppTemplates.Count)"
    }
    Copy-Item -LiteralPath $mppTemplates[0].FullName -Destination (Join-Path $stagePath "data")

    foreach ($directoryName in @("tools", "tests", "docs", "setup")) {
        $sourceDirectory = Join-Path $projectRoot $directoryName
        Get-ChildItem -LiteralPath $sourceDirectory -Recurse -File |
            Where-Object { $_.FullName -notmatch '[\\/]__pycache__[\\/]' } |
            ForEach-Object {
                $relativePath = $_.FullName.Substring($sourceDirectory.Length).TrimStart('\', '/')
                $destination = Join-Path (Join-Path $stagePath $directoryName) $relativePath
                [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destination)) | Out-Null
                Copy-Item -LiteralPath $_.FullName -Destination $destination
            }
    }

    $skillsSource = Join-Path $projectRoot ".agents"
    Get-ChildItem -LiteralPath $skillsSource -Recurse -File |
        Where-Object { $_.FullName -notmatch '[\\/]__pycache__[\\/]' } |
        ForEach-Object {
            $relativePath = $_.FullName.Substring($skillsSource.Length).TrimStart('\', '/')
            $destination = Join-Path (Join-Path $stagePath ".agents") $relativePath
            [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($destination)) | Out-Null
            Copy-Item -LiteralPath $_.FullName -Destination $destination
        }

    Copy-Item -LiteralPath (Join-Path $projectRoot ".codex/hooks.template.json") `
        -Destination (Join-Path $stagePath ".codex")
    [System.IO.Directory]::CreateDirectory((Join-Path $stagePath ".codex/hooks")) | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot ".codex/hooks/agent1_hook.py") `
        -Destination (Join-Path $stagePath ".codex/hooks")

    $items = @(Get-ChildItem -LiteralPath $stagePath | ForEach-Object FullName)
    Compress-Archive -LiteralPath $items -DestinationPath $archivePath -CompressionLevel Optimal
}
finally {
    if ([System.IO.Directory]::Exists($stagePath)) {
        $resolvedStage = [System.IO.Path]::GetFullPath($stagePath)
        if (-not ([System.IO.Path]::GetDirectoryName($resolvedStage)).Equals(
            $packageDirectory, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Refusing to remove unsafe staging path: $resolvedStage"
        }
        [System.IO.Directory]::Delete($resolvedStage, $true)
    }
}

$hash = Get-FileHash -Algorithm SHA256 -LiteralPath $archivePath
[System.IO.File]::WriteAllText(
    $checksumPath,
    "$($hash.Hash)  $archiveName`r`n",
    [System.Text.ASCIIEncoding]::new()
)

Write-Host "Created package: $archivePath" -ForegroundColor Green
Write-Host "SHA256: $($hash.Hash)"
