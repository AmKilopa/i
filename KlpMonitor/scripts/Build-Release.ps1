$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$distRoot = Join-Path $projectRoot 'dist'
$pcRoot = Join-Path $distRoot 'PC'
$laptopRoot = Join-Path $distRoot 'Laptop'

Push-Location $projectRoot
try {
    cargo test --all-targets
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed' }
    cargo build --release
    if ($LASTEXITCODE -ne 0) { throw 'Release build failed' }
    New-Item -ItemType Directory -Path $pcRoot, $laptopRoot -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $projectRoot 'target\release\computer-agent.exe') -Destination (Join-Path $pcRoot 'Computer.Agent.exe') -Force
    Copy-Item -LiteralPath (Join-Path $projectRoot 'target\release\computer-dashboard.exe') -Destination (Join-Path $laptopRoot 'Klpfetch.exe') -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Start-Agent-Lan.ps1') -Destination $pcRoot -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Start-Dashboard-Fullscreen.ps1') -Destination $laptopRoot -Force
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Generate-Token.ps1') -Destination $distRoot -Force
    Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination $distRoot -Force
} finally {
    Pop-Location
}

Write-Host "Ready: $distRoot" -ForegroundColor Green
