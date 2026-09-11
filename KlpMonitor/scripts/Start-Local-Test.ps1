param(
    [string]$Token = 'computer-monitor-local-test-2026'
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$agentPath = Join-Path $projectRoot 'target\release\computer-agent.exe'
$dashboardPath = Join-Path $projectRoot 'target\release\computer-dashboard.exe'

if (-not (Test-Path -LiteralPath $agentPath -PathType Leaf) -or -not (Test-Path -LiteralPath $dashboardPath -PathType Leaf)) {
    & (Join-Path $PSScriptRoot 'Build-Release.ps1')
}

$agent = Start-Process -FilePath $agentPath -ArgumentList @('--listen', '127.0.0.1:47824', '--token', $Token) -WindowStyle Hidden -PassThru
try {
    Start-Sleep -Milliseconds 900
    & $dashboardPath --connect '127.0.0.1:47824' --token $Token
} finally {
    if (-not $agent.HasExited) {
        Stop-Process -Id $agent.Id -Force
        $agent.WaitForExit()
    }
}
