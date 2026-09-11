param(
    [Parameter(Mandatory = $true)]
    [ValidateLength(12, 256)]
    [string]$Token,
    [string]$Listen = '0.0.0.0:47821',
    [switch]$Hidden
)

$ErrorActionPreference = 'Stop'
$agentPath = Join-Path $PSScriptRoot 'Computer.Agent.exe'
if (-not (Test-Path -LiteralPath $agentPath -PathType Leaf)) {
    $agentPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'target\release\computer-agent.exe'
}
if (-not (Test-Path -LiteralPath $agentPath -PathType Leaf)) {
    throw "Computer.Agent.exe was not found: $agentPath"
}

if ($Hidden) {
    Start-Process -FilePath $agentPath -ArgumentList @('--listen', $Listen, '--token', $Token, '--interval-ms', '500') -WindowStyle Hidden
} else {
    & $agentPath --listen $Listen --token $Token --interval-ms 500
}
