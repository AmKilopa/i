param(
    [Parameter(Mandatory = $true)]
    [string]$Computer,
    [Parameter(Mandatory = $true)]
    [ValidateLength(12, 256)]
    [string]$Token,
    [switch]$Maximized
)

$ErrorActionPreference = 'Stop'
$dashboardPath = Join-Path $PSScriptRoot 'Klpfetch.exe'
if (-not (Test-Path -LiteralPath $dashboardPath -PathType Leaf)) {
    $dashboardPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'target\release\computer-dashboard.exe'
}
if (-not (Test-Path -LiteralPath $dashboardPath -PathType Leaf)) {
    throw "Klpfetch.exe was not found: $dashboardPath"
}

$terminal = Get-Command 'wt.exe' -ErrorAction SilentlyContinue
if ($terminal) {
    $mode = if ($Maximized) { '-M' } else { '-F' }
    Start-Process -FilePath $terminal.Source -ArgumentList @(
        $mode,
        '-w',
        'new',
        'new-tab',
        '--title',
        'Klpfetch',
        $dashboardPath,
        '--connect',
        "$Computer`:47821",
        '--token',
        $Token
    )
} else {
    & $dashboardPath --connect "$Computer`:47821" --token $Token
}
