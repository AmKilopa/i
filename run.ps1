#Requires -Version 5.1

param(
    [switch]$Automatic,
    [string]$BasePath = 'D:\',
    [switch]$PlanOnly,
    [switch]$SkipApplications
)

$ErrorActionPreference = 'Stop'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

trap {
    Write-Host ''
    Write-Host "KlpInstall: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-QuotedArgument {
    param([string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Invoke-DownloadFile {
    param([string]$Uri, [string]$Destination)
    $partial = "$Destination.partial"
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            $previousProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $Uri -OutFile $partial -UseBasicParsing -TimeoutSec 180
            $ProgressPreference = $previousProgress
            if (-not (Test-Path -LiteralPath $partial -PathType Leaf) -or (Get-Item -LiteralPath $partial).Length -eq 0) { throw 'The downloaded file is empty' }
            Move-Item -LiteralPath $partial -Destination $Destination -Force
            return
        } catch {
            $ProgressPreference = $previousProgress
            Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            if ($attempt -eq 3) { throw }
            Start-Sleep -Seconds ($attempt * 2)
        }
    }
}

if ([string]::IsNullOrWhiteSpace($PSCommandPath)) { throw 'run.ps1 must be started as a file' }
if (-not [System.IO.Path]::IsPathRooted($BasePath)) { throw 'BasePath must be an absolute path' }
$base = [System.IO.Path]::GetFullPath($BasePath)
$dataRoot = [System.IO.Path]::GetPathRoot($base)
$systemRoot = [System.IO.Path]::GetPathRoot($env:SystemRoot)
if ($dataRoot.TrimEnd('\') -ieq $systemRoot.TrimEnd('\')) { throw "The target cannot be the system drive $systemRoot" }
if (-not (Test-Path -LiteralPath $dataRoot -PathType Container)) { throw "Drive $dataRoot is unavailable" }

if (-not $PlanOnly -and -not (Test-IsAdministrator)) {
    $elevationArguments = @(
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (ConvertTo-QuotedArgument $PSCommandPath),
        '-BasePath',
        (ConvertTo-QuotedArgument $base)
    )
    if ($Automatic) { $elevationArguments += '-Automatic' }
    if ($SkipApplications) { $elevationArguments += '-SkipApplications' }
    $process = Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList ($elevationArguments -join ' ') -Wait -PassThru
    exit $process.ExitCode
}

$source = Join-Path $PSScriptRoot 'i.ps1'
$downloadUri = 'https://raw.githubusercontent.com/AmKilopa/i/main/i.ps1'
$bootstrapRoot = Join-Path $dataRoot 'KlpInstall\Bootstrap'
if ($PlanOnly -and (Test-Path -LiteralPath $source -PathType Leaf)) {
    $target = $source
} else {
    $target = Join-Path $bootstrapRoot 'i.ps1'
    New-Item -ItemType Directory -Path $bootstrapRoot -Force | Out-Null
    if (Test-Path -LiteralPath $source -PathType Leaf) {
        if ([System.IO.Path]::GetFullPath($source) -ine [System.IO.Path]::GetFullPath($target)) {
            Copy-Item -LiteralPath $source -Destination $target -Force
        }
    } else {
        Invoke-DownloadFile -Uri $downloadUri -Destination $target
    }
}

if (-not (Test-Path -LiteralPath $target -PathType Leaf) -or (Get-Item -LiteralPath $target).Length -eq 0) { throw 'The main installer could not be prepared' }

$arguments = @('-NoLogo', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $target, '-BasePath', $base)
if ($Automatic) { $arguments += @('-Silent', '-SkipConfirm') }
if ($PlanOnly) { $arguments += '-PlanOnly' }
if ($SkipApplications) { $arguments += '-SkipApplications' }
& powershell.exe @arguments
exit $LASTEXITCODE
