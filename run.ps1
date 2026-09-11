#Requires -Version 5.1

param(
    [switch]$Automatic,
    [string]$BasePath = 'D:\',
    [switch]$PlanOnly,
    [switch]$SkipApplications,
    [switch]$ElevatedChild
)

$ErrorActionPreference = 'Stop'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

trap {
    Write-Host ''
    Write-Host "KlpInstall: $($_.Exception.Message)" -ForegroundColor Red
    if ($ElevatedChild) {
        Read-Host 'Press Enter to close this administrator window' | Out-Null
    }
    exit 1
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-PowerShellLiteral {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
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
    # Start-Process joins ArgumentList into one Windows command line. Building quoted
    # path arguments by hand breaks values such as D:\ because the trailing slash
    # sits next to a closing quote. Encode the PowerShell command instead so paths,
    # spaces, apostrophes and trailing backslashes survive UAC elevation unchanged.
    $commandParts = @(
        '&',
        (ConvertTo-PowerShellLiteral $PSCommandPath),
        '-BasePath',
        (ConvertTo-PowerShellLiteral $base),
        '-ElevatedChild'
    )
    if ($Automatic) { $commandParts += '-Automatic' }
    if ($SkipApplications) { $commandParts += '-SkipApplications' }

    $elevatedCommand = $commandParts -join ' '
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($elevatedCommand))
    $elevationArguments = @(
        '-NoLogo',
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-EncodedCommand',
        $encodedCommand
    )

    $process = Start-Process -FilePath 'powershell.exe' -Verb RunAs -ArgumentList $elevationArguments -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        Write-Host ''
        Write-Host "KlpInstall failed in the administrator window with exit code $($process.ExitCode)." -ForegroundColor Red
    }
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
$exitCode = $LASTEXITCODE

if ($ElevatedChild -and $exitCode -ne 0) {
    Write-Host ''
    Write-Host "KlpInstall failed with exit code $exitCode." -ForegroundColor Red
    Read-Host 'Press Enter to close this administrator window' | Out-Null
}

exit $exitCode
