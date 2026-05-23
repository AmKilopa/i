#Requires -RunAsAdministrator
#Requires -Version 5.1

param(
    [string]$BasePath,
    [switch]$Silent,
    [switch]$SkipConfirm
)

$ErrorActionPreference = 'Stop'
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

# ── Config ──────────────────────────────────────────────────────────────────

$SCRIPT_VERSION = '2.0.0'
$MIN_DISK_SPACE_GB = 10
$DOWNLOAD_RETRIES = 3
$DOWNLOAD_TIMEOUT = 120

$FOLDERS = @('Download', 'Project', 'Program', 'Games', 'Discord', 'Telegram')

$WINGET_APPS = [ordered]@{
    'Python'    = 'Python.Python.3.12'
    'Opera GX'  = 'Opera.OperaGX'
    'Node.js'   = 'OpenJS.NodeJS.LTS'
    'VS Code'   = 'Microsoft.VisualStudioCode'
    'Discord'   = 'Discord.Discord'
    'Git'       = 'Git.Git'
    'Telegram'  = 'Telegram.TelegramDesktop'
}

$NPM_PACKAGES = @('pnpm')

$CHOCO_PACKAGES = @('spotify')

# ── State tracking ──────────────────────────────────────────────────────────

$Results = [ordered]@{}
$LogFile = $null
$StartTime = Get-Date

# ── UI Functions ────────────────────────────────────────────────────────────

function Write-Ok($m) {
    $msg = "  [ OK ] $m"
    Write-Host '  [' -NoNewline; Write-Host ' OK ' -ForegroundColor Green -NoNewline; Write-Host "] $m"
    Add-Log $msg
}

function Write-Info($m) {
    $msg = "  [INFO] $m"
    Write-Host '  [' -NoNewline; Write-Host 'INFO' -ForegroundColor Cyan -NoNewline; Write-Host "] $m"
    Add-Log $msg
}

function Write-Warn($m) {
    $msg = "  [WARN] $m"
    Write-Host '  [' -NoNewline; Write-Host 'WARN' -ForegroundColor Yellow -NoNewline; Write-Host "] $m"
    Add-Log $msg
}

function Write-Fail($m) {
    $msg = "  [FAIL] $m"
    Write-Host '  [' -NoNewline; Write-Host 'FAIL' -ForegroundColor Red -NoNewline; Write-Host "] $m"
    Add-Log $msg
}

function Write-Section($m) {
    Write-Host ''
    Write-Host "  -- $m --" -ForegroundColor Cyan
    Add-Log "-- $m --"
}

function Write-Banner {
    $w = 52
    $line = "+" + ("=" * $w) + "+"
    Clear-Host
    Write-Host ''
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ("  |" + "  Windows Setup Script v$SCRIPT_VERSION".PadRight($w) + "|") -ForegroundColor Cyan
    Write-Host ("  |" + "  Clean install, custom base path".PadRight($w) + "|") -ForegroundColor Cyan
    Write-Host ("  |" + "  $(Get-Date -Format 'yyyy-MM-dd HH:mm')".PadRight($w) + "|") -ForegroundColor Cyan
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ''
}

function Write-ProgressStep($current, $total, $label) {
    if ($total -le 0) { return }
    $pct = [math]::Round(($current / $total) * 100)
    $barLen = 30
    $filled = [math]::Round($barLen * $current / $total)
    $empty = $barLen - $filled
    $bar = ("#" * $filled) + ("-" * $empty)
    Write-Host "`r  [$bar] $pct% - $label    " -NoNewline -ForegroundColor Gray
}

function Show-Report {
    $w = 52
    $line = "+" + ("-" * $w) + "+"
    $elapsed = (Get-Date) - $StartTime

    Write-Host ''
    Write-Host ''
    Write-Host "  $line" -ForegroundColor White
    Write-Host ("  |" + "  INSTALLATION REPORT".PadRight($w) + "|") -ForegroundColor White
    Write-Host "  $line" -ForegroundColor White

    $ok = 0; $fail = 0; $skip = 0
    foreach ($key in $Results.Keys) {
        $status = $Results[$key]
        $padW = $w - 2
        switch ($status) {
            'OK' {
                $text = "  [+] $key"
                Write-Host ("  |" + $text.PadRight($padW) + "|") -ForegroundColor Green
                $ok++
            }
            'SKIP' {
                $text = "  [-] $key (skipped)"
                Write-Host ("  |" + $text.PadRight($padW) + "|") -ForegroundColor DarkGray
                $skip++
            }
            'EXISTS' {
                $text = "  [~] $key (already installed)"
                Write-Host ("  |" + $text.PadRight($padW) + "|") -ForegroundColor DarkGray
                $ok++
            }
            default {
                $text = "  [X] $key : $status"
                if ($text.Length -gt $padW) { $text = $text.Substring(0, $padW) }
                Write-Host ("  |" + $text.PadRight($padW) + "|") -ForegroundColor Red
                $fail++
            }
        }
    }

    Write-Host "  $line" -ForegroundColor White
    $summaryText = "  OK: $ok  Failed: $fail  Skipped: $skip"
    Write-Host ("  |" + $summaryText.PadRight($w) + "|") -ForegroundColor White
    $timeText = "  Time: $($elapsed.ToString('mm\:ss'))"
    Write-Host ("  |" + $timeText.PadRight($w) + "|") -ForegroundColor White
    if ($LogFile) {
        $logText = "  Log: $LogFile"
        if ($logText.Length -gt $w) { $logText = $logText.Substring(0, $w) }
        Write-Host ("  |" + $logText.PadRight($w) + "|") -ForegroundColor White
    }
    Write-Host "  $line" -ForegroundColor White
    Write-Host ''
}

# ── Utility Functions ───────────────────────────────────────────────────────

function Add-Log($m) {
    if ($LogFile) {
        $ts = Get-Date -Format 'HH:mm:ss'
        "[$ts] $m" | Out-File -FilePath $LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

function Refresh-Path {
    $m = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $u = [System.Environment]::GetEnvironmentVariable('PATH', 'User')
    $env:PATH = "$m;$u"
}

function Add-ToSystemPath($dir) {
    $current = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine')
    if ($current -notlike "*$dir*") {
        [System.Environment]::SetEnvironmentVariable('PATH', "$dir;$current", 'Machine')
        $env:PATH = "$dir;$env:PATH"
        return $true
    }
    return $false
}

function Backup-PathVariable {
    $pathBackup = [System.Environment]::GetEnvironmentVariable('PATH', 'Machine')
    $backupFile = "$BASE\Download\PATH_BACKUP_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
    $pathBackup | Out-File -FilePath $backupFile -Encoding UTF8
    Write-Info "PATH backed up to $backupFile"
}

function Test-Internet {
    try {
        $r = Invoke-WebRequest -Uri 'https://www.google.com' -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop
        return $r.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Test-DiskSpace($path, $minGB) {
    try {
        $drive = $path.Substring(0, 1)
        $disk = Get-PSDrive -Name $drive -ErrorAction Stop
        $freeGB = [math]::Round($disk.Free / 1GB, 2)
        return @{ OK = ($freeGB -ge $minGB); FreeGB = $freeGB }
    } catch {
        return @{ OK = $true; FreeGB = -1 }
    }
}

function Download-File {
    param(
        [string]$Url,
        [string]$Dest,
        [string]$Label,
        [int]$Retries = $DOWNLOAD_RETRIES
    )

    for ($i = 1; $i -le $Retries; $i++) {
        $wc = $null
        try {
            $wc = New-Object System.Net.WebClient
            $wc.Headers.Add('User-Agent', 'Mozilla/5.0')
            $wc.DownloadFile($Url, $Dest)

            if ((Test-Path $Dest) -and (Get-Item $Dest).Length -gt 0) {
                Write-Ok "$Label"
                return $true
            } else {
                throw "Downloaded file is empty"
            }
        } catch {
            if ($i -lt $Retries) {
                Write-Warn "$Label attempt $i/$Retries failed: $_"
                Start-Sleep -Seconds 3
            } else {
                Write-Fail "$Label after $Retries attempts: $_"
                return $false
            }
        } finally {
            if ($wc) { $wc.Dispose() }
        }
    }
    return $false
}

function Install-WingetPackage($id, $label) {
    try {
        $check = winget list --id $id --exact --source winget 2>&1
        if ($LASTEXITCODE -eq 0 -and $check -match $id) {
            Write-Ok "$label (already installed)"
            return 'EXISTS'
        }
    } catch {}

    Write-Host "    Installing $label..." -ForegroundColor Gray
    try {
        $output = winget install --id $id --exact --source winget --silent `
            --accept-package-agreements --accept-source-agreements 2>&1

        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            Write-Ok $label
            return 'OK'
        } else {
            Write-Fail "$label (winget exit: $LASTEXITCODE)"
            return "winget exit $LASTEXITCODE"
        }
    } catch {
        Write-Fail "$label : $_"
        return "$_"
    }
}

function Prompt-YesNo($question, $default = 'y') {
    if ($SkipConfirm) { return $true }
    $hint = if ($default -eq 'y') { '(Y/n)' } else { '(y/N)' }
    $r = Read-Host "  $question $hint"
    if ([string]::IsNullOrWhiteSpace($r)) { $r = $default }
    return $r.Trim().ToLower() -eq 'y'
}

function Select-Components {
    Write-Section 'Component Selection'

    $components = [ordered]@{}

    Write-Host '  Choose what to install:' -ForegroundColor Yellow
    Write-Host ''

    foreach ($key in $WINGET_APPS.Keys) {
        $components[$key] = Prompt-YesNo "Install ${key}?"
    }

    Write-Host ''

    $components['Chocolatey'] = Prompt-YesNo 'Install Chocolatey?'
    $components['Spotify']    = Prompt-YesNo 'Install Spotify (via Chocolatey)?'
    $components['Rust']       = Prompt-YesNo 'Install Rust?'
    $components['npm_tools']  = Prompt-YesNo 'Install npm global tools (pnpm)?'
    $components['tg_proxy']   = Prompt-YesNo 'Download tg-ws-proxy?'

    Write-Host ''
    Write-Host '  Selected:' -ForegroundColor Cyan
    foreach ($key in $components.Keys) {
        if ($components[$key]) {
            Write-Host "    [+] $key" -ForegroundColor Green
        } else {
            Write-Host "    [-] $key" -ForegroundColor DarkGray
        }
    }
    Write-Host ''

    return $components
}

# ── Main ────────────────────────────────────────────────────────────────────

Write-Banner

# ── Internet check ──────────────────────────────────────────────────────────

Write-Section 'Pre-flight Checks'
if (-not (Test-Internet)) {
    Write-Fail 'No internet connection detected'
    Read-Host '  Enter to exit'
    exit 1
}
Write-Ok 'Internet connection'

# ── Base path ───────────────────────────────────────────────────────────────

if (-not $BasePath) {
    Write-Host '  Base path [Enter = D:\]: ' -ForegroundColor Yellow -NoNewline
    $inputPath = Read-Host
    if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = 'D:\' }
    $BASE = $inputPath.TrimEnd('\')
} else {
    $BASE = $BasePath.TrimEnd('\')
}

$driveLetter = $BASE.Substring(0, 1)
if (-not (Test-Path "${driveLetter}:\")) {
    Write-Fail "Drive ${driveLetter}: does not exist"
    Read-Host '  Enter to exit'
    exit 1
}
Write-Ok "Base path: $BASE"

# ── Disk space ──────────────────────────────────────────────────────────────

$diskCheck = Test-DiskSpace $BASE $MIN_DISK_SPACE_GB
if ($diskCheck.FreeGB -ge 0) {
    if ($diskCheck.OK) {
        Write-Ok "Disk space: $($diskCheck.FreeGB) GB free"
    } else {
        Write-Warn "Low disk space: $($diskCheck.FreeGB) GB (need $MIN_DISK_SPACE_GB GB)"
        if (-not (Prompt-YesNo 'Continue anyway?')) { exit 1 }
    }
}

# ── Winget check ────────────────────────────────────────────────────────────

$hasWinget = [bool](Get-Command winget -ErrorAction SilentlyContinue)
if ($hasWinget) {
    Write-Ok 'winget available'
    winget source update 2>&1 | Out-Null
} else {
    Write-Warn 'winget not found - winget-based installs will be skipped'
}

# ── Component selection ─────────────────────────────────────────────────────

$components = Select-Components

# ── Create folders ──────────────────────────────────────────────────────────

Write-Section 'Folder Structure'
foreach ($f in $FOLDERS) {
    $p = "$BASE\$f"
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        Write-Host '    ' -NoNewline; Write-Host '+' -ForegroundColor Green -NoNewline; Write-Host " $p"
    } else {
        Write-Host '    ' -NoNewline; Write-Host '~' -ForegroundColor DarkGray -NoNewline; Write-Host " $p (exists)"
    }
}

# ── Init log ────────────────────────────────────────────────────────────────

$LogFile = "$BASE\Download\setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Add-Log "Setup started - Base: $BASE"
Add-Log "Script version: $SCRIPT_VERSION"
Add-Log "OS: $([System.Environment]::OSVersion.VersionString)"

# ── Backup PATH ─────────────────────────────────────────────────────────────

Backup-PathVariable

# ── Count steps ─────────────────────────────────────────────────────────────

$totalSteps = ($components.Values | Where-Object { $_ }).Count
$currentStep = 0

# ── tg-ws-proxy ─────────────────────────────────────────────────────────────

if ($components['tg_proxy']) {
    $currentStep++
    Write-ProgressStep $currentStep $totalSteps 'tg-ws-proxy'
    Write-Section 'tg-ws-proxy'

    $tgProxyDir = "$BASE\Telegram"
    try {
        $release = Invoke-RestMethod -Uri "https://api.github.com/repos/Flowseal/tg-ws-proxy/releases/latest" -TimeoutSec 15
        $asset = $release.assets | Where-Object {
            $_.name -match 'windows|win' -or $_.name -match '\.exe$'
        } | Select-Object -First 1

        if ($asset) {
            $proxyDest = "$tgProxyDir\$($asset.name)"
            if (-not (Test-Path $proxyDest)) {
                $dl = Download-File -Url $asset.browser_download_url -Dest $proxyDest -Label "tg-ws-proxy ($($asset.name))"
                $Results['tg-ws-proxy'] = if ($dl) { 'OK' } else { 'Download failed' }
            } else {
                Write-Ok 'tg-ws-proxy (already downloaded)'
                $Results['tg-ws-proxy'] = 'EXISTS'
            }
        } else {
            Write-Warn 'tg-ws-proxy: no Windows binary in release'
            $Results['tg-ws-proxy'] = 'No Windows asset'
        }
    } catch {
        Write-Fail "tg-ws-proxy: $_"
        $Results['tg-ws-proxy'] = "$_"
    }
} else {
    $Results['tg-ws-proxy'] = 'SKIP'
}

# ── Winget apps ─────────────────────────────────────────────────────────────

if ($hasWinget) {
    Write-Section 'Winget Applications'
    foreach ($appName in $WINGET_APPS.Keys) {
        if ($components[$appName]) {
            $currentStep++
            Write-ProgressStep $currentStep $totalSteps $appName
            Write-Host ''
            $Results[$appName] = Install-WingetPackage $WINGET_APPS[$appName] $appName
        } else {
            $Results[$appName] = 'SKIP'
        }
    }
} else {
    foreach ($appName in $WINGET_APPS.Keys) {
        if ($components[$appName]) {
            $Results[$appName] = 'SKIP (no winget)'
        } else {
            $Results[$appName] = 'SKIP'
        }
    }
}

Refresh-Path

# ── Chocolatey ──────────────────────────────────────────────────────────────

if ($components['Chocolatey']) {
    $currentStep++
    Write-ProgressStep $currentStep $totalSteps 'Chocolatey'
    Write-Section 'Chocolatey'

    $chocoPath = "$BASE\Chocolatey"
    [System.Environment]::SetEnvironmentVariable('ChocolateyInstall', $chocoPath, 'Machine')
    $env:ChocolateyInstall = $chocoPath

    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Ok 'Chocolatey (already installed)'
        $Results['Chocolatey'] = 'EXISTS'
    } elseif (Test-Path "$chocoPath\bin\choco.exe") {
        $env:Path = "$chocoPath\bin;$env:Path"
        Write-Ok 'Chocolatey (recovered from path)'
        $Results['Chocolatey'] = 'OK'
    } else {
        try {
            if (Test-Path $chocoPath) { Remove-Item -Path $chocoPath -Recurse -Force -ErrorAction SilentlyContinue }
            $installScript = (New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1')
            Invoke-Expression $installScript
            Refresh-Path
            $env:Path = "$chocoPath\bin;$env:Path"

            if (Get-Command choco -ErrorAction SilentlyContinue) {
                Write-Ok 'Chocolatey'
                $Results['Chocolatey'] = 'OK'
            } else {
                Write-Fail 'Chocolatey (not found after install)'
                $Results['Chocolatey'] = 'Post-install check failed'
            }
        } catch {
            Write-Fail "Chocolatey: $_"
            $Results['Chocolatey'] = "$_"
        }
    }
} else {
    $Results['Chocolatey'] = 'SKIP'
}

# ── Spotify ─────────────────────────────────────────────────────────────────

if ($components['Spotify']) {
    $currentStep++
    Write-ProgressStep $currentStep $totalSteps 'Spotify'
    Write-Section 'Spotify'

    if (Get-Command choco -ErrorAction SilentlyContinue) {
        try {
            choco install spotify -y 2>&1 | Out-Null
            if ($LASTEXITCODE -eq 0) {
                Write-Ok 'Spotify'
                $Results['Spotify'] = 'OK'
            } else {
                Write-Fail "Spotify (choco exit: $LASTEXITCODE)"
                $Results['Spotify'] = "choco exit $LASTEXITCODE"
            }
        } catch {
            Write-Fail "Spotify: $_"
            $Results['Spotify'] = "$_"
        }
    } else {
        Write-Warn 'Spotify: Chocolatey not available'
        $Results['Spotify'] = 'No Chocolatey'
    }
} else {
    $Results['Spotify'] = 'SKIP'
}

# ── Rust ────────────────────────────────────────────────────────────────────

if ($components['Rust']) {
    $currentStep++
    Write-ProgressStep $currentStep $totalSteps 'Rust'
    Write-Section 'Rust'

    $rustDir = "$BASE\Rust"
    $rustupHome = "$rustDir\.rustup"
    $cargoHome = "$rustDir\.cargo"
    $rustupExe = "$BASE\Download\rustup-init.exe"

    if (Get-Command rustc -ErrorAction SilentlyContinue) {
        $ver = (rustc --version 2>&1) -replace 'rustc ',''
        Write-Ok "Rust ($ver)"
        $Results['Rust'] = 'EXISTS'
    } else {
        try {
            if (-not (Test-Path $rustDir)) { New-Item -ItemType Directory -Path $rustDir -Force | Out-Null }

            [System.Environment]::SetEnvironmentVariable('RUSTUP_HOME', $rustupHome, 'Machine')
            [System.Environment]::SetEnvironmentVariable('CARGO_HOME', $cargoHome, 'Machine')
            $env:RUSTUP_HOME = $rustupHome
            $env:CARGO_HOME = $cargoHome

            if (-not (Test-Path $rustupExe)) {
                $dl = Download-File `
                    -Url 'https://static.rust-lang.org/rustup/dist/x86_64-pc-windows-msvc/rustup-init.exe' `
                    -Dest $rustupExe `
                    -Label 'rustup-init'
                if (-not $dl) { throw "Failed to download rustup-init.exe" }
            }

            Start-Process -FilePath $rustupExe -ArgumentList '-y','--no-modify-path' -Wait -NoNewWindow

            Add-ToSystemPath "$cargoHome\bin" | Out-Null
            Refresh-Path
            $env:PATH = "$cargoHome\bin;$env:PATH"

            if (Get-Command rustc -ErrorAction SilentlyContinue) {
                $ver = (rustc --version 2>&1) -replace 'rustc ',''
                Write-Ok "Rust ($ver)"
                $Results['Rust'] = 'OK'
            } else {
                Write-Fail 'Rust (not found after install)'
                $Results['Rust'] = 'Post-install check failed'
            }
        } catch {
            Write-Fail "Rust: $_"
            $Results['Rust'] = "$_"
        }
    }
} else {
    $Results['Rust'] = 'SKIP'
}

# ── npm tools ───────────────────────────────────────────────────────────────

if ($components['npm_tools']) {
    $currentStep++
    Write-ProgressStep $currentStep $totalSteps 'npm tools'
    Write-Section 'npm Global Packages'

    Refresh-Path

    if (Get-Command npm -ErrorAction SilentlyContinue) {
        $npmVer = (npm --version 2>&1)
        Write-Info "npm $npmVer"

        foreach ($pkg in $NPM_PACKAGES) {
            try {
                npm install -g $pkg 2>&1 | Out-Null
                if (Get-Command $pkg -ErrorAction SilentlyContinue) {
                    $pver = & $pkg --version 2>&1
                    Write-Ok "$pkg ($pver)"
                    $Results["npm:$pkg"] = 'OK'
                } else {
                    Write-Warn "$pkg (installed but not in PATH)"
                    $Results["npm:$pkg"] = 'Not in PATH'
                }
            } catch {
                Write-Fail "$pkg : $_"
                $Results["npm:$pkg"] = "$_"
            }
        }
    } else {
        Write-Warn 'npm not found - install Node.js first'
        foreach ($pkg in $NPM_PACKAGES) { $Results["npm:$pkg"] = 'No npm' }
    }
} else {
    foreach ($pkg in $NPM_PACKAGES) { $Results["npm:$pkg"] = 'SKIP' }
}

# ── Cleanup ─────────────────────────────────────────────────────────────────

Write-Section 'Cleanup'
$tempFiles = @("$BASE\Download\rustup-init.exe")
foreach ($tf in $tempFiles) {
    if (Test-Path $tf) {
        if (Prompt-YesNo "Delete temp file $(Split-Path $tf -Leaf)?") {
            Remove-Item $tf -Force -ErrorAction SilentlyContinue
            Write-Ok "Deleted $tf"
        }
    }
}

# ── Report ──────────────────────────────────────────────────────────────────

Write-Host ''
Write-ProgressStep $totalSteps $totalSteps 'Done'
Write-Host ''

Show-Report

Add-Log "Setup completed"
Add-Log "Results: $($Results | Out-String)"

Read-Host '  Enter to exit'
