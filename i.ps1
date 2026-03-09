#Requires -RunAsAdministrator
$ErrorActionPreference = 'SilentlyContinue'

function Write-Ok($m)   { Write-Host '  [' -NoNewline; Write-Host 'OK'    -ForegroundColor Green  -NoNewline; Write-Host "] $m" }
function Write-Info($m) { Write-Host '  [' -NoNewline; Write-Host 'INFO'  -ForegroundColor Cyan   -NoNewline; Write-Host "] $m" }
function Write-Warn($m) { Write-Host '  [' -NoNewline; Write-Host 'WARN'  -ForegroundColor Yellow -NoNewline; Write-Host "] $m" }
function Write-Err($m)  { Write-Host '  [' -NoNewline; Write-Host 'ERROR' -ForegroundColor Red    -NoNewline; Write-Host "] $m" }

function Refresh-Path {
    $m = [System.Environment]::GetEnvironmentVariable('PATH','Machine')
    $u = [System.Environment]::GetEnvironmentVariable('PATH','User')
    $env:PATH = "$m;$u"
}

function Download-File($Url, $Dest, $Label) {
    Write-Host ''
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add('User-Agent','Mozilla/5.0')
    Register-ObjectEvent -InputObject $wc -EventName DownloadProgressChanged -SourceIdentifier 'DLP' -Action {
        $p = $Event.SourceArgs[1].ProgressPercentage
        $r = [math]::Round($Event.SourceArgs[1].BytesReceived/1MB,2)
        $t = [math]::Round($Event.SourceArgs[1].TotalBytesToReceive/1MB,2)
        $b = '#' * [math]::Floor($p/5)
        $s = ' ' * (20 - [math]::Floor($p/5))
        [Console]::Write("`r [$b$s] $p% ($r MB / $t MB)  ")
    } | Out-Null
    try {
        $task = $wc.DownloadFileTaskAsync($Url, $Dest)
        while (-not $task.IsCompleted) { Start-Sleep -Milliseconds 150 }
        Write-Host ''
        if ($task.IsFaulted) { throw $task.Exception.InnerException }
        Write-Ok "$Label - downloaded!"
    } catch {
        Write-Host ''
        Write-Err "Failed to download $Label : $_"
    } finally {
        Unregister-Event -SourceIdentifier 'DLP' -ErrorAction SilentlyContinue
        $wc.Dispose()
    }
}

function Install-Winget($id, $label) {
    Write-Host "  Installing $label ..." -ForegroundColor Gray
    winget install --id $id -e --silent --accept-package-agreements --accept-source-agreements 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
        Write-Ok "$label - done"
    } else {
        Write-Warn "$label - already installed or error (code $LASTEXITCODE)"
    }
    Write-Host ''
}

Clear-Host
Write-Host ''
Write-Host '  +==================================================+' -ForegroundColor Cyan
Write-Host '  |       Installing programs on clean Windows       |' -ForegroundColor Cyan
Write-Host '  +==================================================+' -ForegroundColor Cyan
Write-Host ''

Write-Host '  Enter base folder path [Enter = D:\]: ' -ForegroundColor Yellow -NoNewline
$inputPath = Read-Host
if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = 'D:\' }
$BASE = $inputPath.TrimEnd('\')
Write-Host ''
Write-Info "Base folder: $BASE"
Write-Host ''
Write-Host '  --------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

Write-Host '  [FOLDERS] Creating folder structure...' -ForegroundColor Cyan
Write-Host ''
foreach ($f in @('Download','Project','Program','Games')) {
    $p = "$BASE\$f"
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        Write-Host '  [' -NoNewline; Write-Host '+' -ForegroundColor Green -NoNewline; Write-Host "] Created: $p"
    } else {
        Write-Host '  [' -NoNewline; Write-Host '~' -ForegroundColor DarkYellow -NoNewline; Write-Host "] Already exists: $p"
    }
}
Write-Host ''
Write-Host '  --------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

Write-Host '  [CHOCO] Checking Chocolatey...' -ForegroundColor Cyan
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-Info 'Installing Chocolatey...'
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    Refresh-Path
    if (Get-Command choco -ErrorAction SilentlyContinue) { Write-Ok 'Chocolatey installed!' }
    else { Write-Warn 'Chocolatey could not be installed. Continuing...' }
} else {
    $v = choco --version 2>$null
    Write-Ok "Chocolatey already installed. Version: $v"
}
Write-Host ''
Write-Host '  --------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

Write-Host '  [ZAPRET] Downloading zapret-discord-youtube...' -ForegroundColor Cyan
$zDir  = 'C:\Discord'
$zFile = 'C:\Discord\zapret-discord-youtube-1.9.7b.rar'
$zUrl  = 'https://github.com/Flowseal/zapret-discord-youtube/releases/download/1.9.7b/zapret-discord-youtube-1.9.7b.rar'
if (-not (Test-Path $zDir)) {
    New-Item -ItemType Directory -Path $zDir -Force | Out-Null
    Write-Ok 'Folder C:\Discord created'
} else {
    Write-Host '  [~] C:\Discord already exists' -ForegroundColor DarkYellow
}
if (Test-Path $zFile) {
    Write-Ok "Zapret already downloaded: $zFile"
} else {
    Write-Info "Source: $zUrl"
    Download-File -Url $zUrl -Dest $zFile -Label 'Zapret'
}
Write-Host ''
Write-Host '  --------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Err 'winget not found. Install App Installer from Microsoft Store.'
    Read-Host 'Press Enter to exit'
    exit 1
}
Write-Info 'Starting program installation...'
Write-Host ''
Install-Winget 'Python.Python.3.12'              'Python 3.12        [1/7]'
Install-Winget 'Opera.OperaGX'                   'Opera GX           [2/7]'
Install-Winget 'OpenJS.NodeJS.LTS'               'Node.js 22 LTS     [3/7]'
Install-Winget 'EclipseAdoptium.Temurin.21.JDK'  'Java JDK 21        [4/7]'
Install-Winget 'Anysphere.Cursor'                'Cursor             [5/7]'
Install-Winget 'Discord.Discord'                 'Discord            [6/7]'
Install-Winget 'Git.Git'                         'Git                [7/7]'

Write-Info 'Refreshing PATH...'
Refresh-Path
Write-Ok 'PATH refreshed'
Write-Host ''

Write-Host '  [npm] Installing pnpm...' -ForegroundColor Cyan
if (Get-Command npm -ErrorAction SilentlyContinue) {
    npm install -g pnpm 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'pnpm installed!' } else { Write-Warn "pnpm error (code $LASTEXITCODE)" }
} else {
    Write-Warn 'npm not found. After reboot run: npm install -g pnpm'
}
Write-Host ''

Write-Host '  [npm] Installing klpgit...' -ForegroundColor Cyan
if (Get-Command npm -ErrorAction SilentlyContinue) {
    npm install -g klpgit 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'klpgit installed!' } else { Write-Warn "klpgit error (code $LASTEXITCODE)" }
} else {
    Write-Warn 'npm not found. After reboot run: npm install -g klpgit'
}
Write-Host ''
Write-Host '  --------------------------------------------------' -ForegroundColor DarkGray
Write-Host ''

Write-Host '  +==================================================+' -ForegroundColor Green
Write-Host '  |              Installation complete!              |' -ForegroundColor Green
Write-Host '  +==================================================+' -ForegroundColor Green
Write-Host ''
Write-Host "  Folders in $BASE :" -ForegroundColor White
Write-Host '    Download / Project / Program / Games' -ForegroundColor Gray
Write-Host ''
Write-Host '  Zapret: C:\Discord\zapret-discord-youtube-1.9.7b.rar' -ForegroundColor Gray
Write-Host ''
Write-Host '  Installed:' -ForegroundColor White
foreach ($item in @('Chocolatey','Python 3.12','Opera GX','Node.js 22 LTS','Java JDK 21','Cursor','Discord','Git','pnpm','klpgit')) {
    Write-Host "    v $item" -ForegroundColor Green
}
Write-Host ''
Read-Host '  Press Enter to exit'
