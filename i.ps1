#Requires -RunAsAdministrator
$ErrorActionPreference = 'SilentlyContinue'

Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

function Write-Ok($m)   { Write-Host '  [' -NoNewline; Write-Host 'OK' -ForegroundColor Green -NoNewline; Write-Host "] $m" }
function Write-Info($m) { Write-Host '  [' -NoNewline; Write-Host 'INFO' -ForegroundColor Cyan -NoNewline; Write-Host "] $m" }
function Write-Warn($m) { Write-Host '  [' -NoNewline; Write-Host 'WARN' -ForegroundColor Yellow -NoNewline; Write-Host "] $m" }

function Refresh-Path {
    $m = [System.Environment]::GetEnvironmentVariable('PATH','Machine')
    $u = [System.Environment]::GetEnvironmentVariable('PATH','User')
    $env:PATH = "$m;$u"
}

function Download-File($Url, $Dest, $Label) {
    $wc = $null
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add('User-Agent','Mozilla/5.0')
        $wc.DownloadFile($Url, $Dest)
        Write-Ok "$Label"
    } catch {
        Write-Warn "$Label : $_"
    } finally {
        if ($wc) { $wc.Dispose() }
    }
}

function Install-Winget($id, $label) {
    Write-Host "  $label" -ForegroundColor Gray
    winget install --id $id -e --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) { Write-Ok $label } else { Write-Warn "$label (Код: $LASTEXITCODE)" }
}

$boxW = 50
$boxTop = "  +" + ("=" * $boxW) + "+"
$boxLn = "  |" + ("  Setup: D: drive only, no C: clutter".PadRight($boxW)) + "|"
Clear-Host
Write-Host ''
Write-Host $boxTop -ForegroundColor Cyan
Write-Host $boxLn -ForegroundColor Cyan
Write-Host $boxTop -ForegroundColor Cyan
Write-Host ''

Write-Host '  Base [Enter = D:\]: ' -ForegroundColor Yellow -NoNewline
$inputPath = Read-Host
if ([string]::IsNullOrWhiteSpace($inputPath)) { $inputPath = 'D:\' }
$BASE = $inputPath.TrimEnd('\')
Write-Info "Base: $BASE"
Write-Host ''

Write-Host '  Folders...' -ForegroundColor Cyan
foreach ($f in @('Download','Project','Program','Games','Discord','Telegram')) {
    $p = "$BASE\$f"
    if (-not (Test-Path $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
        Write-Host '  ' -NoNewline; Write-Host '+' -ForegroundColor Green -NoNewline; Write-Host " $p"
    } else {
        Write-Host '  ' -NoNewline; Write-Host '~' -ForegroundColor DarkGray -NoNewline; Write-Host " $p"
    }
}
Write-Host ''

Write-Host '  winget check...' -ForegroundColor Cyan
if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Warn 'winget not found'
    $r = Read-Host '  Continue? (y/n)'
    if ($r -ne 'y') { exit 1 }
} else {
    Write-Ok 'winget'
    winget source update 2>&1 | Out-Null
}
Write-Host ''

Write-Host '  tg-ws-proxy...' -ForegroundColor Cyan
$tgProxyDir = "$BASE\Telegram"
try {
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/Flowseal/tg-ws-proxy/releases/latest"
    if ($release.assets.Count -gt 0) {
        $asset = $release.assets | Where-Object { $_.name -match 'windows|win' -or $_.name -match '\.exe$' } | Select-Object -First 1
        
        if ($asset) {
            $proxyUrl = $asset.browser_download_url
            $proxyDest = "$tgProxyDir\$($asset.name)"
            if (-not (Test-Path $proxyDest)) {
                Download-File -Url $proxyUrl -Dest $proxyDest -Label "tg-ws-proxy ($($asset.name))"
            } else {
                Write-Ok 'tg-ws-proxy already downloaded'
            }
        } else {
            Write-Warn 'tg-ws-proxy: Windows asset not found in release'
        }
    } else {
        Write-Warn 'tg-ws-proxy: No assets found in the latest release'
    }
} catch {
    Write-Warn "tg-ws-proxy fetch failed: $_"
}
Write-Host ''

Write-Host '  Telegram...' -ForegroundColor Cyan
if (Get-Command winget -ErrorAction SilentlyContinue) {
    Install-Winget 'Telegram.TelegramDesktop' 'Telegram'
}
Write-Host ''

Write-Host '  happ...' -ForegroundColor Cyan
Write-Info 'Starting happ download/installation...'

do {
    $userResponse = Read-Host '  [WAITING] Type "yes" to continue with the remaining installations'
} while ($userResponse.Trim().ToLower() -ne 'yes')

Write-Ok 'Confirmation received, continuing...'
Write-Host ''

$chocoPath = "$BASE\Chocolatey"
[System.Environment]::SetEnvironmentVariable('ChocolateyInstall', $chocoPath, 'Machine')
$env:ChocolateyInstall = $chocoPath

Write-Host '  Chocolatey...' -ForegroundColor Cyan
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    if (Test-Path "$chocoPath\bin\choco.exe") {
        $env:Path = "$chocoPath\bin;$env:Path"
        Write-Ok 'Chocolatey (Recovered)'
    } else {
        if (Test-Path $chocoPath) { Remove-Item -Path $chocoPath -Recurse -Force -ErrorAction SilentlyContinue }
        try {
            Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
            Refresh-Path
            $env:Path = "$chocoPath\bin;$env:Path"
            if (Get-Command choco -ErrorAction SilentlyContinue) { Write-Ok 'Chocolatey' } else { Write-Warn 'Chocolatey' }
        } catch { Write-Warn "Chocolatey: $_" }
    }
} else {
    Write-Ok 'Chocolatey'
}
Write-Host ''

Write-Host '  Programs...' -ForegroundColor Cyan
if (Get-Command winget -ErrorAction SilentlyContinue) {
    Install-Winget 'Python.Python.3.12' 'Python'
    Install-Winget 'Opera.OperaGX' 'Opera GX'
    Install-Winget 'OpenJS.NodeJS.LTS' 'Node.js'
    Install-Winget 'Microsoft.VisualStudioCode' 'VS Code'
    Install-Winget 'Discord.Discord' 'Discord'
    Install-Winget 'Git.Git' 'Git'
}
Write-Host ''

Write-Host '  Rust...' -ForegroundColor Cyan
$rustDir = "$BASE\Rust"
$rustupPath = "$rustDir\.rustup"
$cargoPath = "$rustDir\.cargo"
$rustupExe = "$BASE\Download\rustup-init.exe"

if (Get-Command rustc -ErrorAction SilentlyContinue) {
    Write-Ok 'Rust'
} else {
    if (-not (Test-Path $rustDir)) { New-Item -ItemType Directory -Path $rustDir -Force | Out-Null }
    if (-not (Test-Path "$BASE\Download")) { New-Item -ItemType Directory -Path "$BASE\Download" -Force | Out-Null }
    if (-not (Test-Path $rustupExe)) {
        Download-File -Url 'https://static.rust-lang.org/rustup/dist/x86_64-pc-windows-msvc/rustup-init.exe' -Dest $rustupExe -Label 'rustup-init'
    }
    if (Test-Path $rustupExe) {
        $env:RUSTUP_HOME = $rustupPath
        $env:CARGO_HOME = $cargoPath
        [System.Environment]::SetEnvironmentVariable('RUSTUP_HOME', $rustupPath, 'Machine')
        [System.Environment]::SetEnvironmentVariable('CARGO_HOME', $cargoPath, 'Machine')
        [System.Environment]::SetEnvironmentVariable('PATH', "$cargoPath\bin;" + [System.Environment]::GetEnvironmentVariable('PATH','Machine'), 'Machine')
        try {
            Start-Process -FilePath $rustupExe -ArgumentList '-y' -Wait -NoNewWindow
            Refresh-Path
            $env:PATH = "$cargoPath\bin;$env:PATH"
            if (Get-Command rustc -ErrorAction SilentlyContinue) { Write-Ok 'Rust' } else { Write-Warn 'Rust' }
        } catch { Write-Warn "Rust: $_" }
    }
}
Write-Host ''

Write-Host '  Spotify...' -ForegroundColor Cyan
if (Get-Command choco -ErrorAction SilentlyContinue) {
    choco install spotify -y 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'Spotify' } else { Write-Warn 'Spotify' }
} else { Write-Warn 'Spotify' }
Write-Host ''

Write-Host '  npm...' -ForegroundColor Cyan
Refresh-Path
if (Get-Command npm -ErrorAction SilentlyContinue) {
    npm install -g pnpm 2>&1 | Out-Null
    if (Get-Command pnpm -ErrorAction SilentlyContinue) { Write-Ok 'pnpm' } else { Write-Warn 'pnpm' }
    npm install -g klpgit 2>&1 | Out-Null
    if (Get-Command klpgit -ErrorAction SilentlyContinue) { Write-Ok 'klpgit' } else { Write-Warn 'klpgit' }
} else { Write-Warn 'npm' }
Write-Host ''

Write-Host $boxTop -ForegroundColor Green
Write-Host ("  |" + ("  Done.".PadRight($boxW)) + "|") -ForegroundColor Green
Write-Host $boxTop -ForegroundColor Green
Write-Host ''
Read-Host '  Enter to exit'
