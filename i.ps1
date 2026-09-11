#Requires -Version 5.1

param(
    [string]$BasePath = 'D:\',
    [switch]$Silent,
    [switch]$SkipConfirm,
    [switch]$PlanOnly,
    [switch]$SkipApplications
)

$ErrorActionPreference = 'Stop'
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

$ScriptVersion = '4.2.0'
$MinimumFreeSpaceGB = 50
$Results = [ordered]@{}
$LogFile = $null
$StartTime = Get-Date
$RebootRequired = $false
$WingetCommand = $null

$AppSpecs = [ordered]@{
    'Python' = @{ Id = 'Python.Python.3.12'; Scope = 'machine'; Folder = 'Python312'; Executables = @('python.exe'); Installer = 'Python'; PathFolders = @('', 'Scripts') }
    'Firefox' = @{ Id = 'Mozilla.Firefox.ru'; Scope = 'machine'; Folder = 'Firefox'; Executables = @('firefox.exe'); Installer = 'Firefox'; PathFolders = @() }
    'Node.js' = @{ Id = 'OpenJS.NodeJS.LTS'; Scope = 'machine'; Folder = 'NodeJS'; Executables = @('node.exe'); Installer = 'Node'; PathFolders = @('') }
    'VS Code' = @{ Id = 'Microsoft.VisualStudioCode'; Scope = 'machine'; Folder = 'VSCode'; Executables = @('Code.exe'); Installer = 'VSCode'; PathFolders = @('bin') }
    'Discord' = @{ Id = 'Discord.Discord'; Scope = 'user'; Folder = 'Discord'; Executables = @('Update.exe', 'Discord.exe'); Installer = 'Discord'; PathFolders = @() }
    'Git' = @{ Id = 'Git.Git'; Scope = 'machine'; Folder = 'Git'; Executables = @('git.exe'); Installer = 'Git'; PathFolders = @('cmd', 'bin') }
    'Telegram' = @{ Id = 'Telegram.TelegramDesktop'; Scope = 'user'; Folder = 'Telegram'; Executables = @('Telegram.exe'); Installer = 'Telegram'; PathFolders = @('') }
}

function Add-Log {
    param([string]$Message)
    if ($LogFile) {
        "[$(Get-Date -Format 'HH:mm:ss')] $Message" | Out-File -LiteralPath $LogFile -Append -Encoding UTF8 -ErrorAction SilentlyContinue
    }
}

function Write-State {
    param([string]$Type, [string]$Message, [ConsoleColor]$Color)
    Write-Host '  [' -NoNewline
    Write-Host $Type.PadRight(4) -ForegroundColor $Color -NoNewline
    Write-Host "] $Message"
    Add-Log "[$Type] $Message"
}

function Write-Ok { param([string]$Message) Write-State 'OK' $Message Green }
function Write-Info { param([string]$Message) Write-State 'INFO' $Message Cyan }
function Write-Warn { param([string]$Message) Write-State 'WARN' $Message Yellow }
function Write-Fail { param([string]$Message) Write-State 'FAIL' $Message Red }
function Write-Section { param([string]$Message) Write-Host ''; Write-Host "  -- $Message --" -ForegroundColor Cyan; Add-Log "-- $Message --" }

function Write-Banner {
    $width = 56
    $line = '+' + ('=' * $width) + '+'
    try { Clear-Host } catch {}
    Write-Host ''
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ('  |' + "  Windows D: Setup v$ScriptVersion".PadRight($width) + '|') -ForegroundColor Cyan
    Write-Host ('  |' + '  Applications, downloads and caches on D:'.PadRight($width) + '|') -ForegroundColor Cyan
    Write-Host ('  |' + "  $(Get-Date -Format 'yyyy-MM-dd HH:mm')".PadRight($width) + '|') -ForegroundColor Cyan
    Write-Host "  $line" -ForegroundColor Cyan
    Write-Host ''
}

function Prompt-YesNo {
    param([string]$Question, [bool]$Default = $true)
    if ($SkipConfirm -or $Silent) { return $Default }
    $hint = if ($Default) { '(Y/n)' } else { '(y/N)' }
    $answer = Read-Host "  $Question $hint"
    if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
    return $answer.Trim().ToLowerInvariant() -eq 'y'
}

function Get-FullBasePath {
    param([string]$Path)
    if (-not [System.IO.Path]::IsPathRooted($Path)) { throw 'BasePath must be an absolute path on a non-system drive' }
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $root = [System.IO.Path]::GetPathRoot($fullPath)
    $systemRoot = [System.IO.Path]::GetPathRoot($env:SystemRoot)
    if ($root.TrimEnd('\') -ieq $systemRoot.TrimEnd('\')) { throw "BasePath cannot be on the system drive $systemRoot" }
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "Drive $root is unavailable" }
    if ($fullPath.TrimEnd('\') -ieq $root.TrimEnd('\')) { return $root }
    return $fullPath.TrimEnd('\')
}

function Get-FreeSpaceGB {
    param([string]$Path)
    $root = [System.IO.Path]::GetPathRoot($Path)
    $drive = Get-PSDrive -Name $root.Substring(0, 1) -ErrorAction Stop
    return [math]::Round($drive.Free / 1GB, 2)
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Assert-SupportedSystem {
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) { throw 'Windows is required' }
    $version = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
    $build = [int]$version.CurrentBuildNumber
    if ($build -lt 17763) { throw "Windows build $build is not supported. Windows 10 1809 or newer is required." }
    if (-not [Environment]::Is64BitOperatingSystem) { throw 'A 64-bit version of Windows is required' }
    if (-not $PlanOnly -and -not (Test-IsAdministrator)) { throw 'Run install.bat or run.ps1 and approve the administrator prompt' }
}

function Get-DataVolume {
    param([string]$Path)
    $root = [System.IO.Path]::GetPathRoot($Path)
    $deviceId = $root.TrimEnd('\')
    $volume = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$deviceId'" -ErrorAction Stop
    if (-not $volume) { throw "Drive $root was not found" }
    if ([int]$volume.DriveType -ne 3) { throw "Drive $root must be a fixed local drive" }
    if ($volume.FileSystem -ne 'NTFS') { throw "Drive $root must use NTFS. Current filesystem: $($volume.FileSystem)" }
    return $volume
}

function Test-Internet {
    foreach ($uri in @('https://cdn.winget.microsoft.com/cache/source.msix', 'https://api.github.com')) {
        try {
            $response = Invoke-WebRequest -Uri $uri -Method Head -UseBasicParsing -TimeoutSec 20
            if ($response.StatusCode -ge 200 -and $response.StatusCode -lt 400) { return $true }
        } catch {}
    }
    return $false
}

function Invoke-DownloadFile {
    param([string]$Uri, [string]$Destination, [int]$TimeoutSec = 180)
    $parent = Split-Path -Parent $Destination
    if ($parent) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $partial = "$Destination.partial"
    $displayName = [System.IO.Path]::GetFileName($Destination)
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        $previousProgress = $ProgressPreference
        try {
            Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            Write-Info "Downloading $displayName (attempt $attempt/3)..."
            $ProgressPreference = 'Continue'
            Invoke-WebRequest -Uri $Uri -OutFile $partial -UseBasicParsing -TimeoutSec $TimeoutSec
            $ProgressPreference = $previousProgress
            if (-not (Test-Path -LiteralPath $partial -PathType Leaf) -or (Get-Item -LiteralPath $partial).Length -eq 0) { throw 'The downloaded file is empty' }
            Move-Item -LiteralPath $partial -Destination $Destination -Force
            $sizeMB = [math]::Round((Get-Item -LiteralPath $Destination).Length / 1MB, 1)
            Write-Ok "Downloaded $displayName ($sizeMB MB)"
            return
        } catch {
            $ProgressPreference = $previousProgress
            Remove-Item -LiteralPath $partial -Force -ErrorAction SilentlyContinue
            Add-Log "Download attempt $attempt failed for ${Uri}: $($_.Exception.Message)"
            if ($attempt -eq 3) { throw }
            Write-Warn "Download failed; retrying in $($attempt * 2) seconds..."
            Start-Sleep -Seconds ($attempt * 2)
        }
    }
}

function Set-PersistentEnvironment {
    param([string]$Name, [string]$Value, [ValidateSet('User', 'Machine')][string]$Target = 'User')
    [System.Environment]::SetEnvironmentVariable($Name, $Value, $Target)
    Set-Item -Path "Env:$Name" -Value $Value
    Write-Ok "$Name -> $Value"
}

function Export-EnvironmentState {
    $userValues = [ordered]@{}
    $machineValues = [ordered]@{}
    foreach ($name in $StorageEnvironment.Keys) { $userValues[$name] = [Environment]::GetEnvironmentVariable($name, 'User') }
    foreach ($name in $MachineEnvironment.Keys) { $machineValues[$name] = [Environment]::GetEnvironmentVariable($name, 'Machine') }
    $state = [ordered]@{
        Timestamp = (Get-Date).ToString('o')
        User = $userValues
        Machine = $machineValues
        UserPath = [Environment]::GetEnvironmentVariable('Path', 'User')
        MachinePath = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    }
    $path = Join-Path $Layout.Backups "environment_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $state | ConvertTo-Json -Depth 5 | Out-File -LiteralPath $path -Encoding UTF8
    Write-Info "Environment backup -> $path"
}

function Initialize-KnownFolderApi {
    if ('KlpKnownFolder' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class KlpKnownFolder
{
    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHGetKnownFolderPath(ref Guid id, uint flags, IntPtr token, out IntPtr path);

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern int SHSetKnownFolderPath(ref Guid id, uint flags, IntPtr token, string path);

    public static string Get(string value)
    {
        Guid id = new Guid(value);
        IntPtr pointer;
        int result = SHGetKnownFolderPath(ref id, 0, IntPtr.Zero, out pointer);
        if (result != 0) Marshal.ThrowExceptionForHR(result);
        try { return Marshal.PtrToStringUni(pointer); }
        finally { Marshal.FreeCoTaskMem(pointer); }
    }

    public static void Set(string value, string path)
    {
        Guid id = new Guid(value);
        int result = SHSetKnownFolderPath(ref id, 0x2000, IntPtr.Zero, path);
        if (result != 0) Marshal.ThrowExceptionForHR(result);
    }
}
'@
}

function Export-RegistryKey {
    param([string]$Key, [string]$Destination)
    & reg.exe query $Key 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) { return $false }
    & reg.exe export $Key $Destination /y | Out-Null
    return $LASTEXITCODE -eq 0
}

function Set-KnownFolderLayout {
    Initialize-KnownFolderApi
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $registryBackup = Join-Path $Layout.Backups "user-shell-folders_$stamp.reg"
    if (-not (Export-RegistryKey 'HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' $registryBackup)) { throw 'Could not back up current known-folder settings' }
    $before = @()
    $failures = @()
    $changed = $false
    foreach ($spec in $KnownFolderSpecs) {
        try {
            $current = [KlpKnownFolder]::Get($spec.Id)
            $target = $Layout[$spec.LayoutKey]
            $before += [pscustomobject]@{ Name = $spec.Name; Id = $spec.Id; Source = $current; Target = $target }
            New-Item -ItemType Directory -Path $target -Force | Out-Null
            Write-Info "Path only: $($spec.Name) -> $target (existing files are not moved)"
            if ($current.TrimEnd('\') -ine $target.TrimEnd('\')) {
                # Path-only mode: existing files remain untouched.
                [KlpKnownFolder]::Set($spec.Id, $target)
                $changed = $true
            }
            $actual = [KlpKnownFolder]::Get($spec.Id)
            if ($actual.TrimEnd('\') -ine $target.TrimEnd('\')) { throw "Windows returned $actual" }
            Write-Ok "$($spec.Name) -> $target"
        } catch {
            $failures += $spec.Name
            Write-Fail "$($spec.Name): $($_.Exception.Message)"
        }
    }
    $mappingPath = Join-Path $Layout.Backups "known-folders_$stamp.json"
    $before | ConvertTo-Json -Depth 4 | Out-File -LiteralPath $mappingPath -Encoding UTF8
    Write-Info "Known-folder backup -> $mappingPath"
    if ($failures.Count -gt 0) { return "FAILED: $($failures -join ', ')" }
    if ($changed) { $script:RebootRequired = $true }
    return 'OK'
}

function Refresh-Path {
    $machinePath = [System.Environment]::GetEnvironmentVariable('Path', 'Machine')
    $userPath = [System.Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = "$machinePath;$userPath"
}

function Add-ToPath {
    param([string]$Directory, [ValidateSet('User', 'Machine')][string]$Target = 'Machine')
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return }
    $current = [System.Environment]::GetEnvironmentVariable('Path', $Target)
    $parts = @($current -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if (-not ($parts | Where-Object { $_.TrimEnd('\') -ieq $Directory.TrimEnd('\') })) {
        [System.Environment]::SetEnvironmentVariable('Path', (@($Directory) + $parts) -join ';', $Target)
        Write-Ok "PATH -> $Directory"
    }
}

function Test-AppAtTarget {
    param([hashtable]$Spec, [string]$Target)
    if (-not (Test-Path -LiteralPath $Target -PathType Container)) { return $false }
    foreach ($name in $Spec.Executables) {
        if (Get-ChildItem -LiteralPath $Target -Recurse -File -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1) { return $true }
    }
    return $false
}

function Test-WingetPackageInstalled {
    param([string]$Id)
    if (-not $WingetCommand) { return $false }
    $output = & $WingetCommand list --id $Id --exact --source winget --accept-source-agreements --disable-interactivity 2>&1
    $exitCode = $LASTEXITCODE
    return $exitCode -eq 0 -and (($output -join "`n") -match [regex]::Escape($Id))
}

function Resolve-WingetCommand {
    Refresh-Path
    $command = Get-Command winget.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { return $command.Source }
    $aliasPath = Join-Path $env:LOCALAPPDATA 'Microsoft\WindowsApps\winget.exe'
    if (Test-Path -LiteralPath $aliasPath -PathType Leaf) { return $aliasPath }
    $package = Get-AppxPackage -Name 'Microsoft.DesktopAppInstaller' -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1
    if ($package) {
        $packageCommand = Join-Path $package.InstallLocation 'winget.exe'
        if (Test-Path -LiteralPath $packageCommand -PathType Leaf) { return $packageCommand }
    }
    return $null
}

function Ensure-Winget {
    $script:WingetCommand = Resolve-WingetCommand
    if (-not $script:WingetCommand) {
        try {
            $addAppx = Get-Command Add-AppxPackage -ErrorAction Stop
            if ($addAppx.Parameters.ContainsKey('RegisterByFamilyName')) {
                Add-AppxPackage -RegisterByFamilyName -MainPackage 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe' -ErrorAction Stop
            }
        } catch { Add-Log "WinGet registration: $($_.Exception.Message)" }
        $script:WingetCommand = Resolve-WingetCommand
    }
    if (-not $script:WingetCommand) {
        try {
            $oldProgress = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            Install-PackageProvider -Name NuGet -Force -Scope AllUsers -Confirm:$false | Out-Null
            if (-not (Get-Module -ListAvailable -Name Microsoft.WinGet.Client)) {
                Install-Module -Name Microsoft.WinGet.Client -Repository PSGallery -Scope AllUsers -Force -AllowClobber -Confirm:$false
            }
            Import-Module Microsoft.WinGet.Client -Force
            $repair = Get-Command Repair-WinGetPackageManager -ErrorAction Stop
            $repairArguments = @{ AllUsers = $true }
            if ($repair.Parameters.ContainsKey('Force')) { $repairArguments.Force = $true }
            if ($repair.Parameters.ContainsKey('Latest')) { $repairArguments.Latest = $true }
            & $repair @repairArguments | Out-Null
            $ProgressPreference = $oldProgress
        } catch {
            $ProgressPreference = $oldProgress
            Add-Log "WinGet repair: $($_.Exception.Message)"
        }
        $script:WingetCommand = Resolve-WingetCommand
    }
    if (-not $script:WingetCommand) {
        try {
            $bundleDirectory = Join-Path $Layout.Installers 'WinGet'
            $bundlePath = Join-Path $bundleDirectory 'Microsoft.DesktopAppInstaller.msixbundle'
            New-Item -ItemType Directory -Path $bundleDirectory -Force | Out-Null
            Invoke-DownloadFile -Uri 'https://aka.ms/getwinget' -Destination $bundlePath
            Add-AppxPackage -Path $bundlePath -ErrorAction Stop
        } catch { Add-Log "WinGet App Installer: $($_.Exception.Message)" }
        $script:WingetCommand = Resolve-WingetCommand
    }
    if (-not $script:WingetCommand) { throw 'WinGet could not be installed automatically' }
    $sourceOutput = & $script:WingetCommand source update --disable-interactivity 2>&1
    $sourceOutput | ForEach-Object { Add-Log "winget source: $_" }
    Write-Ok "WinGet $(& $script:WingetCommand --version)"
    return 'OK'
}

function Ensure-Junction {
    param([string]$Link, [string]$Target)
    New-Item -ItemType Directory -Path $Target -Force | Out-Null
    New-Item -ItemType Directory -Path (Split-Path -Parent $Link) -Force | Out-Null
    if (Test-Path -LiteralPath $Link) {
        $item = Get-Item -LiteralPath $Link -Force
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
            $currentTarget = @($item.Target)[0]
            if ($currentTarget -and ([System.IO.Path]::GetFullPath($currentTarget).TrimEnd('\') -ieq [System.IO.Path]::GetFullPath($Target).TrimEnd('\'))) {
                Write-Ok "$Link -> $Target"
                return $true
            }
            Write-Warn "Existing link points elsewhere: $Link"
            return $false
        }
        if ($item.PSIsContainer -and -not (Get-ChildItem -LiteralPath $Link -Force | Select-Object -First 1)) {
            Remove-Item -LiteralPath $Link -Force
        } else {
            Write-Warn "Existing data was not moved: $Link"
            return $false
        }
    }
    New-Item -ItemType Junction -Path $Link -Target $Target | Out-Null
    Write-Ok "$Link -> $Target"
    return $true
}

function Get-AppJunctions {
    param([string]$Name, [string]$ProgramTarget)
    switch ($Name) {
        'Firefox' { return @(@{ Link = Join-Path $env:APPDATA 'Mozilla'; Target = Join-Path $Layout.AppData 'Firefox\Roaming' }, @{ Link = Join-Path $env:LOCALAPPDATA 'Mozilla'; Target = Join-Path $Layout.AppData 'Firefox\Local' }) }
        'VS Code' { return @(@{ Link = Join-Path $env:APPDATA 'Code'; Target = Join-Path $Layout.AppData 'VSCode\Roaming' }, @{ Link = Join-Path $env:USERPROFILE '.vscode'; Target = Join-Path $Layout.AppData 'VSCode\Profile' }) }
        'Discord' { return @(@{ Link = Join-Path $env:LOCALAPPDATA 'Discord'; Target = $ProgramTarget }, @{ Link = Join-Path $env:APPDATA 'discord'; Target = Join-Path $Layout.AppData 'Discord\Roaming' }) }
        'Telegram' { return @(@{ Link = Join-Path $env:APPDATA 'Telegram Desktop'; Target = Join-Path $Layout.AppData 'Telegram' }) }
        'Spotify' { return @(@{ Link = Join-Path $env:APPDATA 'Spotify'; Target = $ProgramTarget }, @{ Link = Join-Path $env:LOCALAPPDATA 'Spotify'; Target = Join-Path $Layout.AppData 'Spotify\Local' }) }
        default { return @() }
    }
}

function Initialize-AppStorage {
    param([string]$Name, [string]$ProgramTarget)
    foreach ($junction in @(Get-AppJunctions -Name $Name -ProgramTarget $ProgramTarget)) {
        if (-not (Ensure-Junction -Link $junction.Link -Target $junction.Target)) { return $false }
    }
    return $true
}

function Download-WingetInstaller {
    param([string]$Id, [string]$Scope, [string]$Destination)
    if (-not $WingetCommand) { return $null }
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    $selectors = @(
        @{ Scope = $true; Architecture = $true },
        @{ Scope = $false; Architecture = $true },
        @{ Scope = $true; Architecture = $false },
        @{ Scope = $false; Architecture = $false }
    )
    $exitCode = -1
    foreach ($selector in $selectors) {
        $arguments = @('download', '--id', $Id, '--exact', '--source', 'winget')
        if ($selector.Scope) { $arguments += @('--scope', $Scope) }
        if ($selector.Architecture) { $arguments += @('--architecture', $Architecture) }
        $arguments += @('--download-directory', $Destination, '--accept-package-agreements', '--accept-source-agreements', '--disable-interactivity')
        Write-Info "WinGet download: $Id"
        & $WingetCommand @arguments 2>&1 | ForEach-Object {
            Write-Host $_
            Add-Log "winget: $_"
        }
        $exitCode = $LASTEXITCODE
        Add-Log "winget download $Id exit code: $exitCode"
        if ($exitCode -eq 0) { break }
        Write-Warn "WinGet download attempt failed (exit $exitCode); trying a compatible selector..."
    }
    if ($exitCode -ne 0) {
        Write-Fail "Download failed for $Id (winget exit $exitCode)"
        return $null
    }
    $installer = Get-ChildItem -LiteralPath $Destination -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in @('.exe', '.msi', '.msix', '.msixbundle') } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $installer) {
        Write-Fail "Installer file not found for $Id"
        return $null
    }
    Write-Ok "Downloaded to $($installer.FullName)"
    return $installer.FullName
}

function Quote-Argument { param([string]$Value) return '"' + $Value.Replace('"', '\"') + '"' }

function Set-FirefoxConfiguration {
    param([string]$Target)
    $distribution = Join-Path $Target 'distribution'
    $policyPath = Join-Path $distribution 'policies.json'
    $policy = [ordered]@{
        policies = [ordered]@{
            DisableFirefoxStudies = $true
            DisablePocket = $true
            DisableTelemetry = $true
            DefaultDownloadDirectory = $Layout.Downloads
            DownloadDirectory = $Layout.Downloads
            FirefoxHome = [ordered]@{
                Search = $true
                TopSites = $true
                SponsoredTopSites = $false
                Highlights = $false
                Pocket = $false
                Stories = $false
                SponsoredPocket = $false
                SponsoredStories = $false
                Snippets = $false
            }
            NoDefaultBookmarks = $true
            UserMessaging = [ordered]@{
                ExtensionRecommendations = $false
                FeatureRecommendations = $false
                MoreFromMozilla = $false
                SkipOnboarding = $true
                UrlbarInterventions = $false
                WhatsNew = $false
            }
        }
    }
    New-Item -ItemType Directory -Path $distribution -Force | Out-Null
    $policy | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $policyPath -Encoding UTF8
    Write-Ok "Firefox clean configuration -> $policyPath"
}

function Set-BrowserDownloadPolicies {
    Get-Process -Name msedge, chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Export-RegistryKey 'HKLM\SOFTWARE\Policies\Microsoft\Edge' (Join-Path $Layout.Backups "edge-policies_$stamp.reg") | Out-Null
    Export-RegistryKey 'HKLM\SOFTWARE\Policies\Google\Chrome' (Join-Path $Layout.Backups "chrome-policies_$stamp.reg") | Out-Null
    $browsers = @(
        @{ Name = 'Edge'; RegistryPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Edge'; Source = Join-Path $env:LOCALAPPDATA 'Microsoft\Edge\User Data'; Profile = $Layout.EdgeProfile; Cache = $Layout.EdgeCache; PolicyProfile = $BrowserPolicyPaths.EdgeProfile; PolicyCache = $BrowserPolicyPaths.EdgeCache },
        @{ Name = 'Chrome'; RegistryPath = 'HKLM:\SOFTWARE\Policies\Google\Chrome'; Source = Join-Path $env:LOCALAPPDATA 'Google\Chrome\User Data'; Profile = $Layout.ChromeProfile; Cache = $Layout.ChromeCache; PolicyProfile = $BrowserPolicyPaths.ChromeProfile; PolicyCache = $BrowserPolicyPaths.ChromeCache }
    )
    $mapping = @($browsers | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Source = $_.Source; Target = $_.Profile; Cache = $_.Cache } })
    $mapping | ConvertTo-Json -Depth 4 | Out-File -LiteralPath (Join-Path $Layout.Backups "browser-storage_$stamp.json") -Encoding UTF8
    foreach ($browser in $browsers) {
                # Path-only mode: existing files remain untouched.
        New-Item -Path $browser.RegistryPath -Force | Out-Null
        New-ItemProperty -Path $browser.RegistryPath -Name 'DownloadDirectory' -Value $Layout.Downloads -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $browser.RegistryPath -Name 'UserDataDir' -Value $browser.PolicyProfile -PropertyType String -Force | Out-Null
        New-ItemProperty -Path $browser.RegistryPath -Name 'DiskCacheDir' -Value $browser.PolicyCache -PropertyType String -Force | Out-Null
    }
    Write-Ok "Browser downloads -> $($Layout.Downloads)"
    Write-Ok 'Edge and Chrome profiles and caches -> D:'
    return 'OK'
}

function Disable-OneDrive {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $backups = @(
        @{ Key = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\OneDrive'; Name = "onedrive-machine_$stamp.reg" },
        @{ Key = 'HKLM\SOFTWARE\Policies\Microsoft\OneDrive'; Name = "onedrive-machine-sync_$stamp.reg" },
        @{ Key = 'HKCU\SOFTWARE\Policies\Microsoft\OneDrive'; Name = "onedrive-user_$stamp.reg" }
    )
    foreach ($backup in $backups) { Export-RegistryKey $backup.Key (Join-Path $Layout.Backups $backup.Name) | Out-Null }
    Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $windowsPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
    $machinePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
    $userPolicy = 'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive'
    New-Item -Path $windowsPolicy, $machinePolicy, $userPolicy -Force | Out-Null
    New-ItemProperty -Path $windowsPolicy -Name 'DisableFileSyncNGSC' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $windowsPolicy -Name 'DisableFileSync' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $machinePolicy -Name 'KFMBlockOptIn' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $machinePolicy -Name 'DisableAutoConfig' -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $machinePolicy -Name 'PreventNetworkTrafficPreUserSignIn' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $userPolicy -Name 'DisablePersonalSync' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $userPolicy -Name 'EnableAutoStart' -Value 0 -PropertyType DWord -Force | Out-Null
    Remove-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDrive' -ErrorAction SilentlyContinue
    if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
        Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like 'OneDrive*' } | ForEach-Object {
            Disable-ScheduledTask -InputObject $_ -ErrorAction SilentlyContinue | Out-Null
        }
    }
    foreach ($registryPath in @(
        'Registry::HKEY_CLASSES_ROOT\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}',
        'Registry::HKEY_CLASSES_ROOT\WOW6432Node\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}'
    )) {
        if (Test-Path -LiteralPath $registryPath) {
            New-ItemProperty -LiteralPath $registryPath -Name 'System.IsPinnedToNameSpaceTree' -Value 0 -PropertyType DWord -Force | Out-Null
        }
    }
    $uninstallers = @(
        (Join-Path $env:SystemRoot 'System32\OneDriveSetup.exe'),
        (Join-Path $env:SystemRoot 'SysWOW64\OneDriveSetup.exe'),
        (Join-Path $env:LOCALAPPDATA 'Microsoft\OneDrive\OneDriveSetup.exe')
    ) | Select-Object -Unique
    foreach ($uninstaller in $uninstallers) {
        if (Test-Path -LiteralPath $uninstaller -PathType Leaf) {
            try { Start-Process -FilePath $uninstaller -ArgumentList '/uninstall' -Wait -ErrorAction Stop | Out-Null } catch { Add-Log "OneDrive uninstaller: $($_.Exception.Message)" }
        }
    }
    Get-AppxPackage -Name 'Microsoft.OneDriveSync' -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    if ($WingetCommand) {
        $output = & $WingetCommand uninstall --id Microsoft.OneDrive --exact --silent --disable-interactivity --accept-source-agreements 2>&1
        $output | ForEach-Object { Add-Log "OneDrive winget: $_" }
    }
    $guardDirectory = Join-Path $env:ProgramData 'KlpInstall'
    $guardPath = Join-Path $guardDirectory 'OneDriveGuard.ps1'
    $guardContent = @'
$ErrorActionPreference = 'SilentlyContinue'
$windowsPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive'
$machinePolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive'
New-Item -Path $windowsPolicy, $machinePolicy -Force | Out-Null
New-ItemProperty -Path $windowsPolicy -Name 'DisableFileSyncNGSC' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $windowsPolicy -Name 'DisableFileSync' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $machinePolicy -Name 'KFMBlockOptIn' -Value 1 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $machinePolicy -Name 'DisableAutoConfig' -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $machinePolicy -Name 'PreventNetworkTrafficPreUserSignIn' -Value 1 -PropertyType DWord -Force | Out-Null
Get-Process -Name OneDrive | Stop-Process -Force
foreach ($uninstaller in @("$env:SystemRoot\System32\OneDriveSetup.exe", "$env:SystemRoot\SysWOW64\OneDriveSetup.exe")) {
    if (Test-Path -LiteralPath $uninstaller -PathType Leaf) { Start-Process -FilePath $uninstaller -ArgumentList '/uninstall' -Wait }
}
'@
    New-Item -ItemType Directory -Path $guardDirectory -Force | Out-Null
    [System.IO.File]::WriteAllText($guardPath, $guardContent, (New-Object System.Text.UTF8Encoding($false)))
    if (-not (Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) { return 'GUARD_UNAVAILABLE' }
    $guardAction = New-ScheduledTaskAction -Execute (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe') -Argument "-NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$guardPath`""
    $guardTrigger = New-ScheduledTaskTrigger -AtStartup
    $guardPrincipal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
    $guardSettings = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
    Register-ScheduledTask -TaskName 'KlpInstall-OneDriveGuard' -Action $guardAction -Trigger $guardTrigger -Principal $guardPrincipal -Settings $guardSettings -Force | Out-Null
    Get-Process -Name OneDrive -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    $policy = Get-ItemProperty -LiteralPath $windowsPolicy -Name 'DisableFileSyncNGSC' -ErrorAction SilentlyContinue
    $startup = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'OneDrive' -ErrorAction SilentlyContinue
    $guardTask = Get-ScheduledTask -TaskName 'KlpInstall-OneDriveGuard' -ErrorAction SilentlyContinue
    if ($policy.DisableFileSyncNGSC -ne 1 -or $startup.OneDrive -or -not $guardTask) { return 'VERIFICATION_FAILED' }
    $script:RebootRequired = $true
    Write-Ok 'OneDrive is uninstalled and blocked by policy'
    Write-Info 'Existing OneDrive files were preserved'
    return 'OK'
}

function Set-PageFileLocation {
    $memoryKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
    $backupPath = Join-Path $Layout.Backups "memory-management_$(Get-Date -Format 'yyyyMMdd_HHmmss').reg"
    Export-RegistryKey 'HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management' $backupPath | Out-Null
    $ramMB = [math]::Ceiling((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1MB)
    $initialMB = [int][math]::Min([math]::Max([math]::Ceiling(($ramMB * 0.5) / 1024) * 1024, 4096), 16384)
    $maximumMB = [int][math]::Min([math]::Max([math]::Ceiling($ramMB / 1024) * 1024, 8192), 32768)
    if ($maximumMB -lt $initialMB) { $maximumMB = $initialMB }
    $driveRoot = [System.IO.Path]::GetPathRoot($Base).TrimEnd('\')
    $pageFile = "$driveRoot\pagefile.sys"
    $computerSystem = Get-CimInstance -ClassName Win32_ComputerSystem
    if ($computerSystem.AutomaticManagedPagefile) {
        Set-CimInstance -InputObject $computerSystem -Property @{ AutomaticManagedPagefile = $false } | Out-Null
    }
    Set-ItemProperty -LiteralPath $memoryKey -Name 'PagingFiles' -Value @("$pageFile $initialMB $maximumMB")
    $configured = @((Get-ItemProperty -LiteralPath $memoryKey -Name 'PagingFiles').PagingFiles)
    if (-not ($configured | Where-Object { $_ -like "$pageFile *" })) { return 'VERIFICATION_FAILED' }
    $script:RebootRequired = $true
    Write-Ok "Page file -> $pageFile ($initialMB-$maximumMB MB)"
    return 'OK'
}

function Set-WindowsStoragePolicy {
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    Export-RegistryKey 'HKCU\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' (Join-Path $Layout.Backups "storage-sense_$stamp.reg") | Out-Null
    Export-RegistryKey 'HKLM\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps' (Join-Path $Layout.Backups "local-dumps_$stamp.reg") | Out-Null
    Export-RegistryKey 'HKLM\SYSTEM\CurrentControlSet\Control\CrashControl' (Join-Path $Layout.Backups "crash-control_$stamp.reg") | Out-Null
    $storageSense = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy'
    New-Item -Path $storageSense -Force | Out-Null
    New-ItemProperty -Path $storageSense -Name '01' -Value 1 -PropertyType DWord -Force | Out-Null
    $localDumps = 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps'
    New-Item -Path $localDumps -Force | Out-Null
    New-ItemProperty -Path $localDumps -Name 'DumpFolder' -Value $Layout.CrashDumps -PropertyType ExpandString -Force | Out-Null
    New-ItemProperty -Path $localDumps -Name 'DumpCount' -Value 5 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $localDumps -Name 'DumpType' -Value 1 -PropertyType DWord -Force | Out-Null
    $crashControl = 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl'
    New-ItemProperty -Path $crashControl -Name 'DumpFile' -Value (Join-Path $Layout.CrashDumps 'MEMORY.DMP') -PropertyType ExpandString -Force | Out-Null
    New-ItemProperty -Path $crashControl -Name 'MinidumpDir' -Value (Join-Path $Layout.CrashDumps 'Minidump') -PropertyType ExpandString -Force | Out-Null
    & powercfg.exe /hibernate off 2>&1 | ForEach-Object { Add-Log "powercfg: $_" }
    if ($LASTEXITCODE -ne 0) { return 'HIBERNATION_ERROR' }
    Write-Ok 'Storage Sense is enabled, crash dumps use D:, hibernation is disabled'
    return 'OK'
}

function Invoke-AppInstaller {
    param([string]$Kind, [string]$InstallerPath, [string]$Target)
    $filePath = $InstallerPath
    $arguments = @()
    switch ($Kind) {
        'Python' { $arguments = @('/quiet', 'InstallAllUsers=1', "TargetDir=$Target", 'PrependPath=1', 'Include_launcher=1', 'InstallLauncherAllUsers=1', 'Include_test=0') }
        'Firefox' { $arguments = @('/S', "/InstallDirectoryPath=$(Quote-Argument $Target)", '/DesktopShortcut=true', '/StartMenuShortcut=true', '/TaskbarShortcut=false', '/PrivateBrowsingShortcut=false', '/PreventRebootRequired=true') }
        'Node' { $filePath = 'msiexec.exe'; $arguments = @('/i', (Quote-Argument $InstallerPath), '/qn', '/norestart', "INSTALLDIR=$(Quote-Argument $Target)") }
        'VSCode' { $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/MERGETASKS=!runcode', "/DIR=$(Quote-Argument $Target)") }
        'Discord' { $arguments = @('/s') }
        'Git' { $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', '/NOCANCEL', '/SP-', "/DIR=$(Quote-Argument $Target)") }
        'Telegram' { $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/DIR=$(Quote-Argument $Target)") }
        default { throw "Unknown installer type: $Kind" }
    }
    Add-Log "Starting installer: $filePath $($arguments -join ' ')"
    $process = Start-Process -FilePath $filePath -ArgumentList $arguments -Wait -PassThru
    Add-Log "Installer exit code: $($process.ExitCode)"
    return $process.ExitCode -in @(0, 1641, 3010)
}

function Install-Application {
    param([string]$Name, [hashtable]$Spec)
    $target = Join-Path $Layout.Programs $Spec.Folder
    if (Test-AppAtTarget -Spec $Spec -Target $target) { Write-Ok "$Name is already installed on D:"; return 'EXISTS' }
    if (Test-WingetPackageInstalled -Id $Spec.Id) { Write-Warn "$Name is installed outside $target. Existing installation was not moved."; return 'EXISTS_OUTSIDE_D_LAYOUT' }
    if (-not (Initialize-AppStorage -Name $Name -ProgramTarget $target)) { Write-Warn "$Name was skipped because existing profile data must be moved manually"; return 'EXISTING_DATA_ON_C' }
    $downloadDirectory = Join-Path $Layout.Installers ($Spec.Id -replace '[^A-Za-z0-9._-]', '_')
    $installerPath = Download-WingetInstaller -Id $Spec.Id -Scope $Spec.Scope -Destination $downloadDirectory
    if (-not $installerPath) { return 'DOWNLOAD_FAILED' }
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    Write-Info "Installing $Name to $target..."
    if (-not (Invoke-AppInstaller -Kind $Spec.Installer -InstallerPath $installerPath -Target $target)) { Write-Fail "$Name installer returned an error"; return 'INSTALLER_FAILED' }
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        if (Test-AppAtTarget -Spec $Spec -Target $target) {
            if ($Name -eq 'Firefox') { Set-FirefoxConfiguration -Target $target }
            foreach ($relativePath in $Spec.PathFolders) {
                $pathEntry = if ([string]::IsNullOrWhiteSpace($relativePath)) { $target } else { Join-Path $target $relativePath }
                $pathTarget = if ($Spec.Scope -eq 'machine') { 'Machine' } else { 'User' }
                Add-ToPath -Directory $pathEntry -Target $pathTarget
            }
            Refresh-Path
            Write-Ok "$Name -> $target"
            return 'OK'
        }
        Start-Sleep -Seconds 1
    }
    Write-Fail "$Name was not found in $target after installation"
    return 'WRONG_LOCATION'
}

function Install-Chocolatey {
    $target = Join-Path $Layout.Programs 'Chocolatey'
    $executable = Join-Path $target 'bin\choco.exe'
    Set-PersistentEnvironment -Name 'ChocolateyInstall' -Value $target -Target 'Machine'
    if (Test-Path -LiteralPath $executable -PathType Leaf) { Add-ToPath (Split-Path -Parent $executable) Machine; Refresh-Path; Write-Ok "Chocolatey -> $target"; return 'EXISTS' }
    if ((Test-Path -LiteralPath $target) -and (Get-ChildItem -LiteralPath $target -Force | Select-Object -First 1)) { Write-Warn "Chocolatey target is not empty: $target"; return 'TARGET_NOT_EMPTY' }
    $scriptPath = Join-Path $Layout.Installers 'Chocolatey\install.ps1'
    New-Item -ItemType Directory -Path (Split-Path -Parent $scriptPath) -Force | Out-Null
    Invoke-DownloadFile -Uri 'https://community.chocolatey.org/install.ps1' -Destination $scriptPath -TimeoutSec 120
    Write-Ok "Downloaded to $scriptPath"
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $executable -PathType Leaf)) { Write-Fail 'Chocolatey installation failed'; return 'INSTALLER_FAILED' }
    Add-ToPath (Split-Path -Parent $executable) Machine
    Refresh-Path
    & $executable config set cacheLocation $Layout.ChocolateyCache --limit-output | Out-Null
    Write-Ok "Chocolatey -> $target"
    return 'OK'
}

function Install-Spotify {
    $target = Join-Path $Layout.Programs 'Spotify'
    if (-not (Initialize-AppStorage 'Spotify' $target)) { Write-Warn 'Spotify profile data already exists on C:'; return 'EXISTING_DATA_ON_C' }
    $choco = Join-Path $Layout.Programs 'Chocolatey\bin\choco.exe'
    if (-not (Test-Path -LiteralPath $choco -PathType Leaf)) { return 'NO_CHOCOLATEY' }
    & $choco install spotify -y --no-progress --limit-output
    if ($LASTEXITCODE -notin @(0, 1641, 3010)) { return 'INSTALLER_FAILED' }
    if (-not (Get-ChildItem -LiteralPath $target -Recurse -File -Filter 'Spotify.exe' -ErrorAction SilentlyContinue | Select-Object -First 1)) { return 'WRONG_LOCATION' }
    Write-Ok "Spotify -> $target"
    return 'OK'
}

function Install-Rust {
    $rustRoot = Join-Path $Layout.Programs 'Rust'
    $rustupHome = $Layout.RustupHome
    $cargoHome = $Layout.CargoHome
    $installerPath = Join-Path $Layout.Installers 'Rust\rustup-init.exe'
    Set-PersistentEnvironment 'RUSTUP_HOME' $rustupHome User
    Set-PersistentEnvironment 'CARGO_HOME' $cargoHome User
    if (Test-Path -LiteralPath (Join-Path $cargoHome 'bin\rustc.exe') -PathType Leaf) { Add-ToPath (Join-Path $cargoHome 'bin') User; Refresh-Path; return 'EXISTS' }
    New-Item -ItemType Directory -Path (Split-Path -Parent $installerPath) -Force | Out-Null
    $rustTarget = if ($Architecture -eq 'arm64') { 'aarch64-pc-windows-msvc' } else { 'x86_64-pc-windows-msvc' }
    Invoke-DownloadFile -Uri "https://static.rust-lang.org/rustup/dist/$rustTarget/rustup-init.exe" -Destination $installerPath
    Write-Ok "Downloaded to $installerPath"
    $process = Start-Process -FilePath $installerPath -ArgumentList @('-y', '--no-modify-path') -Wait -PassThru
    if ($process.ExitCode -ne 0 -or -not (Test-Path -LiteralPath (Join-Path $cargoHome 'bin\rustc.exe') -PathType Leaf)) { return 'INSTALLER_FAILED' }
    Add-ToPath (Join-Path $cargoHome 'bin') User
    Refresh-Path
    Write-Ok "Rust -> $rustRoot"
    return 'OK'
}

function Install-NpmTools {
    $prefix = $Layout.NpmGlobal
    $cache = $Layout.NpmCache
    New-Item -ItemType Directory -Path $prefix, $cache, $Layout.PnpmHome, $Layout.PnpmStore -Force | Out-Null
    Set-PersistentEnvironment 'NPM_CONFIG_PREFIX' $prefix User
    Set-PersistentEnvironment 'NPM_CONFIG_CACHE' $cache User
    Add-ToPath $prefix User
    Add-ToPath $Layout.PnpmHome User
    Refresh-Path
    if (-not (Get-Command npm -ErrorAction SilentlyContinue)) { return 'NO_NPM' }
    & npm config set prefix $prefix --location=user
    if ($LASTEXITCODE -ne 0) { return 'NPM_CONFIG_FAILED' }
    & npm config set cache $cache --location=user
    if ($LASTEXITCODE -ne 0) { return 'NPM_CONFIG_FAILED' }
    & npm install --global pnpm
    if ($LASTEXITCODE -ne 0) { return 'INSTALLER_FAILED' }
    Refresh-Path
    if (-not (Get-Command pnpm -ErrorAction SilentlyContinue)) { return 'NOT_IN_PATH' }
    & pnpm config set store-dir $Layout.PnpmStore --global
    if ($LASTEXITCODE -ne 0) { return 'PNPM_CONFIG_FAILED' }
    Write-Ok "pnpm -> $prefix"
    return 'OK'
}

function Install-TelegramProxy {
    $target = Join-Path $Layout.Programs 'TelegramProxy'
    New-Item -ItemType Directory -Path $target -Force | Out-Null
    try {
        $release = Invoke-RestMethod -Uri 'https://api.github.com/repos/Flowseal/tg-ws-proxy/releases/latest' -TimeoutSec 30
        $asset = $release.assets | Where-Object { $_.name -match '(?i)(windows|win).*(x64|amd64).*\.exe$|(?i)(x64|amd64).*(windows|win).*\.exe$' } | Select-Object -First 1
        if (-not $asset) { $asset = $release.assets | Where-Object { $_.name -match '(?i)\.exe$' } | Select-Object -First 1 }
        if (-not $asset) { return 'NO_WINDOWS_ASSET' }
        $destination = Join-Path $target $asset.name
        Invoke-DownloadFile -Uri $asset.browser_download_url -Destination $destination
        if (-not (Test-Path -LiteralPath $destination -PathType Leaf) -or (Get-Item -LiteralPath $destination).Length -eq 0) { return 'DOWNLOAD_FAILED' }
        Write-Ok "tg-ws-proxy -> $destination"
        return 'OK'
    } catch { Write-Fail $_.Exception.Message; return 'DOWNLOAD_FAILED' }
}

function Select-Components {
    $selection = [ordered]@{}
    Write-Section 'Component selection'
    Write-Host '  Choose what KlpInstall should install or configure.' -ForegroundColor White
    Write-Host '  The component name is shown before every Y/N question.' -ForegroundColor DarkGray

    $items = @(
        @{ Key = 'Python'; Label = 'Python 3.12'; Question = 'Install Python 3.12?'; Default = $true; Application = $true },
        @{ Key = 'Firefox'; Label = 'Mozilla Firefox'; Question = 'Install Firefox?'; Default = $true; Application = $true },
        @{ Key = 'Node.js'; Label = 'Node.js LTS'; Question = 'Install Node.js LTS?'; Default = $true; Application = $true },
        @{ Key = 'VS Code'; Label = 'Visual Studio Code'; Question = 'Install Visual Studio Code?'; Default = $true; Application = $true },
        @{ Key = 'Discord'; Label = 'Discord'; Question = 'Install Discord?'; Default = $true; Application = $true },
        @{ Key = 'Git'; Label = 'Git'; Question = 'Install Git?'; Default = $true; Application = $true },
        @{ Key = 'Telegram'; Label = 'Telegram Desktop'; Question = 'Install Telegram Desktop?'; Default = $true; Application = $true },
        @{ Key = 'Chocolatey'; Label = 'Chocolatey'; Question = 'Install Chocolatey?'; Default = $true; Application = $true },
        @{ Key = 'Spotify'; Label = 'Spotify'; Question = 'Install Spotify?'; Default = $true; Application = $true },
        @{ Key = 'Rust'; Label = 'Rust toolchain (rustup + cargo)'; Question = 'Install Rust?'; Default = $true; Application = $true },
        @{ Key = 'npm_tools'; Label = 'pnpm'; Question = 'Install pnpm?'; Default = $true; Application = $true },
        @{ Key = 'tg_proxy'; Label = 'tg-ws-proxy'; Question = 'Download tg-ws-proxy?'; Default = $false; Application = $true },
        @{ Key = 'KnownFolders'; Label = 'Windows personal folders (paths only)'; Question = 'Set Desktop, Documents, Downloads, Pictures, Music, Videos and Saved Games paths to D: without moving existing files?'; Default = $true; Application = $false },
        @{ Key = 'DisableOneDrive'; Label = 'OneDrive'; Question = 'Uninstall and permanently block OneDrive?'; Default = $true; Application = $false },
        @{ Key = 'BrowserDownloads'; Label = 'Browser storage'; Question = 'Force browser downloads, profiles and caches to D:?'; Default = $true; Application = $false },
        @{ Key = 'WindowsStorage'; Label = 'Windows storage'; Question = 'Move the page file to D: and disable hibernation?'; Default = $true; Application = $false }
    )

    $index = 0
    foreach ($item in $items) {
        $index++
        Write-Host ''
        Write-Host ("  [{0}/{1}] {2}" -f $index, $items.Count, $item.Label) -ForegroundColor Cyan
        if ($SkipApplications -and $item.Application) {
            $selection[$item.Key] = $false
            Write-Host '       SKIP - application installation disabled' -ForegroundColor DarkGray
            continue
        }
        $selection[$item.Key] = Prompt-YesNo $item.Question $item.Default
        $choice = if ($selection[$item.Key]) { 'YES' } else { 'NO' }
        $choiceColor = if ($selection[$item.Key]) { 'Green' } else { 'DarkGray' }
        Write-Host "       -> $choice" -ForegroundColor $choiceColor
    }

    if ($selection['DisableOneDrive']) { $selection['KnownFolders'] = $true }
    if ($selection['Spotify']) { $selection['Chocolatey'] = $true }
    Write-Host ''
    Write-Host '  Installation set:' -ForegroundColor Cyan
    foreach ($name in $selection.Keys) {
        $mark = if ($selection[$name]) { '[+]' } else { '[-]' }
        $color = if ($selection[$name]) { 'Green' } else { 'DarkGray' }
        Write-Host "  $mark $name" -ForegroundColor $color
    }
    return $selection
}

function Show-Plan {
    Write-Section 'Storage plan'
    foreach ($key in $Layout.Keys) { Write-Host "  $($key.PadRight(16)) $($Layout[$key])" }
    Write-Host ''
    Write-Host '  Existing personal files are NOT copied or moved.' -ForegroundColor Yellow
    Write-Host '  OneDrive is disabled only after known folders are redirected.' -ForegroundColor Yellow
    Write-Host '  Windows, WinGet itself, drivers and small system metadata remain on C:.' -ForegroundColor Yellow
    Write-Host '  Existing non-empty application folders on C: are never moved automatically.' -ForegroundColor Yellow
}

function Show-Report {
    Write-Section 'Installation report'
    $ok = 0; $failed = 0; $skipped = 0
    foreach ($name in $Results.Keys) {
        $status = $Results[$name]
        if ($status -in @('OK', 'EXISTS', 'CONFIGURED')) { Write-Host "  [OK]   $name - $status" -ForegroundColor Green; $ok++ }
        elseif ($status -eq 'SKIP') { Write-Host "  [SKIP] $name" -ForegroundColor DarkGray; $skipped++ }
        else { Write-Host "  [FAIL] $name - $status" -ForegroundColor Red; $failed++ }
    }
    $elapsed = (Get-Date) - $StartTime
    Write-Host ''
    Write-Host "  Successful: $ok  Failed: $failed  Skipped: $skipped"
    Write-Host "  Time: $($elapsed.ToString('hh\:mm\:ss'))"
    Write-Host "  Log: $LogFile"
    if ($RebootRequired) { Write-Host '  Restart Windows to finish applying storage and OneDrive policies.' -ForegroundColor Yellow }
    return $failed
}

function Test-FinalConfiguration {
    $checks = [ordered]@{}
    foreach ($entry in $StorageEnvironment.GetEnumerator()) {
        $actual = [Environment]::GetEnvironmentVariable($entry.Key, 'User')
        $checks["Environment:$($entry.Key)"] = $actual -ieq $entry.Value
    }
    foreach ($entry in $MachineEnvironment.GetEnumerator()) {
        $actual = [Environment]::GetEnvironmentVariable($entry.Key, 'Machine')
        $checks["MachineEnvironment:$($entry.Key)"] = $actual -ieq $entry.Value
    }
    $userPath = [Environment]::GetEnvironmentVariable('Path', 'User')
    $userPathParts = @($userPath -split ';' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    foreach ($directory in $PathDirectories) {
        $checks["Path:$directory"] = [bool]($userPathParts | Where-Object { $_.TrimEnd('\') -ieq $directory.TrimEnd('\') })
    }
    if ($Components['KnownFolders']) {
        Initialize-KnownFolderApi
        foreach ($spec in $KnownFolderSpecs) {
            $actual = [KlpKnownFolder]::Get($spec.Id)
            $checks["KnownFolder:$($spec.Name)"] = $actual.TrimEnd('\') -ieq $Layout[$spec.LayoutKey].TrimEnd('\')
        }
    }
    if ($Components['DisableOneDrive']) {
        $policy = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive' -ErrorAction SilentlyContinue
        $machinePolicy = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Policies\Microsoft\OneDrive' -ErrorAction SilentlyContinue
        $userPolicy = Get-ItemProperty -LiteralPath 'HKCU:\SOFTWARE\Policies\Microsoft\OneDrive' -ErrorAction SilentlyContinue
        $checks['OneDrive:Policy'] = $policy.DisableFileSyncNGSC -eq 1
        $checks['OneDrive:KFM'] = $machinePolicy.KFMBlockOptIn -eq 1
        $checks['OneDrive:AutoConfig'] = $machinePolicy.DisableAutoConfig -eq 3
        $checks['OneDrive:PersonalSync'] = $userPolicy.DisablePersonalSync -eq 1
        $checks['OneDrive:Process'] = -not [bool](Get-Process -Name OneDrive -ErrorAction SilentlyContinue)
        $checks['OneDrive:Guard'] = [bool](Get-ScheduledTask -TaskName 'KlpInstall-OneDriveGuard' -ErrorAction SilentlyContinue)
        $checks['OneDrive:GuardFile'] = Test-Path -LiteralPath (Join-Path $env:ProgramData 'KlpInstall\OneDriveGuard.ps1') -PathType Leaf
    }
    if ($Components['BrowserDownloads']) {
        $edge = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Policies\Microsoft\Edge' -ErrorAction SilentlyContinue
        $checks['Browser:Downloads'] = $edge.DownloadDirectory -ieq $Layout.Downloads
        $checks['Browser:EdgeProfile'] = $edge.UserDataDir -ieq $BrowserPolicyPaths.EdgeProfile
        $checks['Browser:EdgeCache'] = $edge.DiskCacheDir -ieq $BrowserPolicyPaths.EdgeCache
        $chrome = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Policies\Google\Chrome' -ErrorAction SilentlyContinue
        $checks['Browser:ChromeProfile'] = $chrome.UserDataDir -ieq $BrowserPolicyPaths.ChromeProfile
        $checks['Browser:ChromeCache'] = $chrome.DiskCacheDir -ieq $BrowserPolicyPaths.ChromeCache
    }
    if ($Components['WindowsStorage']) {
        $memoryManagement = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management'
        $pageFile = @($memoryManagement.PagingFiles)
        $dataRoot = [System.IO.Path]::GetPathRoot($Base).TrimEnd('\')
        $checks['Windows:PageFile'] = [bool]($pageFile | Where-Object { $_ -like "$dataRoot\pagefile.sys *" })
        $power = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\Power' -ErrorAction SilentlyContinue
        $checks['Windows:Hibernation'] = $power.HibernateEnabled -eq 0
        $storageSense = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy' -ErrorAction SilentlyContinue
        $checks['Windows:StorageSense'] = $storageSense.'01' -eq 1
        $localDumps = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting\LocalDumps' -ErrorAction SilentlyContinue
        $checks['Windows:LocalDumps'] = $localDumps.DumpFolder -ieq $Layout.CrashDumps
        $crashControl = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Control\CrashControl' -ErrorAction SilentlyContinue
        $checks['Windows:CrashDump'] = $crashControl.DumpFile -ieq (Join-Path $Layout.CrashDumps 'MEMORY.DMP')
    }
    foreach ($path in $Layout.Values) { $checks["Directory:$path"] = Test-Path -LiteralPath $path -PathType Container }
    $failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
    $report = [ordered]@{
        Timestamp = (Get-Date).ToString('o')
        ScriptVersion = $ScriptVersion
        SystemDriveFreeGB = Get-FreeSpaceGB ([System.IO.Path]::GetPathRoot($env:SystemRoot))
        DataDriveFreeGB = Get-FreeSpaceGB $Base
        RebootRequired = $RebootRequired
        Checks = $checks
    }
    $reportPath = Join-Path $Layout.Logs "verification_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    $report | ConvertTo-Json -Depth 6 | Out-File -LiteralPath $reportPath -Encoding UTF8
    if ($failed.Count -gt 0) {
        foreach ($item in $failed) { Write-Fail "Verification: $($item.Key)" }
        Write-Info "Verification report -> $reportPath"
        return 'VERIFICATION_FAILED'
    }
    Write-Ok "All storage paths verified -> $reportPath"
    return 'OK'
}

Write-Banner
try {
    Assert-SupportedSystem
    $Base = Get-FullBasePath -Path $BasePath
} catch {
    Write-Fail $_.Exception.Message
    if (-not $Silent) { Read-Host '  Enter to exit' }
    exit 1
}

$currentProfileName = [Environment]::UserName
$Layout = [ordered]@{
    Programs = Join-Path $Base 'Program'
    AppData = Join-Path $Base 'AppData'
    Desktop = Join-Path $Base 'Desktop'
    Documents = Join-Path $Base 'Documents'
    Downloads = Join-Path $Base 'Downloads'
    Pictures = Join-Path $Base 'Pictures'
    Music = Join-Path $Base 'Music'
    Videos = Join-Path $Base 'Videos'
    SavedGames = Join-Path $Base 'Games\Saved Games'
    Screenshots = Join-Path $Base 'Pictures\Screenshots'
    CameraRoll = Join-Path $Base 'Pictures\Camera Roll'
    Projects = Join-Path $Base 'Project'
    Games = Join-Path $Base 'Games'
    Models = Join-Path $Base 'Models'
    InstallerRoot = Join-Path $Base 'KlpInstall'
    Installers = Join-Path $Base 'Downloads\Installers'
    UserTemp = Join-Path $Base 'KlpInstall\Temp\User'
    SystemTemp = Join-Path $Base 'KlpInstall\Temp\System'
    Cache = Join-Path $Base 'KlpInstall\Cache'
    PipCache = Join-Path $Base 'KlpInstall\Cache\pip'
    NuGetCache = Join-Path $Base 'KlpInstall\Cache\NuGet'
    NpmCache = Join-Path $Base 'KlpInstall\Cache\npm'
    NpmGlobal = Join-Path $Base 'AppData\npm'
    PnpmHome = Join-Path $Base 'AppData\pnpm'
    PnpmStore = Join-Path $Base 'KlpInstall\Cache\pnpm-store'
    CorepackHome = Join-Path $Base 'KlpInstall\Cache\Corepack'
    YarnCache = Join-Path $Base 'KlpInstall\Cache\Yarn'
    RustupHome = Join-Path $Base 'Program\Rust\rustup'
    CargoHome = Join-Path $Base 'Program\Rust\cargo'
    PythonUserBase = Join-Path $Base 'AppData\Python'
    PipxHome = Join-Path $Base 'AppData\Python\pipx'
    PipxBin = Join-Path $Base 'AppData\Python\pipx-bin'
    PoetryCache = Join-Path $Base 'KlpInstall\Cache\Poetry'
    DotnetHome = Join-Path $Base 'AppData\dotnet'
    DotnetTools = Join-Path $Base 'AppData\dotnet-tools'
    GoHome = Join-Path $Base 'AppData\Go'
    GoCache = Join-Path $Base 'KlpInstall\Cache\Go\Build'
    GoModules = Join-Path $Base 'KlpInstall\Cache\Go\Modules'
    GradleHome = Join-Path $Base 'KlpInstall\Cache\Gradle'
    UvCache = Join-Path $Base 'KlpInstall\Cache\uv'
    BunHome = Join-Path $Base 'AppData\Bun'
    XdgCache = Join-Path $Base 'KlpInstall\Cache\XDG'
    HuggingFaceCache = Join-Path $Base 'KlpInstall\Cache\HuggingFace'
    TorchCache = Join-Path $Base 'KlpInstall\Cache\Torch'
    CudaCache = Join-Path $Base 'KlpInstall\Cache\CUDA'
    OllamaModels = Join-Path $Base 'Models\Ollama'
    PlaywrightCache = Join-Path $Base 'KlpInstall\Cache\Playwright'
    PuppeteerCache = Join-Path $Base 'KlpInstall\Cache\Puppeteer'
    CypressCache = Join-Path $Base 'KlpInstall\Cache\Cypress'
    ElectronCache = Join-Path $Base 'KlpInstall\Cache\Electron'
    ElectronBuilderCache = Join-Path $Base 'KlpInstall\Cache\ElectronBuilder'
    VcpkgCache = Join-Path $Base 'KlpInstall\Cache\vcpkg'
    Ccache = Join-Path $Base 'KlpInstall\Cache\ccache'
    Sccache = Join-Path $Base 'KlpInstall\Cache\sccache'
    AndroidSdk = Join-Path $Base 'AppData\Android\Sdk'
    CodexHome = Join-Path $Base 'AppData\Codex'
    EdgeProfile = Join-Path $Base "AppData\Browsers\$currentProfileName\Edge\User Data"
    EdgeCache = Join-Path $Base "KlpInstall\Cache\Browsers\$currentProfileName\Edge"
    ChromeProfile = Join-Path $Base "AppData\Browsers\$currentProfileName\Chrome\User Data"
    ChromeCache = Join-Path $Base "KlpInstall\Cache\Browsers\$currentProfileName\Chrome"
    ChocolateyCache = Join-Path $Base 'KlpInstall\Cache\Chocolatey'
    CrashDumps = Join-Path $Base 'KlpInstall\CrashDumps'
    Logs = Join-Path $Base 'KlpInstall\Logs'
    Backups = Join-Path $Base 'KlpInstall\Backups'
}

$BrowserPolicyPaths = [ordered]@{
    EdgeProfile = Join-Path $Base 'AppData\Browsers\${user_name}\Edge\User Data'
    EdgeCache = Join-Path $Base 'KlpInstall\Cache\Browsers\${user_name}\Edge'
    ChromeProfile = Join-Path $Base 'AppData\Browsers\${user_name}\Chrome\User Data'
    ChromeCache = Join-Path $Base 'KlpInstall\Cache\Browsers\${user_name}\Chrome'
}

$KnownFolderSpecs = @(
    @{ Name = 'Desktop'; Id = 'B4BFCC3A-DB2C-424C-B029-7FE99A87C641'; LayoutKey = 'Desktop' },
    @{ Name = 'Documents'; Id = 'FDD39AD0-238F-46AF-ADB4-6C85480369C7'; LayoutKey = 'Documents' },
    @{ Name = 'Downloads'; Id = '374DE290-123F-4565-9164-39C4925E467B'; LayoutKey = 'Downloads' },
    @{ Name = 'Pictures'; Id = '33E28130-4E1E-4676-835A-98395C3BC3BB'; LayoutKey = 'Pictures' },
    @{ Name = 'Music'; Id = '4BD8D571-6D19-48D3-BE97-422220080E43'; LayoutKey = 'Music' },
    @{ Name = 'Videos'; Id = '18989B1D-99B5-455B-841C-AB7C74E4DDFC'; LayoutKey = 'Videos' },
    @{ Name = 'Saved Games'; Id = '4C5C32FF-BB9D-43B0-B5B4-2D72E54EAAA4'; LayoutKey = 'SavedGames' },
    @{ Name = 'Screenshots'; Id = 'B7BEDE81-DF94-4682-A7D8-57A52620B86F'; LayoutKey = 'Screenshots' },
    @{ Name = 'Camera Roll'; Id = 'AB5FB87B-7CE2-4F83-915D-550846C9537B'; LayoutKey = 'CameraRoll' }
)

$StorageEnvironment = [ordered]@{
    TEMP = $Layout.UserTemp
    TMP = $Layout.UserTemp
    TMPDIR = $Layout.UserTemp
    PIP_CACHE_DIR = $Layout.PipCache
    PYTHONUSERBASE = $Layout.PythonUserBase
    NUGET_PACKAGES = $Layout.NuGetCache
    NPM_CONFIG_PREFIX = $Layout.NpmGlobal
    NPM_CONFIG_CACHE = $Layout.NpmCache
    NPM_CONFIG_USERCONFIG = Join-Path $Layout.NpmGlobal '.npmrc'
    PNPM_HOME = $Layout.PnpmHome
    COREPACK_HOME = $Layout.CorepackHome
    YARN_CACHE_FOLDER = $Layout.YarnCache
    RUSTUP_HOME = $Layout.RustupHome
    CARGO_HOME = $Layout.CargoHome
    PIPX_HOME = $Layout.PipxHome
    PIPX_BIN_DIR = $Layout.PipxBin
    POETRY_CACHE_DIR = $Layout.PoetryCache
    DOTNET_CLI_HOME = $Layout.DotnetHome
    DOTNET_BUNDLE_EXTRACT_BASE_DIR = Join-Path $Layout.UserTemp 'dotnet-bundle'
    GOPATH = $Layout.GoHome
    GOCACHE = $Layout.GoCache
    GOMODCACHE = $Layout.GoModules
    GRADLE_USER_HOME = $Layout.GradleHome
    UV_CACHE_DIR = $Layout.UvCache
    BUN_INSTALL = $Layout.BunHome
    XDG_CACHE_HOME = $Layout.XdgCache
    HF_HOME = $Layout.HuggingFaceCache
    TORCH_HOME = $Layout.TorchCache
    CUDA_CACHE_PATH = $Layout.CudaCache
    OLLAMA_MODELS = $Layout.OllamaModels
    PLAYWRIGHT_BROWSERS_PATH = $Layout.PlaywrightCache
    PUPPETEER_CACHE_DIR = $Layout.PuppeteerCache
    CYPRESS_CACHE_FOLDER = $Layout.CypressCache
    ELECTRON_CACHE = $Layout.ElectronCache
    ELECTRON_BUILDER_CACHE = $Layout.ElectronBuilderCache
    VCPKG_DEFAULT_BINARY_CACHE = $Layout.VcpkgCache
    CCACHE_DIR = $Layout.Ccache
    SCCACHE_DIR = $Layout.Sccache
    ANDROID_HOME = $Layout.AndroidSdk
    ANDROID_SDK_ROOT = $Layout.AndroidSdk
    CODEX_HOME = $Layout.CodexHome
}

$MachineEnvironment = [ordered]@{
    TEMP = $Layout.SystemTemp
    TMP = $Layout.SystemTemp
}

$PathDirectories = @(
    $Layout.NpmGlobal,
    $Layout.PnpmHome,
    $Layout.PipxBin,
    (Join-Path $Layout.CargoHome 'bin'),
    (Join-Path $Layout.BunHome 'bin'),
    $Layout.DotnetTools,
    (Join-Path $Layout.GoHome 'bin')
)

$Architecture = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
$Components = Select-Components
Show-Plan
if ($PlanOnly) { Write-Info 'PlanOnly completed without changing the system'; if (-not $Silent) { Read-Host '  Enter to exit' }; exit 0 }
if (-not (Prompt-YesNo 'Apply this storage plan?')) { Write-Info 'Cancelled'; exit 0 }

Write-Section 'Pre-flight checks'
try {
    $volume = Get-DataVolume -Path $Base
    Write-Ok "$($volume.DeviceID) is a fixed NTFS drive"
} catch {
    Write-Fail $_.Exception.Message
    if (-not $Silent) { Read-Host '  Enter to exit' }
    exit 1
}
$freeSpace = Get-FreeSpaceGB -Path $Base
if ($freeSpace -lt $MinimumFreeSpaceGB) { Write-Warn "Only $freeSpace GB is free on $([System.IO.Path]::GetPathRoot($Base))"; if (-not (Prompt-YesNo 'Continue with low free space?' $false)) { exit 1 } } else { Write-Ok "$freeSpace GB is free on $([System.IO.Path]::GetPathRoot($Base))" }
$systemDrive = [System.IO.Path]::GetPathRoot($env:SystemRoot)
$systemFreeSpace = Get-FreeSpaceGB -Path $systemDrive
if ($systemFreeSpace -lt 10) { Write-Warn "The system drive has only $systemFreeSpace GB free. This setup prevents new growth but does not delete existing data." } else { Write-Info "$systemFreeSpace GB is free on the system drive" }
$internetRequired = $false
foreach ($name in $AppSpecs.Keys) { if ($Components[$name]) { $internetRequired = $true } }
foreach ($name in @('Chocolatey', 'Spotify', 'Rust', 'npm_tools', 'tg_proxy')) { if ($Components[$name]) { $internetRequired = $true } }
if ($internetRequired) {
    if (-not (Test-Internet)) { Write-Fail 'Internet connection is unavailable'; if (-not $Silent) { Read-Host '  Enter to exit' }; exit 1 }
    Write-Ok 'Internet connection'
}

foreach ($directory in $Layout.Values) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
foreach ($directory in $PathDirectories) { New-Item -ItemType Directory -Path $directory -Force | Out-Null }
$LogFile = Join-Path $Layout.Logs "setup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Add-Log "Setup started, version $ScriptVersion, base path $Base"
$aclOutput = & icacls.exe $Layout.SystemTemp /inheritance:e /grant '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' '*S-1-5-11:(OI)(CI)M' 2>&1
$aclExitCode = $LASTEXITCODE
$aclOutput | ForEach-Object { Add-Log "icacls: $_" }
if ($aclExitCode -ne 0) { Write-Fail 'Could not configure permissions for the system temporary directory'; if (-not $Silent) { Read-Host '  Enter to exit' }; exit 1 }
Write-Ok "System temporary directory permissions -> $($Layout.SystemTemp)"

Write-Section 'Persistent storage locations'
Export-EnvironmentState
foreach ($entry in $StorageEnvironment.GetEnumerator()) { Set-PersistentEnvironment -Name $entry.Key -Value $entry.Value -Target User }
foreach ($entry in $MachineEnvironment.GetEnumerator()) { Set-PersistentEnvironment -Name $entry.Key -Value $entry.Value -Target Machine }
$env:TEMP = $Layout.UserTemp
$env:TMP = $Layout.UserTemp
$env:TMPDIR = $Layout.UserTemp
foreach ($directory in $PathDirectories) { Add-ToPath -Directory $directory -Target User }
Refresh-Path

Write-Section 'Windows folders and storage'
if ($Components['KnownFolders']) {
    try { $Results['Known folders'] = Set-KnownFolderLayout } catch { Write-Fail "Known folders: $($_.Exception.Message)"; $Results['Known folders'] = 'ERROR' }
} else { $Results['Known folders'] = 'SKIP' }
if ($Components['BrowserDownloads']) {
    try { $Results['Browser downloads'] = Set-BrowserDownloadPolicies } catch { Write-Fail "Browser downloads: $($_.Exception.Message)"; $Results['Browser downloads'] = 'ERROR' }
} else { $Results['Browser downloads'] = 'SKIP' }
if ($Components['WindowsStorage']) {
    try { $Results['Page file'] = Set-PageFileLocation } catch { Write-Fail "Page file: $($_.Exception.Message)"; $Results['Page file'] = 'ERROR' }
    try { $Results['Windows storage'] = Set-WindowsStoragePolicy } catch { Write-Fail "Windows storage: $($_.Exception.Message)"; $Results['Windows storage'] = 'ERROR' }
} else {
    $Results['Page file'] = 'SKIP'
    $Results['Windows storage'] = 'SKIP'
}

$wingetRequired = $false
foreach ($name in $AppSpecs.Keys) { if ($Components[$name]) { $wingetRequired = $true } }
if ($wingetRequired) {
    try { $Results['WinGet'] = Ensure-Winget } catch { Write-Fail "WinGet: $($_.Exception.Message)"; $Results['WinGet'] = 'ERROR' }
} else {
    $script:WingetCommand = Resolve-WingetCommand
    $Results['WinGet'] = 'SKIP'
}

if ($Components['DisableOneDrive']) {
    if ($Results['Known folders'] -eq 'OK') {
        try { $Results['OneDrive'] = Disable-OneDrive } catch { Write-Fail "OneDrive: $($_.Exception.Message)"; $Results['OneDrive'] = 'ERROR' }
    } else {
        Write-Fail 'OneDrive was not disabled because known-folder redirection failed'
        $Results['OneDrive'] = 'KNOWN_FOLDER_FAILED'
    }
} else { $Results['OneDrive'] = 'SKIP' }

Write-Section 'Applications'
foreach ($name in $AppSpecs.Keys) {
    if ($Components[$name]) {
        if ($Results['WinGet'] -eq 'OK') {
            try { $Results[$name] = Install-Application -Name $name -Spec $AppSpecs[$name] } catch { Write-Fail "${name}: $($_.Exception.Message)"; $Results[$name] = 'ERROR' }
        } else {
            $Results[$name] = 'NO_WINGET'
        }
    } else { $Results[$name] = 'SKIP' }
}

Write-Section 'Additional tools'
if ($Components['Chocolatey']) { try { $Results['Chocolatey'] = Install-Chocolatey } catch { Write-Fail "Chocolatey: $($_.Exception.Message)"; $Results['Chocolatey'] = 'ERROR' } } else { $Results['Chocolatey'] = 'SKIP' }
if ($Components['Spotify']) { try { $Results['Spotify'] = Install-Spotify } catch { Write-Fail "Spotify: $($_.Exception.Message)"; $Results['Spotify'] = 'ERROR' } } else { $Results['Spotify'] = 'SKIP' }
if ($Components['Rust']) { try { $Results['Rust'] = Install-Rust } catch { Write-Fail "Rust: $($_.Exception.Message)"; $Results['Rust'] = 'ERROR' } } else { $Results['Rust'] = 'SKIP' }
if ($Components['npm_tools']) { try { $Results['pnpm'] = Install-NpmTools } catch { Write-Fail "pnpm: $($_.Exception.Message)"; $Results['pnpm'] = 'ERROR' } } else { $Results['pnpm'] = 'SKIP' }
if ($Components['tg_proxy']) { $Results['tg-ws-proxy'] = Install-TelegramProxy } else { $Results['tg-ws-proxy'] = 'SKIP' }

Refresh-Path
Write-Section 'Final verification'
try { $Results['Configuration'] = Test-FinalConfiguration } catch { Write-Fail "Configuration: $($_.Exception.Message)"; $Results['Configuration'] = 'ERROR' }
$failedCount = Show-Report
Add-Log "Results: $($Results | Out-String)"
if (-not $Silent) { Read-Host '  Enter to exit' }
if ($failedCount -gt 0) { exit 1 }
exit 0
