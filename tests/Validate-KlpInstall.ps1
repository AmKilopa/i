$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$files = @(
    (Join-Path $repositoryRoot 'i.ps1'),
    (Join-Path $repositoryRoot 'run.ps1')
)

foreach ($file in $files) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count -gt 0) { throw "$file contains PowerShell syntax errors: $($errors.Message -join '; ')" }
}

$installer = Get-Content -LiteralPath (Join-Path $repositoryRoot 'i.ps1') -Raw
$requiredValues = @(
    'DisableFileSyncNGSC',
    'KFMBlockOptIn',
    'SHSetKnownFolderPath',
    'Microsoft.WinGet.Client',
    'PLAYWRIGHT_BROWSERS_PATH',
    'VCPKG_DEFAULT_BINARY_CACHE',
    'UserDataDir',
    'DiskCacheDir',
    'OneDriveGuard'
)
foreach ($value in $requiredValues) {
    if ($installer -notmatch [regex]::Escape($value)) { throw "Required setting is missing: $value" }
}
foreach ($obsoleteValue in @('Set-DownloadsKnownFolder', "Components['DownloadsFolder']")) {
    if ($installer.Contains($obsoleteValue)) { throw "Obsolete setting remains: $obsoleteValue" }
}

$temporaryDirectory = Join-Path ([System.IO.Path]::GetTempPath()) "KlpInstallTest-$([guid]::NewGuid().ToString('N'))"
New-Item -ItemType Directory -Path $temporaryDirectory -Force | Out-Null
$driveName = $null
foreach ($letter in @('T', 'U', 'V', 'W', 'X', 'Y', 'Z')) {
    if (-not (Test-Path -LiteralPath "${letter}:\")) { $driveName = "${letter}:"; break }
}
if (-not $driveName) { throw 'No temporary drive letter is available for validation' }

try {
    & subst.exe $driveName $temporaryDirectory
    if ($LASTEXITCODE -ne 0) { throw 'Could not create the temporary validation drive' }
    $testBase = "$driveName\"
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repositoryRoot 'run.ps1') -Automatic -SkipApplications -PlanOnly -BasePath $testBase
    if ($LASTEXITCODE -ne 0) { throw "PlanOnly returned exit code $LASTEXITCODE" }
    if (Get-ChildItem -LiteralPath $temporaryDirectory -Force | Select-Object -First 1) { throw 'PlanOnly changed the target drive' }
} finally {
    if ($driveName) { & subst.exe $driveName /D 2>$null }
    if (Test-Path -LiteralPath $temporaryDirectory) { Remove-Item -LiteralPath $temporaryDirectory -Force }
}

Write-Host 'KlpInstall validation passed.' -ForegroundColor Green
