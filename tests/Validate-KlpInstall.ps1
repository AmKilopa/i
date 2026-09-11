$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$installerPath = Join-Path $repositoryRoot 'i.ps1'
$runnerPath = Join-Path $repositoryRoot 'run.ps1'
$bootstrapPath = Join-Path $repositoryRoot 'install.bat'
$files = @($installerPath, $runnerPath)

foreach ($file in $files) {
    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors.Count -gt 0) { throw "$file contains PowerShell syntax errors: $($errors.Message -join '; ')" }
}

$installer = Get-Content -LiteralPath $installerPath -Raw
$runner = Get-Content -LiteralPath $runnerPath -Raw
$bootstrap = Get-Content -LiteralPath $bootstrapPath -Raw

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

# Regression guards for the Windows command-line bug that corrupted a drive-root
# argument such as D:\ while elevating through Start-Process -Verb RunAs.
if ($runner -notmatch [regex]::Escape('-EncodedCommand')) { throw 'run.ps1 must use EncodedCommand for UAC elevation' }
if ($runner -notmatch [regex]::Escape('ConvertTo-PowerShellLiteral')) { throw 'run.ps1 must safely quote values inside the encoded elevation command' }
if ($runner -notmatch [regex]::Escape('-ElevatedChild')) { throw 'run.ps1 must mark the elevated child so failures remain visible' }
if ($runner -match [regex]::Escape('ConvertTo-QuotedArgument')) { throw 'Legacy Windows command-line quoting helper must not return' }
if ($bootstrap -notmatch [regex]::Escape('-BasePath "D:\."')) { throw 'install.bat must pass the safe D:\. root spelling' }

# Verify that the literal conversion preserves a trailing backslash and apostrophes.
function ConvertTo-TestPowerShellLiteral {
    param([string]$Value)
    return "'" + $Value.Replace("'", "''") + "'"
}
$rootLiteral = ConvertTo-TestPowerShellLiteral 'D:\'
if ($rootLiteral -ne "'D:\'") { throw "Drive-root literal was corrupted: $rootLiteral" }
$quotedLiteral = ConvertTo-TestPowerShellLiteral "D:\O'Brien\"
if ($quotedLiteral -ne "'D:\O''Brien\'") { throw "Apostrophe/trailing-backslash literal was corrupted: $quotedLiteral" }

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

    # Exercise both accepted spellings. They must normalize to the same drive root.
    foreach ($testBase in @("$driveName\", "$driveName\.")) {
        & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $runnerPath -Automatic -SkipApplications -PlanOnly -BasePath $testBase
        if ($LASTEXITCODE -ne 0) { throw "PlanOnly returned exit code $LASTEXITCODE for $testBase" }
    }

    if (Get-ChildItem -LiteralPath $temporaryDirectory -Force | Select-Object -First 1) { throw 'PlanOnly changed the target drive' }
} finally {
    if ($driveName) { & subst.exe $driveName /D 2>$null }
    if (Test-Path -LiteralPath $temporaryDirectory) { Remove-Item -LiteralPath $temporaryDirectory -Recurse -Force }
}

Write-Host 'KlpInstall validation passed.' -ForegroundColor Green
