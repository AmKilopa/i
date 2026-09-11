@echo off
title KlpInstall - Automatic D Setup
set "KLP_SELF=%~f0"
set "KLP_RUNNER=%~dp0run.ps1"
set "KLP_BOOTSTRAP=%TEMP%\KlpInstallBootstrap"

fltmc >nul 2>&1
if errorlevel 1 (
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath $env:KLP_SELF -Verb RunAs"
    if errorlevel 1 (
        echo.
        echo  Administrator access was not granted.
        pause >nul
    )
    exit /b
)

echo.
echo  KlpInstall
echo  Configuring clean Windows with data storage on drive D:
echo.

if not exist "%KLP_RUNNER%" (
    if not exist "%KLP_BOOTSTRAP%" mkdir "%KLP_BOOTSTRAP%"
    set "KLP_RUNNER=%KLP_BOOTSTRAP%\run.ps1"
    powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $ProgressPreference='SilentlyContinue'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $target=Join-Path $env:KLP_BOOTSTRAP 'run.ps1'; for($i=1;$i -le 3;$i++){try{Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/AmKilopa/i/main/run.ps1' -OutFile $target -TimeoutSec 180; break}catch{if($i -eq 3){throw}; Start-Sleep -Seconds ($i*2)}}"
    if errorlevel 1 goto download_error
)

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%KLP_RUNNER%" -Automatic -BasePath "D:\"
set "KLP_EXIT=%ERRORLEVEL%"
echo.
if "%KLP_EXIT%"=="0" (
    echo  KlpInstall completed successfully.
) else (
    echo  KlpInstall failed with exit code %KLP_EXIT%.
)
echo  Press any key to close this window.
pause >nul
exit /b %KLP_EXIT%

:download_error
echo.
echo  KlpInstall could not be downloaded from GitHub.
echo  Check the internet connection and try again.
pause >nul
exit /b 1
