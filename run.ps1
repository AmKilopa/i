$src = Join-Path $PSScriptRoot "i.ps1"
if (-not (Test-Path $src)) { exit 1 }
New-Item -ItemType Directory -Path "D:\i" -Force | Out-Null
Copy-Item $src "D:\i\" -Force
powershell -ExecutionPolicy Bypass -Command "Start-Process powershell -Verb RunAs -ArgumentList '-ExecutionPolicy','Bypass','-NoExit','-File','D:\i\i.ps1'"
