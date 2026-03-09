```powershell
1. New-Item -ItemType Directory -Path "D:\i" -Force
2. powershell -Command "Start-Process powershell -Verb RunAs -ArgumentList '-NoExit','-Command','cd D:\\i; .\\i.ps1'"
```
