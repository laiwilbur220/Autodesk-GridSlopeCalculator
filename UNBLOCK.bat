@echo off
echo ============================================
echo   Unblock DLLs for AutoCAD NETLOAD
echo ============================================
echo.
echo If you downloaded this project as a ZIP from GitHub,
echo Windows will block the DLL files. This script removes
echo the block so AutoCAD can load them.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem '%~dp0' -Recurse -Filter *.dll | ForEach-Object { Unblock-File $_.FullName; Write-Host ('Unblocked: ' + $_.Name) }"
echo.
echo Done! You can now use NETLOAD in AutoCAD.
pause
