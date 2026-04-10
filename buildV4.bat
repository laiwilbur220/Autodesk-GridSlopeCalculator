@echo off
setlocal enabledelayedexpansion

:: Auto-detect version from CommandMethod attribute in the .cs file
set "VERSION="
for /f "tokens=*" %%L in ('findstr /C:"CommandMethod" GridSlopeCalculatorV4.cs') do (
    set "LINE=%%L"
)

:: Extract version string between quotes: CalcGridSlopeCSV4_N -> V4_N
for /f "tokens=2 delims=()""" %%V in ("%LINE%") do (
    set "CMD=%%V"
)
:: CMD is now e.g. "CalcGridSlopeCSV4_3", strip the prefix to get "V4_3"
set "VERSION=%CMD:CalcGridSlopeCS=%"

if "%VERSION%"=="" (
    echo ERROR: Could not detect version from GridSlopeCalculatorV4.cs
    pause
    exit /b 1
)

echo Compiling Civil 3D Grid Slope Tool %VERSION%...
echo.

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:library /out:GridSlopeCalculator%VERSION%.dll /reference:acmgd.dll,acdbmgd.dll,accoremgd.dll GridSlopeCalculatorV4.cs

echo.
if exist "GridSlopeCalculator%VERSION%.dll" (
    echo SUCCESS: GridSlopeCalculator%VERSION%.dll created.
    echo AutoCAD command: CalcGridSlopeCS%VERSION%
) else (
    echo FAILED: Check errors above.
)
pause
