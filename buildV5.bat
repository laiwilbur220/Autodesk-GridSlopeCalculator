@echo off
setlocal enabledelayedexpansion

echo Compiling Civil 3D Grid Slope Tool V5...
echo.

set "WINBASE=C:\Windows\Microsoft.NET\assembly\GAC_MSIL\WindowsBase\v4.0_4.0.0.0__31bf3856ad364e35\WindowsBase.dll"
set "OUTFILE=GridSlopeCalculatorV5.dll"
echo Target: !OUTFILE!

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:library /out:"!OUTFILE!" /reference:acmgd.dll,acdbmgd.dll,accoremgd.dll,"%WINBASE%" GridSlopeCalculatorV5.cs

echo.
if exist "!OUTFILE!" (
    echo SUCCESS: !OUTFILE! created.
    echo AutoCAD commands: CalcGridSlopeCSV5, UpdateGridSlopeCSV5
) else (
    echo FAILED: Check errors above.
)
pause
