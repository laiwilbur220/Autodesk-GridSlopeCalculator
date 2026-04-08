@echo off
@echo off
echo Compiling Civil 3D Grid Slope Tool...

"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" /target:library /out:GridSlopeCalculatorV3.dll /reference:acmgd.dll,acdbmgd.dll,accoremgd.dll GridSlopeCalculatorV3.cs

echo.
echo Compilation finished. Look for GridSlopeCalculator.dll in the folder.
pause