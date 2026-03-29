@echo off
echo ============================================
echo  Webcam Settings Manager - Build Script
echo ============================================

REM Find the latest .NET Framework csc.exe
set CSC=
if exist "%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe" (
    set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
) else if exist "%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe" (
    set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) else (
    echo ERROR: .NET Framework csc.exe not found!
    echo Please ensure .NET Framework 4.x is installed.
    pause
    exit /b 1
)

echo Using compiler: %CSC%
echo.

"%CSC%" -target:winexe -out:WebcamSettingsManager.exe -optimize+ -platform:anycpu -reference:System.dll -reference:System.Drawing.dll -reference:System.Windows.Forms.dll -reference:System.Core.dll WebcamSettingsManager.cs

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ============================================
    echo  BUILD SUCCESSFUL: WebcamSettingsManager.exe
    echo ============================================
) else (
    echo.
    echo ============================================
    echo  BUILD FAILED
    echo ============================================
)

pause
