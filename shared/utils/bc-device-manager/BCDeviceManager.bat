@echo off
:: BCDeviceManager.bat
:: BriComp Computers, LLC
:: Double-click launcher for Invoke-BCDeviceManager.ps1
:: Works from Explorer, desktop shortcuts, or cmd.exe
:: No PowerShell version dependency - cmd.exe always runs this.

setlocal

:: Script lives alongside this .bat file
set "SCRIPT=%~dp0Invoke-BCDeviceManager.ps1"

:: Check the script exists
if not exist "%SCRIPT%" (
    echo.
    echo  ERROR: Invoke-BCDeviceManager.ps1 not found.
    echo  Expected: %SCRIPT%
    echo.
    echo  Ensure BCDeviceManager.bat and Invoke-BCDeviceManager.ps1
    echo  are in the same folder.
    echo.
    pause
    exit /b 1
)

:: Check if pwsh (PowerShell 7) is available
where pwsh >nul 2>&1
if %ERRORLEVEL% EQU 0 goto :launch

:: PS7 not found - check winget
echo.
echo  -------------------------------------------------------
echo   PowerShell 7+ Required
echo  -------------------------------------------------------
echo.
echo   BriComp Device Manager requires PowerShell 7 or later.
echo   PowerShell 7 is not currently installed.
echo.

where winget >nul 2>&1
if %ERRORLEVEL% EQU 0 goto :offer_install

:: No winget either - show manual instructions
echo   Install PowerShell 7 using one of these methods:
echo.
echo   Option 1 - Microsoft Store (Windows 10/11):
echo     Search for "PowerShell" in the Microsoft Store
echo.
echo   Option 2 - Direct download:
echo     https://aka.ms/powershell
echo     Download the .msi for your architecture (x64 recommended)
echo.
echo   Option 3 - Install winget first, then PowerShell 7:
echo     https://aka.ms/getwinget
echo     Then run: winget install Microsoft.PowerShell
echo.
pause
exit /b 1

:offer_install
echo   winget is available. Would you like to install PowerShell 7 now?
echo.
set /p "CHOICE=   Install now? (Y/N): "
if /i "%CHOICE%"=="Y" goto :install
echo.
echo   To install manually: winget install Microsoft.PowerShell
echo   Or download from:    https://aka.ms/powershell
echo.
pause
exit /b 0

:install
echo.
echo   Installing PowerShell 7 via winget...
echo.
winget install Microsoft.PowerShell --accept-source-agreements --accept-package-agreements
if %ERRORLEVEL% EQU 0 (
    echo.
    echo   Installation complete.
    echo   Relaunching BriComp Device Manager...
    echo.
    :: Refresh PATH so pwsh is found
    where pwsh >nul 2>&1
    if %ERRORLEVEL% EQU 0 goto :launch
    echo   Please close and reopen this launcher after installation.
    pause
    exit /b 0
) else (
    echo.
    echo   Installation may have failed. Try manually:
    echo     winget install Microsoft.PowerShell
    echo.
    pause
    exit /b 1
)

:launch
:: Check for Administrator rights - Device Manager needs them
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo   Requesting Administrator privileges...
    echo.
    :: Re-launch this .bat as Administrator using PowerShell elevation
    pwsh -WindowStyle Hidden -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b 0
)

:: Launch the GUI in pwsh - no extra console window
:: -WindowStyle Hidden hides the pwsh console; the WPF window takes over
start "" pwsh -WindowStyle Hidden -ExecutionPolicy Bypass -File "%SCRIPT%"
exit /b 0
