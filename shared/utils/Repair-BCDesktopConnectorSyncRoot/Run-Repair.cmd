@echo off
REM ============================================================================
REM  Repair-BCDesktopConnectorSyncRoot - launcher
REM  BriComp IT Consulting Services   https://bricomp.com
REM  Version 1.1.4
REM
REM  HOW TO USE
REM    Right-click this file and choose "Run as administrator".
REM    Approve the Windows security prompt if it appears.
REM    The computer will restart automatically when the repair is done.
REM
REM    To preview without changing anything:   Run-Repair.cmd /whatif
REM    To repair but not restart:              Run-Repair.cmd /norestart
REM
REM  Logs and the registry backup are written to:
REM    C:\ProgramData\BriComp\DCSyncRootFix
REM ============================================================================

setlocal
set "SCRIPTDIR=%~dp0"
set "PS1=%SCRIPTDIR%Repair-BCDesktopConnectorSyncRoot.ps1"
set "PSARGS="

if /i "%~1"=="/whatif"    set "PSARGS=-WhatIf"
REM /dryrun still accepted so older habits and existing RMM jobs keep working
if /i "%~1"=="/dryrun"    set "PSARGS=-WhatIf"
if /i "%~1"=="/norestart" set "PSARGS=-NoRestart"

if not exist "%PS1%" (
    echo.
    echo  ERROR: Repair-BCDesktopConnectorSyncRoot.ps1 was not found next to this file.
    echo  Keep both files together in the same folder.
    echo.
    pause
    exit /b 1
)

REM --- Are we elevated? -------------------------------------------------------
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  Requesting administrator rights...
    echo.
    REM -ArgumentList must be omitted entirely when there is no argument.
    REM Passing an empty string fails with "Cannot validate argument on parameter
    REM 'ArgumentList'", which is what a plain double-click used to hit.
    REM Single lines on purpose: a caret continuation inside a nested block is a
    REM well known batch parsing trap.
    if "%~1"=="" powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    if not "%~1"=="" powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -ArgumentList '%~1' -Verb RunAs"
    exit /b 0
)

REM --- Is script execution locked down by Group Policy? -----------------------
REM Execution policy governs script FILES, not -Command input, so this check still
REM runs on a machine where the .ps1 itself would be refused. A Group Policy setting
REM overrides the -ExecutionPolicy switch used below, so detect that and say so
REM plainly rather than failing with a confusing error.
powershell -NoProfile -Command "$bad=$false; foreach($e in (Get-ExecutionPolicy -List)){ if(($e.Scope -eq 'MachinePolicy' -or $e.Scope -eq 'UserPolicy') -and $e.ExecutionPolicy -ne 'Undefined'){$bad=$true} }; if($bad){exit 1}else{exit 0}"
if %errorlevel% equ 1 (
    echo.
    echo  CANNOT RUN ON THIS COMPUTER
    echo.
    echo  This computer enforces PowerShell script policy through Group Policy,
    echo  which overrides the setting this launcher uses. The repair cannot run
    echo  until the BriComp Computers, LLC code signing certificate is added to
    echo  the Trusted Publishers store on this computer.
    echo.
    echo  Nothing was changed. Please send this message to your IT contact.
    echo.
    pause
    exit /b 5
)

echo.
echo  Running Desktop Connector sync root repair...
echo.

REM Bypass is used so a downloaded copy (which carries Mark-of-the-Web) runs without
REM prompting a non-technical operator. If your environment wants signature
REM enforcement instead, change Bypass to RemoteSigned or AllSigned below - the
REM script is Authenticode-signed, so it will satisfy either.
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" %PSARGS%
set "RC=%errorlevel%"

echo.
if "%RC%"=="0" echo  RESULT: No changes were needed. This computer is fine.
if "%RC%"=="1" echo  RESULT: An error occurred. Please send the log file to IT.
if "%RC%"=="2" echo  RESULT: Administrator rights are required.
if "%RC%"=="3" echo  RESULT: Repair applied. The computer will restart.
if "%RC%"=="4" echo  RESULT: WHAT IF - this computer NEEDS the repair. Re-run without /whatif to apply it.
echo.
echo  Logs: C:\ProgramData\BriComp\DCSyncRootFix
echo.

REM Only pause when a human is watching and nothing is about to reboot
if not "%RC%"=="3" pause
exit /b %RC%
