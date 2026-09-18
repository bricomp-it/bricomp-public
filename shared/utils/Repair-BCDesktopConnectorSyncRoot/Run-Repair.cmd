@echo off
REM ============================================================================
REM  Repair-BCDesktopConnectorSyncRoot - launcher
REM  BriComp IT Consulting Services   https://bricomp.com
REM  Version 1.0.0
REM
REM  HOW TO USE
REM    Right-click this file and choose "Run as administrator".
REM    Approve the Windows security prompt if it appears.
REM    The computer will restart automatically when the repair is done.
REM
REM    To preview without changing anything:   Run-Repair.cmd /dryrun
REM    To repair but not restart:              Run-Repair.cmd /norestart
REM
REM  Logs and the registry backup are written to:
REM    C:\ProgramData\BriComp\DCSyncRootFix
REM ============================================================================

setlocal
set "SCRIPTDIR=%~dp0"
set "PS1=%SCRIPTDIR%Repair-BCDesktopConnectorSyncRoot.ps1"
set "PSARGS="

if /i "%~1"=="/dryrun"    set "PSARGS=-DryRun"
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
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "Start-Process -FilePath '%~f0' -ArgumentList '%~1' -Verb RunAs"
    exit /b 0
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
echo.
echo  Logs: C:\ProgramData\BriComp\DCSyncRootFix
echo.

REM Only pause when a human is watching and nothing is about to reboot
if not "%RC%"=="3" pause
exit /b %RC%
