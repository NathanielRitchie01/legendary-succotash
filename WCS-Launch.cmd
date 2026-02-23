@echo off
REM ============================================================
REM  WCS-Launch.cmd — Launcher for WCS-CHECKER.ps1
REM  Author: Nathaniel Ritchie | Site: AUKC01
REM ============================================================
REM
REM  USAGE:
REM    WCS-Launch.cmd              Launch normally (production)
REM    WCS-Launch.cmd /DEBUG       Launch with separate debug window
REM    WCS-Launch.cmd /TEST        Run automated test suite
REM    WCS-Launch.cmd /HELP        Show this help
REM
REM  NOTES:
REM    - /DEBUG opens a second console that live-tails the debug
REM      log. Close it anytime; it won't affect the main app.
REM    - /TEST runs WCS-TestFramework.ps1 with mock data (no DB).
REM      Returns exit code 0 on pass, 1 on failure.
REM    - All files must be in the same directory as this .cmd.
REM
REM ============================================================

setlocal enabledelayedexpansion

REM ---- Resolve script directory ----
set "SCRIPT_DIR=%~dp0"
set "MAIN_SCRIPT=%SCRIPT_DIR%WCS-CHECKER.ps1"
set "TEST_SCRIPT=%SCRIPT_DIR%WCS-TestFramework.ps1"
set "DEBUG_LOG=%SCRIPT_DIR%wcs-debug.log"

REM ---- Parse arguments ----
set "MODE=NORMAL"
if /i "%~1"=="/DEBUG" set "MODE=DEBUG"
if /i "%~1"=="/debug" set "MODE=DEBUG"
if /i "%~1"=="-debug" set "MODE=DEBUG"
if /i "%~1"=="debug"  set "MODE=DEBUG"

if /i "%~1"=="/TEST"  set "MODE=TEST"
if /i "%~1"=="/test"  set "MODE=TEST"
if /i "%~1"=="-test"  set "MODE=TEST"
if /i "%~1"=="test"   set "MODE=TEST"

if /i "%~1"=="/HELP"  set "MODE=HELP"
if /i "%~1"=="/help"  set "MODE=HELP"
if /i "%~1"=="-help"  set "MODE=HELP"
if /i "%~1"=="help"   set "MODE=HELP"
if /i "%~1"=="/?"     set "MODE=HELP"

REM ============================================================
REM  HELP
REM ============================================================
if "%MODE%"=="HELP" (
    echo.
    echo  ============================================================
    echo   WCS-Launch.cmd — Launcher for WCS-CHECKER
    echo  ============================================================
    echo.
    echo   USAGE:
    echo     WCS-Launch.cmd              Normal mode ^(production^)
    echo     WCS-Launch.cmd /DEBUG       Debug mode + live log window
    echo     WCS-Launch.cmd /TEST        Run automated tests
    echo     WCS-Launch.cmd /HELP        This help screen
    echo.
    echo   FILES REQUIRED ^(same directory^):
    echo     WCS-CHECKER.ps1             Main dashboard script
    echo     WCS-TestFramework.ps1       Test harness ^(for /TEST^)
    echo.
    echo   DEBUG MODE:
    echo     Opens a second console window that live-scrolls all
    echo     Write-Debug messages from the main application.
    echo     Close the debug window anytime — it does not affect
    echo     the main app. The log file is: wcs-debug.log
    echo.
    echo   TEST MODE:
    echo     Runs all automated tests with mock data ^(no DB needed^).
    echo     Returns exit code 0 if all pass, 1 if any fail.
    echo     Safe to run anytime — never touches production data.
    echo.
    echo  ============================================================
    echo.
    pause
    exit /b 0
)

REM ---- Verify files exist ----
if not exist "%MAIN_SCRIPT%" (
    echo.
    echo  [ERROR] Cannot find WCS-CHECKER.ps1
    echo  Expected at: %MAIN_SCRIPT%
    echo  Make sure all files are in the same directory.
    echo.
    pause
    exit /b 1
)

REM ============================================================
REM  TEST MODE
REM ============================================================
if "%MODE%"=="TEST" (
    if not exist "%TEST_SCRIPT%" (
        echo.
        echo  [ERROR] Cannot find WCS-TestFramework.ps1
        echo  Expected at: %TEST_SCRIPT%
        echo.
        pause
        exit /b 1
    )

    echo.
    echo  ============================================================
    echo   WCS-CHECKER TEST MODE
    echo  ============================================================
    echo   Running automated tests with mock data...
    echo   No database connection will be made.
    echo  ============================================================
    echo.

    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%TEST_SCRIPT%"

    set "TEST_EXIT=!errorlevel!"
    echo.
    if "!TEST_EXIT!"=="0" (
        echo  [RESULT] All tests PASSED.
    ) else (
        echo  [RESULT] Some tests FAILED. Review output above.
    )
    echo.
    pause
    exit /b !TEST_EXIT!
)


REM ============================================================
REM  DEBUG MODE
REM ============================================================
if "%MODE%"=="DEBUG" (
    echo.
    echo  ============================================================
    echo   WCS-CHECKER DEBUG MODE
    echo  ============================================================
    echo   Main app will launch with full debug tracing.
    echo   A separate window will show live debug output.
    echo  ============================================================
    echo.

    REM Clear previous debug log
    if exist "%DEBUG_LOG%" del "%DEBUG_LOG%"
    echo [%date% %time%] Debug session started > "%DEBUG_LOG%"

    REM Launch the debug tail window (separate process)
    REM Uses PowerShell's Get-Content -Wait to tail the log file
    start "WCS Debug Monitor" cmd /c "title WCS Debug Monitor - LIVE LOG & color 0A & echo. & echo  ============================================ & echo   WCS-CHECKER DEBUG MONITOR & echo   Watching: %DEBUG_LOG% & echo   Close this window anytime. & echo  ============================================ & echo. & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Get-Content -Path '%DEBUG_LOG%' -Wait -Tail 0""

    REM Give the tail window a moment to start
    timeout /t 2 /nobreak > nul

    REM Launch main app with debug preferences set
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
        "$DebugPreference = 'Continue'; $Script:DebugLogPath = '%DEBUG_LOG%'; $Script:TestMode = $false; & '%MAIN_SCRIPT%'"

    echo.
    echo  Main application exited. Debug log saved at:
    echo  %DEBUG_LOG%
    echo.
    pause
    exit /b 0
)


REM ============================================================
REM  NORMAL MODE (default)
REM ============================================================
echo.
echo  Starting WCS-CHECKER...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%MAIN_SCRIPT%"

exit /b 0