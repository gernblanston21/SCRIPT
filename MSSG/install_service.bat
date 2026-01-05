@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==== Resolve MSSG folder (this .bat’s directory) ====
set "MSSG_DIR=%~dp0"
REM Strip trailing backslash
if "%MSSG_DIR:~-1%"=="\" set "MSSG_DIR=%MSSG_DIR:~0,-1%"

REM ==== Service name & display ====
set "SERVICE_NAME=MSSG-Node"
set "DISPLAY_NAME=MSSG Node Odds Cacher"
set "DESCRIPTION=Fetches FanDuel odds & props to MSSG\cache for Viz Trio."

REM ==== Find NSSM (prefer MSSG\BIN\NSSM\nssm.exe, else MSSG\bin\nssm.exe) ====
set "NSSM_EXE=%MSSG_DIR%\BIN\NSSM\nssm.exe"
if not exist "%NSSM_EXE%" set "NSSM_EXE=%MSSG_DIR%\bin\nssm.exe"
if not exist "%NSSM_EXE%" (
  echo [ERROR] nssm.exe not found. Expected:
  echo   %MSSG_DIR%\BIN\NSSM\nssm.exe  or  %MSSG_DIR%\bin\nssm.exe
  echo Please place NSSM there and re-run.
  exit /b 1
)

REM ==== Find Node (prefer bundled MSSG\BIN\Node\node.exe, else PATH) ====
set "NODE_EXE=%MSSG_DIR%\BIN\Node\node.exe"
if not exist "%NODE_EXE%" set "NODE_EXE=%MSSG_DIR%\bin\node\node.exe"
if not exist "%NODE_EXE%" (
  for /f "usebackq delims=" %%A in (`where node 2^>NUL`) do (
    set "NODE_EXE=%%A"
    goto :found_node
  )
  echo [ERROR] node.exe not found.
  echo Place a Node runtime at %MSSG_DIR%\BIN\Node\node.exe or ensure Node is on PATH.
  exit /b 1
)
:found_node

REM ==== Ensure cache\logs exists ====
if not exist "%MSSG_DIR%\cache" mkdir "%MSSG_DIR%\cache" >NUL 2>&1
if not exist "%MSSG_DIR%\cache\logs" mkdir "%MSSG_DIR%\cache\logs" >NUL 2>&1

REM ==== Ensure .env exists (warn only) ====
if not exist "%MSSG_DIR%\.env" (
  echo [WARN] .env not found at %MSSG_DIR%\.env
  echo Create it with:  FD_API_KEY=YOUR_KEY_HERE
)

REM ==== If service exists, stop & remove silently (idempotent install) ====
"%NSSM_EXE%" stop  "%SERVICE_NAME%" >NUL 2>&1
"%NSSM_EXE%" remove "%SERVICE_NAME%" confirm >NUL 2>&1

REM ==== Install the service ====
"%NSSM_EXE%" install "%SERVICE_NAME%" "%NODE_EXE%" "%MSSG_DIR%\server.js"
if errorlevel 1 (
  echo [ERROR] Failed to install service via NSSM.
  exit /b 1
)

REM ==== Set working dir and I/O redirection ====
"%NSSM_EXE%" set "%SERVICE_NAME%" AppDirectory "%MSSG_DIR%"
"%NSSM_EXE%" set "%SERVICE_NAME%" AppStdout    "%MSSG_DIR%\cache\logs\service.out.log"
"%NSSM_EXE%" set "%SERVICE_NAME%" AppStderr    "%MSSG_DIR%\cache\logs\service.err.log"

REM ==== Keep logs small: rotate when file > 5 MB, rotate online ====
REM NSSM rotates automatically; no watcher process needed.
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateBytes 5242880
"%NSSM_EXE%" set "%SERVICE_NAME%" AppRotateOnline 1

REM ==== Removed unsupported AppKillProcessTree line ====

REM ==== Set start mode, display name, description ====
"%NSSM_EXE%" set "%SERVICE_NAME%" Start SERVICE_AUTO_START
"%NSSM_EXE%" set "%SERVICE_NAME%" DisplayName "%DISPLAY_NAME%"
"%NSSM_EXE%" set "%SERVICE_NAME%" Description "%DESCRIPTION%"

REM ==== Start the service ====
"%NSSM_EXE%" start "%SERVICE_NAME%"

echo.
echo [OK] %DISPLAY_NAME% installed and started.
echo Logs: %MSSG_DIR%\cache\logs\service.out.log  /  service.err.log  (auto-rotated at 5 MB)
echo Working dir: %MSSG_DIR%
echo Using Node:  %NODE_EXE%
echo.

pause

exit /b 0