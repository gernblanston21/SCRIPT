@echo off
setlocal

set "SERVICE_NAME=MSSG-Node"

REM Try both common NSSM locations relative to this script
set "MSSG_DIR=%~dp0"
if "%MSSG_DIR:~-1%"=="\" set "MSSG_DIR=%MSSG_DIR:~0,-1%"
set "NSSM_EXE=%MSSG_DIR%\BIN\NSSM\nssm.exe"
if not exist "%NSSM_EXE%" set "NSSM_EXE=%MSSG_DIR%\bin\nssm.exe"

if exist "%NSSM_EXE%" (
  "%NSSM_EXE%" stop "%SERVICE_NAME%" >NUL 2>&1
  "%NSSM_EXE%" remove "%SERVICE_NAME%" confirm >NUL 2>&1
  if errorlevel 1 (
    echo [WARN] Service might not have existed or removal failed.
  ) else (
    echo [OK] %SERVICE_NAME% removed.
  )
) else (
  REM Fall back to sc if NSSM not present
  sc stop "%SERVICE_NAME%" >NUL 2>&1
  sc delete "%SERVICE_NAME%" >NUL 2>&1
  echo [INFO] NSSM not found. Attempted removal via SC. [cite: 10]
)

echo Done.
pause
exit /b 0