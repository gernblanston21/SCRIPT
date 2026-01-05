@echo off
setlocal EnableExtensions

set "MSSG_DIR=%~dp0"
if "%MSSG_DIR:~-1%"=="\" set "MSSG_DIR=%MSSG_DIR:~0,-1%"

set "NODE_EXE=%MSSG_DIR%\BIN\Node\node.exe"
if not exist "%NODE_EXE%" set "NODE_EXE=%MSSG_DIR%\bin\node\node.exe"
if not exist "%NODE_EXE%" set "NODE_EXE=node"

pushd "%MSSG_DIR%"
echo Using Node: %NODE_EXE%
echo Working dir: %CD%
echo.
"%NODE_EXE%" "%MSSG_DIR%\server.js"
popd
pause