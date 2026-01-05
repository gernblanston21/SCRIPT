@echo off
setlocal
set "SERVICE_NAME=MSSG-Node"
sc stop "%SERVICE_NAME%"
pause