@echo off
setlocal
set "SERVICE_NAME=MSSG-Node"
sc start "%SERVICE_NAME%"
pause