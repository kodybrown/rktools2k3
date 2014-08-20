@echo off
REM
REM Script to dump FSMO role owners on the server designated by %1
REM

if ""=="%1" goto usage

ntdsutil roles Connections "Connect to server %1" Quit "select Operation Target" "List roles for connected server" Quit Quit Quit 

goto done

:usage

@echo Please provide the name of a domain controller (i.e. dumpfsmos MYDC)
@echo.

:done
