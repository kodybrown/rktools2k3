@echo off
setlocal

set ServicePath=HKLM\System\CurrentControlSet\Services\Rqs
set EventPath=HKLM\System\CurrentControlSet\Services\EventLog\System\Rqs
set ServiceName=Remote Access Quarantine Agent

REM
REM EVENTLOG_SUCCESS                0x0000
REM EVENTLOG_ERROR_TYPE             0x0001
REM EVENTLOG_WARNING_TYPE           0x0002
REM EVENTLOG_INFORMATION_TYPE       0x0004
REM EVENTLOG_AUDIT_SUCCESS          0x0008
REM EVENTLOG_AUDIT_FAILURE          0x0010
REM

if "%1"=="/install" goto Install
if "%1"=="/remove" goto Remove
goto Usage

:Install
REM Copy the RQS binaries
copy /y rqs*.* %windir%\system32\ras

REM Register the service as auto start
sc.exe create rqs type= own start= auto binPath= "%windir%\System32\Ras\rqs.exe" depend= remoteaccess DisplayName= "%ServiceName%"

REM add failure actions to restart RQS should something unforseen occur
sc failure RQS reset= 86400 actions= restart/60000/restart/60000/restart/60000

REM Add the allowed version strings.  Note we have REM-ed out this line so it can be done manually.
REM Edit the following line with your version strings to make the batch setup fully automated.
REM REG ADD %ServicePath% /v AllowedSet /t REG_MULTI_SZ /d Version1\0Version1a\0Test

REM Setup the Event log messages
REG ADD %EventPath% /v EventMessageFile /t REG_EXPAND_SZ /d %windir%\System32\Ras\Rqsmsg.dll
REG ADD %EventPath% /v TypesSupported /t REG_DWORD /d 7

REM Start RQS
REM sc.exe start rqs

echo.
echo.
echo You must add your version string to the AllowedSet value of 
echo %ServicePath%
echo and then start the service using 'net start rqs'.  Or you can modify this
echo batch file to fully automate installing and configuring RQS.
echo.
echo.

goto Done

:Remove
net stop rqs
SC.exe DELETE rqs
REG DELETE %EventPath%
del %windir%\system32\ras\rqs*.*
goto Done

:Usage
echo %ServiceName% installation utility v1.0
echo ========================================================
echo To install:    rqs_setup /install
echo To remove:  rqs_setup /remove
goto Done

:Done
endlocal
echo on

