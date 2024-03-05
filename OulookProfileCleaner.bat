@echo off

set outlookpst=C:\Users\%username%\AppData\Roaming\Microsoft\Outlook

REM check for ReportQueue folder, change directory to it if it does exist
if exist %outlookpst% (
	cd /d %outlookpst%
) Else (
	echo "Outlook PST folder not found."
)
	
REM doublecheck that system is currently in PST folder. Then preform changes.
if %cd%==%outlookpst% (
	cp -R %outlookpst% %outlookpst%\PSTbackup
	for /F "delims=" %%i in ('dir /b') do (del "%%i" /S/Q)
	echo "PST backup created."
) Else (
	echo "Current directory is not PST folder. Exiting..."
)
pause