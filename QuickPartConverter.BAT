@echo off

set templatesDir=C:\Users\%USERNAME%\AppData\Roaming\Microsoft\Templates

cd %templatesDir%


REM if Pre15's exist, delete old ones and rename
IF EXIST NormalEmail15Pre.dotm (
	DEL /Q NormalEmail.dotm
	ren NormalEmail15Pre.dotm NormalEmail.dotm
) ELSE (
	echo "(Send this to your System Administrator) ERROR 2: You have no Outlook 2010 Quickparts to convert!"
	pause
	exit
)
pause