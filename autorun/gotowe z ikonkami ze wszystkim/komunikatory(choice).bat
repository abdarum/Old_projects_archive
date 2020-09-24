@echo off
:4
echo -----KOMUNIKATORY-----
echo ----------------------
echo [ 1 ].Open Skype and GG
echo [ 2 ].Close Skype and GG
echo [ 3 ].End task
echo [ 4 ].Refresh
echo ----------------------


choice /C:1234
IF ERRORLEVEL 4 GOTO 4
IF ERRORLEVEL 3 GOTO 3
IF ERRORLEVEL 2 GOTO 2
IF ERRORLEVEL 1 GOTO 1


:1
start skype.exe
start C:\Users\klub\AppData\Local\GG\Application\gghub.exe
goto 3

:2
taskkill /t /f /im skype.exe 
taskkill /t /f /im gghub.exe
taskkill /t /f /im ggapp.exe
goto 3

:3
exit 