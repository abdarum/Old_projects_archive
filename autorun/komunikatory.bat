@echo off
title Komunikatory
echo Write start or kill. Start to open Skype and GG or kill to close Skype and GG
set/p "cho=>"
if %cho%==start goto start
if %cho%==kill goto kill

:start
start skype.exe
start C:\Users\klub\AppData\Local\GG\Application\gghub.exe
goto end

:kill

taskkill /t /f /im skype.exe 
taskkill /t /f /im ggapp.exe
taskkill /t /f /im gghub.exe
taskkill /t /f /im ggapp.exe
goto end
pause

:end
