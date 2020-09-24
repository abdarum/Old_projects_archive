@ECHO OFF
REM  QBFC Project Options Begin
REM  HasVersionInfo: No
REM  Companyname: 
REM  Productname: 
REM  Filedescription: 
REM  Copyrights: 
REM  Trademarks: 
REM  Originalname: 
REM  Comments: 
REM  Productversion:  0. 0. 0. 0
REM  Fileversion:  0. 0. 0. 0
REM  Internalname: 
REM  Appicon: ..\ikony\IKONY\Folders\Group; Public\Folder-Sharepoint.ico
REM  AdministratorManifest: No
REM  Embeddedfile: vbs\kill.vbe
REM  Embeddedfile: vbs\Komunikatory.vbe
REM  Embeddedfile: vbs\Komunikatory1.vbe
REM  Embeddedfile: vbs\open.vbe
REM  QBFC Project Options End
ECHO ON
@echo off
title Komunikatory
call %MYFILES%\Komunikatory.vbe
:if
set/p "cho=>"
if %cho%==start goto start
if %cho%==kill goto kill

call D:\autorun\vbs\Komunikatory1.vbe
goto if

:start
start skype.exe
start C:\Users\klub\AppData\Local\GG\Application\gghub.exe
call %MYFILES%\open.vbe
goto end 

:kill

taskkill /t /f /im skype.exe 
taskkill /t /f /im gghub.exe
taskkill /t /f /im ggapp.exe
call %MYFILES%\kill.vbe
goto end



:end
