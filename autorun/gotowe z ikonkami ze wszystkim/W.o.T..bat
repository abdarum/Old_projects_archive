@echo off
title W.o.T.
:8

cls
echo ----------------
echo -----W.o.T.-----
echo ----------------
echo ================
echo [ 1 ].Open W.o.T.
echo [ 2 ].Open W.o.T. with Luncher
echo [ 3 ].Kill W.o.T.
echo [ 4 ].Restart W.o.T.
echo [ 5 ].Exit
echo [ 6 ].Delete files: python.log and Influx_PS.log
echo [ 7 ].Open page W.o.T.
echo ================


CHOICE /C:1234567 /N

IF ERRORLEVEL 7 GOTO 7
IF ERRORLEVEL 6 GOTO 6
IF ERRORLEVEL 5 GOTO 5
IF ERRORLEVEL 4 GOTO 4
IF ERRORLEVEL 3 GOTO 3
IF ERRORLEVEL 2 GOTO 2
IF ERRORLEVEL 1 GOTO 1

:1
start C:\Games\World_of_Tanks\WorldOfTanks.exe
cls
goto 6

:2
start C:\Games\World_of_Tanks\WoTLauncher.exe
cls
goto 6

:3
taskkill /f /im WorldOfTanks.exe
cls
goto 6

:4
taskkill /f /im WorldOfTanks.exe
start C:\Games\World_of_Tanks\WorldOfTanks.exe
cls
goto 6
:5
exit

:6
timeout /t 40
goto 6.1

:6.1
CHOICE /C:12 /M "Wybierz cyfre 2 aby przejsc do Menu lub 1 aby kontynuowac." /T 3 /D 1

IF ERRORLEVEL 2 GOTO 8

if exist python.log goto 6.2

if exist Influx_PS.log goto 6.2

:6.2
del python.log
del Influx_PS.log
del python.log
del Influx_PS.log
del python.log
del Influx_PS.log
if not exist python.log goto 8
if not exist Influx_PS.log goto 8
goto 6.1

:7
start http://worldoftanks.eu/
goto 8

