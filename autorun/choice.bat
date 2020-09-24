@ ECHO off
echo Write 1, 2 or 3
 @echo off
 @ CHOICE /C:123 
 IF ERRORLEVEL 3 GOTO trzy
 IF ERRORLEVEL 2 GOTO dwa
 IF ERRORLEVEL 1 GOTO jeden
 GOTO koniec
:jeden 
ECHO Naciœniêto "1" ! 
GOTO END
:dwa 
ECHO Naciœniêto "2" ! 
GOTO koniec
:trzy 
ECHO Naciœniêto "3" ! 

:koniec 
PAUSE