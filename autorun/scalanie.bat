@echo off
cls
set /p pathh= Sciezka:
set /p firstt= Plik1:
set /p secc= Plik2:
set /p outt= Wyjscie:
pushd %pathh%
copy /B %firstt% + %secc% %outt% >log.txt
pause
poupd