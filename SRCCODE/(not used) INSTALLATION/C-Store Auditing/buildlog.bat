@echo off
echo .
echo Updating DiskBild database
if "%3" == "b" goto Beta
if "%3" == "B" goto Beta
K:\DISKBILD\DISKBILD.EXE %1 %2 csafiles.LST
goto done

:Beta
K:\DISKBILD\DISKBILD.EXE BETA%1 %2 csafiles.LST
:done