@echo off

if "%1" == "" goto parm_missing

echo .
echo Updating DiskBild database
if "%2" == "" goto Normal_Installation
K:\DISKBILD\DISKBILD.EXE BETAVB6EXECUTRAK %1 EXECTRAK.LST

goto done

:Normal_Installation
K:\DISKBILD\DISKBILD.EXE VB6EXECUTRAK %1 EXECTRAK.LST

:done

echo .
echo Done

goto finished


:parm_missing
echo .
echo Version No is missing.  format: tsbldlog [Version_No] [b]


:finished

