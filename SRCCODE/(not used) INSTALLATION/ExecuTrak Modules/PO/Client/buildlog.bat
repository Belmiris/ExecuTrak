rem %3 = BETA or Blank
@echo off
echo .
echo Updating DiskBild database

if "%3" == "" goto Normal_Installation

K:\DISKBILD\DISKBILD.EXE BETAPO3 %2 MODFILES.LST

goto done

:Normal_Installation
K:\DISKBILD\DISKBILD.EXE PO3 %2 MODFILES.LST

goto done

:done