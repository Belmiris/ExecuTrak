rem %3 = BETA or Blank
@echo off
echo .
echo Updating DiskBild database

if "%3" == "" goto Normal_Installation

K:\DISKBILD\DISKBILD.EXE BETA401K3 %2 MODFILES.LST

goto done

:Normal_Installation
K:\DISKBILD\DISKBILD.EXE 401K3 %2 MODFILES.LST

goto done

:done