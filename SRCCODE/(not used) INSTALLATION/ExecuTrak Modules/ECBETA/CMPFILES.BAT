@echo off

if "%1" == "" goto parm_missing
if NOT "%2" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

del ..\disk\*.z

call buildlog ECBETA %1

echo Compress Application files
icomp G:\program\release\exectrak\ec-beta\bin\*.* ..\disk\factbin.z BIN
icomp G:\program\release\exectrak\ec-beta\ec.hlp ..\disk\factbin.z BIN
icomp G:\program\release\exectrak\ec-beta\dtnctrl\*.* ..\disk\factbin.z DTNCTRL
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\factor.mdb ..\disk\localdb.z

echo .
echo Finished making Z files.
echo .
echo Run the PACKCALC.EXE now to calculate the packing list.
echo And update the SPLITZ.BAT AND/OR SETUP.LST if necessary.
echo .
echo Z File sequence (APPFILES.Z, DTNCTRL.Z, LOCALDB.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.
echo .

echo .
choice /c:yn /t:y,7 Do you want to run the splitz.bat.

if errorlevel 2 goto finished

echo .
call splitz

goto finished


:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished

