@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat CSTORMNT %1 %2

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\program\release\cstormnt\*.* ..\disk\factbin.z -i
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN
echo .
echo Compress all files in project32\Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in project32\Dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* ..\disk\factdll.z

echo .
echo Compress all files in project32\Ole directory to factole.z
icomp G:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp G:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in project32\Rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in project32\Shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* ..\disk\shared.z

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, FACTOLE.Z, CUSTCTL.Z,
echo                  FACTDLL.Z, RTM.Z, SHARED.Z, FACTDB.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.

echo .
rem call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
