@echo off
rem Note: NEED TO copy the files to APPFILES\BIN directory from \PROGRAM\RELEASE\EPB
rem       Move the HELP FILE to APPFILES directory

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat EPB %1

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\program\release\exectrak\epb\*.* ..\disk\factbin.z -i
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z bin

echo .
echo Compress all files in Crystal directory to custctl.z
icomp G:\program\release\exectrak\crystal\*.* ..\disk\crystal.z

echo .
echo Compress all files in Ole directory to factole.z
icomp G:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp G:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in Dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* ..\disk\factdll.z


echo .
echo Compress all files in Rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in Shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* ..\disk\shared.z

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo Z File sequence (FACTBIN.Z, SHARED.Z, FACTDLL.Z, FACTOLE.Z, CUSTCTL.Z
echo                   , FACTDB.Z, CRYSTAL.Z, RTM.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.

echo .
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
