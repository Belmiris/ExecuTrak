@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo.
echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat PROFITRPT %1

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z

icomp G:\PROGRAM\RELEASE\prft_rpt\*.* ..\disk\factbin.z -i
icomp G:\program\release\exectrak\common\*.exe ..\disk\factbin.z bin
icomp G:\program\release\exectrak\common\*.hlp ..\disk\factbin.z 

echo .
echo Compress all files in release32\exectrak\Crystal directory to custctl.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\crystal\*.* ..\disk\crystal.z

echo .
echo Compress all files in release32\exectrak\Custctl directory to custctl.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in release32\exectrak\Dll directory to factdll.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\dll\*.* ..\disk\factdll.z

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Compress all files in release32\exectrak\Ole directory to factole.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\ole\tbkit.dll ..\disk\factole.z
icomp G:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in release32\exectrak\Rtm directory to rtm.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in release32\exectrak\Shared directory to shared.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\shared\*.* ..\disk\shared.z


echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, FACTOLE.Z, CRYSTAL, CUSTCTL.Z, 
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
