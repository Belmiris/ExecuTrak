@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat EXECVISION %1

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\program\release\execvisn\bin\*.* ..\disk\factbin.z bin
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z bin
icomp G:\program\release\execvisn\ini\*.* ..\disk\factbin.z ini
icomp G:\program\release\execvisn\ole\*.* ..\disk\factbin.z ole
icomp G:\program\release\execvisn\qrt_file\*.* ..\disk\factbin.z qrt_file
icomp G:\program\release\execvisn\registry\*.* ..\disk\factbin.z registry

icomp G:\program\release\execvisn\models\finance\*.* ..\disk\factbin.z finance
icomp G:\program\release\execvisn\models\fuel\*.* ..\disk\factbin.z fuel
icomp G:\program\release\execvisn\models\fueld\*.* ..\disk\factbin.z fueld
icomp G:\program\release\execvisn\models\retail\*.* ..\disk\factbin.z retail
icomp G:\program\release\execvisn\models\retaild\*.* ..\disk\factbin.z retaild
icomp G:\program\release\execvisn\models\sales\*.* ..\disk\factbin.z sales

echo .
echo Compress all files in g:\PROGRAM\RELEASE\EXECTRAK\Custctl directory to custctl.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in g:\PROGRAM\RELEASE\EXECTRAK\Dll directory to factdll.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\dll\*.* ..\disk\factdll.z
echo .
echo Compress all files in \EXECVISN\Dll directory to factdll.z
icomp G:\program\release\execvisn\dll\*.* ..\disk\factdll.z

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Compress all files in g:\PROGRAM\RELEASE\EXECTRAK\Ole directory to factole.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\ole\tbkit.dll ..\disk\factole.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in g:\PROGRAM\RELEASE\EXECTRAK\Rtm directory to rtm.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in g:\PROGRAM\RELEASE\EXECTRAK\Shared directory to shared.z
icomp G:\PROGRAM\RELEASE\EXECTRAK\shared\*.* ..\disk\shared.z


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
goto finished


:quitbuild
echo .
echo Disk Build Process has been cancelled.
echo .


:finished
