@echo off
rem Note: NEED TO copy the files to APPFILES\BIN directory from \PROGRAM\RELEASE\SIMONS
rem       Move the HELP FILE to APPFILES directory

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat SIMONS %1

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\PROGRAM\RELEASE\Custom\Simons\*.* ..\disk\factbin.z
icomp g:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN
icomp G:\PROGRAM\RELEASE\Custom\Simons\BIN\*.* ..\disk\factbin.z BIN

echo .
echo Compress all files in TLC directory to tlc.z
icomp g:\program\release\Custom\Simons\TLC\*.* ..\disk\tlc.z

echo .
echo Compress all files in g:\program\release\exectrak\Custctl directory to custctl.z
icomp g:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in g:\program\release\exectrak\pdqcomm directory to custctl.z
icomp g:\program\release\exectrak\pdqcomm\*.* ..\disk\custctl.z

echo .
echo Compress all files in g:\program\release\exectrak\dll directory to factdll.z
icomp g:\program\release\exectrak\dll\*.* ..\disk\factdll.z

echo .
echo Compress Local Database
icomp g:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Compress all files in g:\program\release\exectrak\ole directory to factole.z
icomp g:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp g:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in g:\program\release\exectrak\rtm directory to rtm.z
icomp g:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in g:\program\release\exectrak\shared directory to shared.z
icomp g:\program\release\exectrak\shared\*.* ..\disk\shared.z


:quitbuild
echo .
echo DISK BUILD PROCESS has been cancelled

goto finished

:showinfo
echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, TLC.Z ,FACTOLE.Z, CUSTCTL.Z,
echo                  FACTDLL.Z, RTM.Z, SHARED.Z, FACTDB.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.

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
