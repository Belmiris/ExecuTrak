@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo .
echo call DISKBILD.BAT
call buildlog.bat RPPCA %1

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild

echo Delete z files
del ..\disk\*.z

echo .
echo Compress all files in g:\program\release\execlink
echo and cpylocal.exe to factbin.z
icomp g:\program\release\execlink\*.* ..\disk\factbin.z -i
icomp g:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN

echo .
echo Compress 2 files in g:\program\release\exectrak\ole directory to factole.z
icomp g:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp g:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in g:\program\release\exectrak\custctl directory to custctl.z
icomp g:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in g:\program\release\exectrak\dll directory to factdll.z
icomp g:\program\release\exectrak\dll\*.* ..\disk\factdll.z

echo .
echo Compress all files in Rtm directory to rtm.z
icomp g:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in Shared directory to shared.z
icomp g:\program\release\exectrak\shared\*.* ..\disk\shared.z

echo .
echo Compress Local Database
icomp g:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

goto showinfo


:quitbuild
echo .
echo DISK BUILD PROCESS has been cancelled

goto finished

:showinfo
echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo Z File sequence (FACTBIN.Z, SHARED.Z, FACTDLL.Z, FACTOLE.Z, CUSTCTL.Z
echo                   , RTM.Z, FACTDB.Z)

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
