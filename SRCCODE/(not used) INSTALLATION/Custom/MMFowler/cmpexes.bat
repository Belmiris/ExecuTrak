@echo off
rem Note: NEED TO copy the files to APPFILES\BIN directory from \PROGRAM\RELEASE\CSTORMNT
rem       Move the HELP FILE to APPFILES directory

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z file
del ..\disk\factbin.z

echo .
echo call DISKBILD.BAT
call buildlog.bat MMFOWLER %1

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp g:\program\release\Custom\MMFOWLER\*.* ..\disk\factbin.z -i

icomp g:\program\release\exectrak\common\*.hlp ..\disk\factbin.z
icomp g:\PROGRAM\release\exectrak\common\*.exe ..\disk\factbin.z BIN

goto showinfo

:quitbuild
echo .
echo DISK BUILD PROCESS has been cancelled

goto finished

:showinfo
echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, FACTOLE.Z, CUSTCTL.Z,
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
