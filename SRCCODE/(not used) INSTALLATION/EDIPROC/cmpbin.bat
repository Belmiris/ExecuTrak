rem this batch file will compress EDIPROC.Z only
@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild

echo Delete z files
del ..\disk\ediproc.z

echo .
echo call DISKBILD.BAT
call buildlog.bat EDIPROC %1 %2

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to ediproc.z
icomp G:\program\release\ediproc\*.* ..\disk\ediproc.z -i
echo .

echo .
echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (EDIPROC.Z, FACTOLE.Z, CUSTCTL.Z,
echo                  FACTDLL.Z, RTM.Z, SHARED.Z, FACTDB.Z)

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
goto finished



:quitbuild
echo .
echo Disk Build Process has been cancelled.
echo .


:finished
