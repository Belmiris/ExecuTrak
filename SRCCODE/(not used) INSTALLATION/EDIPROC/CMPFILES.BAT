@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild

echo Delete z files
del ..\disk\*.z

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
echo Compress all files in G:\program\release\exectrak\ole directory to factole.z
icomp G:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp G:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in G:\program\release\exectrak\Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in G:\program\release\exectrak\dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* ..\disk\factdll.z

echo .
echo Compress all files in G:\program\release\exectrak\rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in G:\program\release\exectrak\shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* ..\disk\shared.z

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

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
