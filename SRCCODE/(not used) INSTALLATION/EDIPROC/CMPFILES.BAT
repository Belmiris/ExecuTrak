@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

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
