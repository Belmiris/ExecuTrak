@echo off

if "%1" == "" goto parm_missing

rem CREATE SETUP.BMP AND CHANGE VERSION
call buildbmp.bat %1

echo Delete z files
del ..\disk\*.z

echo .
echo call DISKBILD.BAT
call buildlog.bat EXECVISION40DC %1

echo .
echo Compress Application files

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\program\release\execvisn\40d_cycl\bin\*.* ..\disk\factbin.z bin
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z bin

icomp G:\program\release\execvisn\40d_cycl\qrt_file\*.* ..\disk\factbin.z qrt_file

icomp G:\program\release\execvisn\40d_cycl\models\*.* ..\disk\factbin.z -i


echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, FACTDB.Z)

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
