rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z
del ..\disk\factdb.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat DutchOil3 %1

echo .
echo Compress all files in Release Source directory to factbin.z
icomp G:\PROGRAM\RELEASE\AutoSend\*.* ..\disk\factbin.z -i
icomp G:\program\release\exectrak\common\*.hlp ..\disk\factbin.z
icomp G:\program\release\exectrak\common\*.exe ..\disk\factbin.z BIN

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

rem echo .
rem echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
rem echo and then RUN the SPLITZ.BAT

echo .
echo RUN the SPLITZ.BAT
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
