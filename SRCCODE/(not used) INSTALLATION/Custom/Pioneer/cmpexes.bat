rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat PIONEER %1

echo .
echo Compress all files in appfiles directory to factbin.z
icomp ..\appfiles\*.* ..\disk\factbin.z -i

rem echo .
rem echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
rem echo and then RUN the SPLITZ.BAT
echo .
echo Z File sequence (FACTBIN.Z, CRYSTAL.Z, FACTDB.Z)

echo .
echo RUN the SPLITZ.BAT
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
