rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat %1

echo .
echo Compress all files in appfiles directory to factbin.z
icomp ..\appfiles\*.* ..\disk\factbin.z -i

echo NO split factbin.z
copy ..\disk\factbin.z ..\disk\disk1

echo .
call makedisk

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
