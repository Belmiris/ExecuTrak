@echo off

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

echo Delete z files
del ..\disk\disk1\*.z

call buildlog %1 %2

echo .
echo Compress Application files
icomp G:\program\release\exectrak\cm\*.* ..\disk\disk1\cmbin.z -i
icomp g:\program\release\exectrak\common\*.* ..\disk\disk1\cmbin.z BIN

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* ..\disk\disk1\localdb.z

echo .
echo Done

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Name] [Module Version]
echo .

:finished

