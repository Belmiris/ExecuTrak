@echo off

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

echo Delete files
deltree /Y ..\disk\disk1\*.*

call buildlog %1 %2 %3

echo .
echo Compress Application files
icomp G:\PROGRAM\RELEASE\Exectrak\%1\*.* ..\disk\disk1\%1bin.z -i
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\disk1\%1bin.z BIN

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\factor.mdb ..\disk\disk1\localdb.z

echo .
echo Done

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Name] [Module Version] [Beta]
echo .

:finished

