@echo off

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

echo Delete z files
del ..\disk\disk1\*.z

call buildlog %1 %2

echo .
echo Compress Application files
icomp G:\PROGRAM\RELEASE\Exectrak\PT\*.* ..\disk\disk1\ptbin.z
icomp G:\PROGRAM\RELEASE\Exectrak\PT\BIN\*.* ..\disk\disk1\ptbin.z Bin
icomp G:\PROGRAM\RELEASE\Exectrak\PT\qrt_file\*.* ..\disk\disk1\ptbin.z Bin

icomp G:\PROGRAM\RELEASE\Exectrak\PT\ole\*.* ..\disk\disk1\ptbin.z ole
icomp G:\PROGRAM\RELEASE\Exectrak\PT\ini\*.* ..\disk\disk1\ptbin.z ini

icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\disk1\ptbin.z Bin

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

