@echo off

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

echo Delete files
echo y | del ..\disk\disk1\*.*

call buildlog %1 %2 %3

echo .
echo Compress updated application files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\BIN\*.* ..\disk\disk1\servbin.z 

echo .
echo Compress updated application files
icomp G:\program\release\exectrak\common\cpylocal.exe ..\disk\disk1\servbin.z 

echo .
echo Compress updated custctl files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\custctl\*.* ..\disk\disk1\servctl.z 

echo .
echo Compress Local Database
icomp G:\program\release\exectrak\local_db\factor.mdb ..\disk\disk1\localdb.z

echo .
echo Compress updated dll files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\dll\*.* ..\disk\disk1\servdll.z 

echo .
echo Compress updated ole files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\ole\*.* ..\disk\disk1\servole.z 

echo .
echo Compress updated rtm files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\rtm\*.* ..\disk\disk1\servrtm.z 

echo .
echo Compress updated shared files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\shared\*.* ..\disk\disk1\servshar.z 

echo .
echo Compress updated factor menu files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\*.* ..\disk\disk1\servfact.z 

echo .
echo Compress updated dtnctrl files
icomp G:\PROGRAM\RELEASE\Exectrak\SP_INTRM\DTNCTRL\*.* ..\disk\disk1\servdtnc.z 

echo .
echo Done

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Name] [Module Version] [Beta]
echo .

:finished

