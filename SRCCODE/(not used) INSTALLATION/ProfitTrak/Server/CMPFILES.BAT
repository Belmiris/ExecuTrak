rem Assume: current directory is \Project6\xx\xxserver\setup

@echo off
cls

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing
if "%3" == "" goto parm_missing



if exist ..\..\%2client\setup\nul goto begin

echo .
echo Path:"%2modu~1\%2client\setup\" not found, cmpfiles failed.
echo .
goto finished


:begin

echo .

copy setup.bmp ..\..\%2client\setup

cd 
echo .

if "%1" == "c" goto client_compress
if "%1" == "C" goto client_compress
if "%1" == "b" goto client_compress
if "%1" == "B" goto client_compress

if "%1" == "s" goto server_compressor
if "%1" == "S" goto server_compress

goto parm_missing

:client_compress

if exist ..\..\%2client\disk\disk1\nul goto del_client_zfiles

if exist ..\..\%2client\disk\nul goto check_disk1

echo .
echo Path:"..\..\%2client\disk" not found, create a new one.
mkdir ..\..\%2client\disk

:check_disk1
mkdir ..\..\%2client\disk\disk1

:del_client_zfiles

echo .
echo ******1 Delete all %2 Client z Files
echo .
echo ******2 Compress %2 Application and database files

cd ..\..\%2client\setup
call makedir
call cmpfiles %2 %3

cd ..\..\%2server\setup

echo .
echo 2x -- Compress %2 Client files has finished.

if "%1" == "c" goto info
if "%1" == "C" goto info


:server_compress

if exist ..\disk\nul goto del_server_zfiles

echo .
echo Path:"..\disk" not found, make a new one.
mkdir ..\disk

:del_server_zfiles
call .\makedir

echo .
echo ******1 Delete all %2 Server z Files
del ..\disk\*.z

echo .
echo ******2 Compress client setup files to %2client.z
icomp ..\..\%2client\disk\disk1\*.* ..\disk\%2client.z

echo .
echo Compress all files in Ole directory to factole.z
icomp G:\program\release\exectrak\ole\tbkit.dll ..\disk\factole.z
icomp G:\program\release\exectrak\ole\olecombo.dll ..\disk\factole.z

echo .
echo Compress all files in Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* ..\disk\custctl.z

echo .
echo Compress all files in Dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* ..\disk\factdll.z

echo .
echo Compress all files in Rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* ..\disk\rtm.z

echo .
echo Compress all files in Shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* ..\disk\shared.z

echo Compress all files in Crystal directory to crystal.z
icomp G:\program\release\exectrak\crystal\*.* ..\disk\crystal.z


if "%1" == "b" goto compress_both
if "%1" == "B" goto compress_both

echo .
echo 2x -- Compress Server files has finished.

goto info

:compress_both
echo .
echo Client and Server Installation Files Compression has completed.

:info
echo .

if "%4" == "" goto split_default

echo .
echo call splitz with %2 %4
call splitz %2 %4
goto finished

:split_default
echo .
echo call splitz with %2 700
call splitz %2 
goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [C,S,B] [AppName] [Module Version] [optional SplitSize].
echo C - Compress Client Files Only.
echo S - Compress Server Files Only.
echo B - Compress Both Client and Server Files.
echo AppName - AP, AR, GL, and etc.
echo Module Version - 1.20.
echo SplitSize - this parameter is optional. (Default = 700).
echo .

:finished
