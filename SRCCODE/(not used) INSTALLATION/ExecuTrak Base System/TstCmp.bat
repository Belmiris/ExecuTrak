rem Assume: current directory is \release32\exectrak\setup

@echo off
cls

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

if "%1" == "c" goto client_compress
if "%1" == "C" goto client_compress
if "%1" == "b" goto client_compress
if "%1" == "B" goto client_compress

if "%1" == "s" goto server_compress
if "%1" == "S" goto server_compress


:client_compress

echo .
echo 0 -- Delete all Client z Files
deltree /Y .\clients\exectrak\zfiles
mkdir .\clients\exectrak\zfiles

call buildbmp.bat %2

echo .
echo 1 -- Update Disk Build Log database
call tsbuldlog %2 %3

echo .
echo 2 -- Compress Client Files and put them in Zfiles directory

echo .
echo 2a -- Compress all files in bin directory to sybin.z
icomp G:\program\release\exectrak\sy\bin\*.* .\clients\exectrak\zfiles\sybin.z
icomp G:\program\release\exectrak\common\*.exe .\clients\exectrak\zfiles\sybin.z

echo .
echo 2b -- Compress all files in Crystal directory to crystal.z
icomp G:\program\release\exectrak\crystal\*.* .\clients\exectrak\zfiles\crystal.z

echo .
echo 2c -- Compress all files in Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* .\clients\exectrak\zfiles\custctl.z

echo .
echo 2d -- Compress all files in Dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* .\clients\exectrak\zfiles\factdll.z

echo .
echo 2e -- Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* .\clients\exectrak\zfiles\localdb.z

echo .
echo 2f -- Compress all files in Ole directory to factole.z
icomp G:\program\release\exectrak\ole\*.* .\clients\exectrak\zfiles\factole.z

echo .
echo 2g -- Compress all files in Rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* .\clients\exectrak\zfiles\rtm.z

echo .
echo 2h -- Compress all files in parent directory to factmenu.z
icomp G:\program\release\exectrak\sy\*.* .\clients\exectrak\zfiles\factmenu.z
icomp G:\program\release\exectrak\common\*.hlp .\clients\exectrak\zfiles\factmenu.z

echo .
echo 2i -- Compress all files in Shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* .\clients\exectrak\zfiles\shared.z

echo .
echo 2x -- Compress Client files has finished.

rem if compress client files then continue to compress server files


:server_compress
cd .\clients\exectrak
echo .
choice /c:yn /t:y,7 Do you want to Compile the Client Setup Script
if errorlevel 2 goto nocompclient
compile setup
echo .

:nocompclient

choice /c:yn /t:y,7 Do you want to Compile the Client Package List
if errorlevel 2 goto nopackclient
echo .
packlist setup.lst
echo .

:nopackclient
cd ..\..\
cd .\servers\exectrak
cd

echo .
choice /c:yn /t:y,7 Do you want to Compile the Server Setup Script
if errorlevel 2 goto nocompserver
compile setup
echo .

:nocompserver

choice /c:yn /t:y,7 Do you want to Compile the Server Package List
if errorlevel 2 goto nopackserver
echo .
packlist setup.lst
echo .

:nopackserver

cd ..\..\
cd

echo .
echo 3 -- Delete all Server z Files
deltree /Y .\servers\exectrak\zfiles
mkdir .\servers\exectrak\zfiles

echo .
echo 4 -- Compress Server Files and put them in Zfiles directory

echo .
echo 4a -- Compress all files in clients\exectrak\zfiles directory to client.z
icomp .\clients\exectrak\zfiles\*.* .\servers\exectrak\zfiles\client.z

echo .
echo 4b -- Compress all 'program' files in c:\template\setup directory to setup.z
icomp c:\template\setup\*.* .\servers\exectrak\zfiles\setup.z

echo .
echo 4c -- Add clients\exectrak\setup.ins to setup.z
icomp .\clients\exectrak\setup.ins .\servers\exectrak\zfiles\setup.z

echo .
echo 4d -- Add setup.bmp to setup.z
icomp setup.bmp .\servers\exectrak\zfiles\setup.z

echo .
echo 4e -- Add clients\exectrak\setup.pkg to setup.z
icomp .\clients\exectrak\setup.pkg .\servers\exectrak\zfiles\setup.z

if "%1" == "b" goto compress_both
if "%1" == "B" goto compress_both

echo .
echo 4x -- Compress Server files has finished.

goto info

:compress_both
echo .
echo Client and Server Installation Files Compression has completed.

:info
echo .
echo Now, do the packing list Calculation and edit splitz.bat.  Then run splitz.bat
echo .
choice /c:yn /t:y,7 Do you want to run the splitz.bat.

if errorlevel 2 goto finished

call splitz

echo .
echo split Z files has finished.
choice /c:yn /t:y,5 Do you want to make the Distribution Disk.

if errorlevel 2 goto finished

:makedisk
call makedisk


goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [C,S,B] [Version].
echo C - Compress Both Client and Server Files.
echo S - Compress Server Files Only.
echo B - Compress Both Client and Server Files.
echo Version - 1.00, 1.10, 1.20 ....
echo b - BETA
echo .

:finished
echo .
echo Done
