@echo off
cls

if "%1" == "" goto parm_missing


:server_compress
cd .\clients\exectrak
echo .
choice /c:yn /t:y,7 Do you want to Compile the Client Setup Script
if errorlevel 2 goto nopackclient
compile setup
echo .

packlist setup.lst
echo .

:nopackclient
cd ..\..\
cd .\servers\exectrak
cd

echo .
choice /c:yn /t:y,7 Do you want to Compile the Server Setup Script
if errorlevel 2 goto nopackserver
compile setup
echo .

packlist setup.lst
echo .


:nopackserver
cd ..\..\
cd

rem COMPRESS/MAKE Z FILES (factmenu.z,sybin,localdb.z,crystal.z,
rem custctl,factdll,rtm,shared)

if "%2" == "ALL" goto cmpfiles_all
if "%2" == "all" goto cmpfiles_all
if "%2" == "" goto start_cmpfiles
call %2
call updatez .\clients\exectrak\zfiles\%2.z .\servers\exectrak\zfiles\client.z

if "%3" == "" goto start_cmpfiles
call %3
call updatez .\clients\exectrak\zfiles\%3.z .\servers\exectrak\zfiles\client.z

if "%4" == "" goto start_cmpfiles
call %4
call updatez .\clients\exectrak\zfiles\%4.z .\servers\exectrak\zfiles\client.z

if "%5" == "" goto start_cmpfiles
call %5
call updatez .\clients\exectrak\zfiles\%5.z .\servers\exectrak\zfiles\client.z

if "%6" == "" goto start_cmpfiles
call %6
call updatez .\clients\exectrak\zfiles\%6.z .\servers\exectrak\zfiles\client.z

if "%7" == "" goto start_cmpfiles
call %7
call updatez .\clients\exectrak\zfiles\%7.z .\servers\exectrak\zfiles\client.z

if "%8" == "" goto start_cmpfiles
call %8
call updatez .\clients\exectrak\zfiles\%8.z .\servers\exectrak\zfiles\client.z

goto start_cmpfiles

:cmpfiles_all
call factmenu
call sybin
call factdb
call crystal
call custctl
call factdll
call factole
call rtm
call shared
call updatez .\clients\exectrak\zfiles\*.* .\servers\exectrak\zfiles\client.z


:start_cmpfiles
echo .
echo 1 -- Update Disk Build Log database
REM call buildlog %1

call cupsetup

if "%2" == "" goto ask_copy_setupz


call splitz
goto ask_make_disk


:ask_copy_setupz
call copysetf

echo .
echo SETUP.Z file has updated.
choice /c:yn /t:y,5 Do you want to copy SETUP.Z to Disk#1.

if errorlevel 2 goto ask_make_disk

echo .
echo Insert Disk#1 into Drive A:
pause

echo a | xcopy .\disk\disk1\setup.z a:
echo a | xcopy .\disk\disk1\setup.bmp a:
echo a | xcopy .\disk\disk1\setup.pkg a:
echo a | xcopy .\servers\exectrak\setup.ins a:


:ask_make_disk
echo .
echo split Z files has finished.
choice /c:yn /t:y,5 Do you want to make the Distribution Disk.

if errorlevel 2 goto finished


:makedisk
call makedisk

goto finished


:parm_missing
echo .
echo Version No is missing.  format: excmp Version_No [factmenu] [sybin]
echo        [factdb] [crystal] [custctl] [factdll] [factole] [rtm] [shared]


:finished
echo .
echo Done
