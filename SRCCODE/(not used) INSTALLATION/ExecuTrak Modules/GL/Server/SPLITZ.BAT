@echo off
cls

if "%1" == "" goto parm_missing
if "%2" == "" goto parm_missing

cd
echo .

call .\makedir

echo .
echo ******1 copy server setup files
call copysetf.bat %1

echo .
echo ******2 split client.z
call fsplit ..\disk\%1client.z ..\disk %2
if not exist ..\disk\%1client.1 goto copyfile
move ..\disk\%1client.1 ..\disk\disk1
if not exist ..\disk\%1client.2 goto splitdone
move ..\disk\%1client.2 ..\disk\disk2
if not exist ..\disk\%1client.3 goto splitdone
move ..\disk\%1client.3 ..\disk\disk3
if not exist ..\disk\%1client.4 goto splitdone
move ..\disk\%1client.4 ..\disk\disk4
if not exist ..\disk\%1client.5 goto splitdone
move ..\disk\%1client.5 ..\disk\disk5

:copyfile
copy ..\disk\%1client.z ..\disk\disk1

:splitdone

echo .
echo split Z files has finished.
choice /c:yn /t:n,5 Do you want to make the Distribution Disk.

if errorlevel 2 goto end

:makedisk
call makedisk

goto finished

:parm_missing
echo .
echo missing parameter.  format: splitz [AppName] [SplitSize].
echo AppName - AP, AR, GL, and etc.
echo SplitSize - Size of %1client.1 in bytes that will go to disk#1
echo             For example, 724
echo .

goto finished

:end

echo .
echo split .z files has finished.  Run makedisk.bat to copy distribution disk.
goto finished

:finished


