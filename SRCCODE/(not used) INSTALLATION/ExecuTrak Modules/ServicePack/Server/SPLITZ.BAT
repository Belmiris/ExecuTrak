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
if not exist ..\disk\%1client.6 goto splitdone
move ..\disk\%1client.6 ..\disk\disk6
if not exist ..\disk\%1client.7 goto splitdone
move ..\disk\%1client.7 ..\disk\disk7
if not exist ..\disk\%1client.8 goto splitdone
move ..\disk\%1client.8 ..\disk\disk8
if not exist ..\disk\%1client.9 goto splitdone
move ..\disk\%1client.9 ..\disk\disk9
if not exist ..\disk\%1client.10 goto splitdone
move ..\disk\%1client.10 ..\disk\disk10
if not exist ..\disk\%1client.11 goto splitdone
move ..\disk\%1client.11 ..\disk\disk11
if not exist ..\disk\%1client.12 goto splitdone
move ..\disk\%1client.12 ..\disk\disk12
if not exist ..\disk\%1client.13 goto splitdone
move ..\disk\%1client.13 ..\disk\disk13
if not exist ..\disk\%1client.14 goto splitdone
move ..\disk\%1client.14 ..\disk\disk14
if not exist ..\disk\%1client.15 goto splitdone
move ..\disk\%1client.15 ..\disk\disk15

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


