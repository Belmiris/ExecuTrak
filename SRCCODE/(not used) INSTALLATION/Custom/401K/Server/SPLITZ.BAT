@echo off

echo .
call makedir

echo .
echo ******1 copy server setup files
call copysetf.bat 4k

echo  not split 4kclient.z

copy..\disk\4kclient.z ..\disk\disk1

echo  split shared.z
call fsplit ..\disk\shared.z ..\disk 170
move ..\disk\shared.1 ..\disk\disk1
move ..\disk\shared.2 ..\disk\disk2

echo No split factdll.z
copy ..\disk\factdll.z ..\disk\disk2

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk2

echo  split custctl.z
call fsplit ..\disk\custctl.z ..\disk 890
move ..\disk\custctl.1 ..\disk\disk2
move ..\disk\custctl.2 ..\disk\disk3

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 300
move ..\disk\rtm.1 ..\disk\disk3
move ..\disk\rtm.2 ..\disk\disk4
move ..\disk\rtm.3 ..\disk\disk5
move ..\disk\rtm.4 ..\disk\disk6

echo .
echo split .z files is finished.  Run makedisk.bat to copy distribution disk.

echo .
choice /c:yn /t:y,7 Do you want to Compile the Setup Script
if errorlevel 2 goto skipsetup
compile setup.rul

:skipsetup
echo .
choice /c:yn /t:y,7 Do you want to Compile the Setup.lst
if errorlevel 2 goto skipcompile
packlist setup.lst
echo .


:skipcompile
call copysetf.bat

echo .
call makedisk
