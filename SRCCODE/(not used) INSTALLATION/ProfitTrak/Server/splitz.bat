@echo off

echo .
call makedir

echo  split ptclient.z
call fsplit ..\disk\ptclient.z ..\disk 700
move ..\disk\ptclient.1 ..\disk\disk1
move ..\disk\ptclient.2 ..\disk\disk2

echo  split shared.z
call fsplit ..\disk\shared.z ..\disk 170
move ..\disk\shared.1 ..\disk\disk2
move ..\disk\shared.2 ..\disk\disk3

echo No split factdll.z
copy ..\disk\factdll.z ..\disk\disk3

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk3

echo  split custctl.z
call fsplit ..\disk\custctl.z ..\disk 890
move ..\disk\custctl.1 ..\disk\disk3
move ..\disk\custctl.2 ..\disk\disk4

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 300
move ..\disk\rtm.1 ..\disk\disk4
move ..\disk\rtm.2 ..\disk\disk5
move ..\disk\rtm.3 ..\disk\disk6
move ..\disk\rtm.4 ..\disk\disk7

echo NO split crystal.z
copy ..\disk\crystal.z ..\disk\disk7

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
