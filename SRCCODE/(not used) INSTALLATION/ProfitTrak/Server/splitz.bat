@echo off

echo .
call makedir

echo  split ptclient.z
call fsplit ..\disk\ptclient.z ..\disk 700
move ..\disk\ptclient.1 ..\disk\disk1
move ..\disk\ptclient.2 ..\disk\disk2
move ..\disk\ptclient.3 ..\disk\disk3


echo  no split ptini.z
copy ..\disk\ptini.z ..\disk\disk3

echo  no split ptmodels.z
copy ..\disk\ptmodels.z ..\disk\disk3

echo  split ptiqbin.z
call fsplit ..\disk\ptiqbin.z ..\disk 800
move ..\disk\ptiqbin.1 ..\disk\disk3
move ..\disk\ptiqbin.2 ..\disk\disk4

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk4


echo  split custctl.z
call fsplit ..\disk\custctl.z ..\disk 340
move ..\disk\custctl.1 ..\disk\disk4
move ..\disk\custctl.2 ..\disk\disk5
move ..\disk\custctl.3 ..\disk\disk6

echo No split factdll.z
copy ..\disk\factdll.z ..\disk\disk6

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 30
move ..\disk\rtm.1 ..\disk\disk6
move ..\disk\rtm.2 ..\disk\disk7
move ..\disk\rtm.3 ..\disk\disk8
move ..\disk\rtm.4 ..\disk\disk9

echo  no split shared.z
copy ..\disk\shared.z ..\disk\disk9



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
