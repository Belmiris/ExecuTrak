@echo off
call makedir.bat

echo NO split factbin.z
copy ..\disk\factbin.z ..\disk\disk1

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk2

echo split crystal.z
call fsplit ..\disk\crystal.z ..\disk 1250
move ..\disk\crystal.1 ..\disk\disk2
move ..\disk\crystal.2 ..\disk\disk3

echo split custctl.z
call fsplit ..\disk\custctl.z ..\disk 100
move ..\disk\custctl.1 ..\disk\disk3
move ..\disk\custctl.2 ..\disk\disk4
move ..\disk\custctl.3 ..\disk\disk5


echo NO split factdll.z
copy ..\disk\factdll.z ..\disk\disk5

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 710
move ..\disk\rtm.1 ..\disk\disk5
move ..\disk\rtm.2 ..\disk\disk6
move ..\disk\rtm.3 ..\disk\disk7
move ..\disk\rtm.4 ..\disk\disk8

echo  split shared.z
call fsplit ..\disk\shared.z ..\disk 150
move ..\disk\shared.1 ..\disk\disk7
move ..\disk\shared.2 ..\disk\disk8

echo NO split factdb.z
copy ..\disk\factdb.z ..\disk\disk8

echo .
echo split .z files is finished.  Run makedisk.bat to copy distribution disk.

echo .
choice /c:yn /t:y,7 Do you want to Compile the Setup Script
if errorlevel 2 goto skipcompile
compile setup
echo .
packlist setup.lst
echo .


:skipcompile
call copysetf.bat

echo .
call makedisk
