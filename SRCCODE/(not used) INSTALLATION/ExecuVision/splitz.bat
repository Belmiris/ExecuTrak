@echo off
call makedir.bat

echo split factbin.z
call fsplit ..\disk\factbin.z ..\disk 700
move ..\disk\factbin.1 ..\disk\disk1
move ..\disk\factbin.2 ..\disk\disk2
move ..\disk\factbin.3 ..\disk\disk3

echo no split factole.z
copy ..\disk\factole.z ..\disk\disk3

echo split custctl.z
call fsplit ..\disk\custctl.z ..\disk 880
move ..\disk\custctl.1 ..\disk\disk3
move ..\disk\custctl.2 ..\disk\disk4

echo split factdll.z
call fsplit ..\disk\factdll.z ..\disk 310
move ..\disk\factdll.1 ..\disk\disk4
move ..\disk\factdll.2 ..\disk\disk5

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 570
move ..\disk\rtm.1 ..\disk\disk5
move ..\disk\rtm.2 ..\disk\disk6
move ..\disk\rtm.3 ..\disk\disk7

echo NO split shared.z
copy ..\disk\shared.z ..\disk\disk8

echo no split factdb.z
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
