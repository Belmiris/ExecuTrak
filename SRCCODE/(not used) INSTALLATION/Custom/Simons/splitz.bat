@echo off
call makedir.bat

echo NO split factbin.z
copy ..\disk\factbin.z ..\disk\disk1

echo NO split tlc.z
copy ..\disk\tlc.z ..\disk\disk1

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk1

echo split custctl.z
call fsplit ..\disk\custctl.z ..\disk 250
move ..\disk\custctl.1 ..\disk\disk1
move ..\disk\custctl.2 ..\disk\disk2

echo NO split factdll.z
copy ..\disk\factdll.z ..\disk\disk2

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 619
move ..\disk\rtm.1 ..\disk\disk2
move ..\disk\rtm.2 ..\disk\disk3
move ..\disk\rtm.3 ..\disk\disk4

echo NO split shared.z
copy ..\disk\shared.z ..\disk\disk4

REM echo NO split factdb.z
REM copy ..\disk\factdb.z ..\disk\disk4
echo split factdb.z
call fsplit ..\disk\factdb.z ..\disk 240
move ..\disk\factdb.1 ..\disk\disk4
move ..\disk\factdb.2 ..\disk\disk5


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
