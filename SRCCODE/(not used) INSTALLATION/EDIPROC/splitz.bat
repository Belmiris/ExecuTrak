@echo off
call makedir.bat

echo split custctl.z
call fsplit ..\disk\ediproc.z ..\disk 700
move ..\disk\ediproc.1 ..\disk\disk1
move ..\disk\ediproc.2 ..\disk\disk2

echo NO split factole.z
copy ..\disk\factole.z ..\disk\disk2

echo split custctl.z
call fsplit ..\disk\custctl.z ..\disk 971
move ..\disk\custctl.1 ..\disk\disk2
move ..\disk\custctl.2 ..\disk\disk3

echo NO split factdll.z
copy ..\disk\factdll.z ..\disk\disk3

echo split rtm.z
call fsplit ..\disk\rtm.z ..\disk 201
move ..\disk\rtm.1 ..\disk\disk3
move ..\disk\rtm.2 ..\disk\disk4
move ..\disk\rtm.3 ..\disk\disk5
move ..\disk\rtm.4 ..\disk\disk6

echo NO split shared.z
copy ..\disk\shared.z ..\disk\disk6

echo No split factdb.z
copy ..\disk\factdb.z ..\disk\disk6

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
