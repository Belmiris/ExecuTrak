@echo off
call makedir.bat

echo split factbin.z
call fsplit ..\disk\factbin.z ..\disk 690
move ..\disk\factbin.1 ..\disk\disk1
move ..\disk\factbin.2 ..\disk\disk2

echo no split factdb.z
copy ..\disk\factdb.z ..\disk\disk2

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
