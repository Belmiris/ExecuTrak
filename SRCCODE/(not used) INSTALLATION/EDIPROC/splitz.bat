@echo off
call makedir.bat

echo split custctl.z
call fsplit ..\disk\ediproc.z ..\disk 700
move ..\disk\ediproc.1 ..\disk\disk1
move ..\disk\ediproc.2 ..\disk\disk2
move ..\disk\ediproc.3 ..\disk\disk3
move ..\disk\ediproc.4 ..\disk\disk4
move ..\disk\ediproc.5 ..\disk\disk5
move ..\disk\ediproc.6 ..\disk\disk6

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
