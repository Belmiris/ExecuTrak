@echo off

echo .
call makedir

echo  split ptclient.z
call fsplit ..\disk\ptclient.z ..\disk 700
move ..\disk\ptclient.1 ..\disk\disk1
move ..\disk\ptclient.2 ..\disk\disk2
move ..\disk\ptclient.3 ..\disk\disk3
move ..\disk\ptclient.4 ..\disk\disk4

echo  no split ptmodels.z
copy ..\disk\ptmodels.z ..\disk\disk4

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
