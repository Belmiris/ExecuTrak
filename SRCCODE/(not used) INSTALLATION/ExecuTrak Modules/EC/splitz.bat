call makedir.bat

@echo off
echo split appfiles.z
call fsplit ..\disk\factbin.z ..\disk 700
move ..\disk\factbin.1 ..\disk\disk1
move ..\disk\factbin.2 ..\disk\disk2
move ..\disk\factbin.3 ..\disk\disk3
move ..\disk\factbin.4 ..\disk\disk3

echo .
REMecho split localdb.z
REMcall fsplit ..\disk\localdb.z ..\disk 9
REMmove ..\disk\localdb.1 ..\disk\disk3
REMmove ..\disk\localdb.2 ..\disk\disk4
copy localdb.z ..\disk\disk4

echo .
choice /c:yn /t:y,7 Do you want to Compile the Setup Script
if errorlevel 2 goto skipcompile
compile setup
echo .

:skipcompile

choice /c:yn /t:y,7 Do you want to Compile the Package List
if errorlevel 2 goto skippacking

packlist setup.lst
echo .

:skippacking
call copysetf.bat

echo .
echo split Z files has finished.

choice /c:yn /t:n,5 Do you want to make the Distribution Disk.

if errorlevel 2 goto finished

:makedisk
call makedisk

goto finished

echo .
echo split .z files has finished.  Run makedisk.bat to copy distribution disk.
goto finished

:finished
