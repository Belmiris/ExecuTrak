@echo off
cls

echo .
echo makedir.bat
pause

deltree /Y ..\disk\disk1
md ..\disk\disk1
copy disk.id ..\disk\disk1\disk1.id
deltree /Y ..\disk\disk2
md ..\disk\disk2
copy disk.id ..\disk\disk2\disk2.id

echo .
echo cmpfiles.bat
pause

del ..\disk\appfiles.z
icomp ..\appfiles\*.* ..\disk\appfiles.z -i

echo .
echo splitz.bat
pause

call copysetf.bat
call fsplit ..\disk\appfiles.z ..\disk 635
move ..\disk\appfiles.1 ..\disk\disk1
move ..\disk\appfiles.2 ..\disk\disk2
copy ..\disk\custctl.z ..\disk\disk2
copy ..\disk\factdll.z ..\disk\disk2

echo .
echo makedisk.bat
echo .

echo Insert Setup Disk#1
pause
deltree /Y a:\*.*
xcopy ..\disk\disk1\*.* a:

echo Insert Setup Disk#2
pause
deltree /Y a:\*.* 
xcopy ..\disk\disk2\*.* a:

echo .
echo Note: changes has been made to appfiles.z 
echo Re-make setup disk#1 and disk#2 has finished.

