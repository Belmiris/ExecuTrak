@echo off
echo .
echo Making Distribution Disk
echo .

echo Total Distribution Diskette Required = 2

echo .
echo insert disk#1
pause

deltree /Y a:\*.*
xcopy ..\disk\disk1\*.* a:

echo .
echo insert disk#2
pause

deltree /Y a:\*.*
xcopy ..\disk\disk2\*.* a:

:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .
