@echo off
echo .
echo Making Distribution Disk
echo .

echo Total Distribution Diskette Required = 1

echo .
echo insert disk#1
pause

deltree /Y a:\*.*
xcopy ..\disk\disk1\*.* a:

:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .
