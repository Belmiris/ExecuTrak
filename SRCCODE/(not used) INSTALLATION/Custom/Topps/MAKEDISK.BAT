@echo off
echo .
echo Making Distribution Disk
echo .

if exist ..\disk\disk8\factdb.z goto disk8
if exist ..\disk\disk7\factdb.z goto disk7
if exist ..\disk\disk6\factdb.z goto disk6
if exist ..\disk\disk5\factdb.z goto disk5
if exist ..\disk\disk4\factdb.z goto disk4
if exist ..\disk\disk3\factdb.z goto disk3
if exist ..\disk\disk2\factdb.z goto disk2

echo Total Distribution Diskette Required = 1
goto start

:disk2
echo Total Distribution Diskette Required = 2
goto start

:disk3
echo Total Distribution Diskette Required = 3
goto start

:disk4
echo Total Distribution Diskette Required = 4
goto start

:disk5
echo Total Distribution Diskette Required = 5
goto start

:disk6
echo Total Distribution Diskette Required = 6

:disk7
echo Total Distribution Diskette Required = 7

:disk8
echo Total Distribution Diskette Required = 8
:start

echo .
echo insert disk#1
pause

deltree /Y a:\*.*
xcopy ..\disk\disk1\*.* a:

echo insert disk#2
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk2\*.* a:

if exist ..\disk\disk2\factdb.z goto makedisk_done

echo insert disk#3
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk3\*.* a:

if exist ..\disk\disk3\factdb.z goto makedisk_done

echo insert disk#4
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk4\*.* a:

if exist ..\disk\disk4\factdb.z goto makedisk_done

echo insert disk#5
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk5\*.* a:

if exist ..\disk\disk5\factdb.z goto makedisk_done

echo insert disk#6
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk6\*.* a:

if exist ..\disk\disk6\factdb.z goto makedisk_done

echo insert disk#7
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk7\*.* a:

if exist ..\disk\disk7\factdb.z goto makedisk_done

echo insert disk#8
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk8\*.* a:

if exist ..\disk\disk8\factdb.z goto makedisk_done

if not exist ..\disk\disk8\factdb.z goto need_more_disk

:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .

goto finished

:need_more_disk
echo .
echo One or More Distribution Disk needs to be copied.
echo Copy the disk mannually start with Disk#6.
echo .

:finished
