@echo off
echo .
echo Making Distribution Disk
echo .

if exist ..\disk\disk8\factdb.* goto disk6
if exist ..\disk\disk7\factdb.* goto disk6
if exist ..\disk\disk6\factdb.* goto disk6
if exist ..\disk\disk5\factdb.* goto disk5
if exist ..\disk\disk4\factdb.* goto disk4
if exist ..\disk\disk3\factdb.* goto disk3
if exist ..\disk\disk2\factdb.* goto disk2

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

if exist ..\disk\disk2\factdb.* goto makedisk_done

echo insert disk#3
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk3\*.* a:

if exist ..\disk\disk3\factdb.* goto makedisk_done

echo insert disk#4
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk4\*.* a:

if exist ..\disk\disk4\factdb.* goto makedisk_done

echo insert disk#5
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk5\*.* a:

if exist ..\disk\disk5\factdb.* goto makedisk_done

echo insert disk#6
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk6\*.* a:

if exist ..\disk\disk6\factdb.* goto makedisk_done

echo insert disk#7
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk7\*.* a:

if exist ..\disk\disk7\factdb.* goto makedisk_done

echo insert disk#8
pause

deltree /Y a:\*.* 
xcopy ..\disk\disk8\*.* a:

if not exist ..\disk\disk8\factdb.* goto need_more_disk

:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .

goto finished

:need_more_disk
echo .
echo One or More Distribution Disk needs to be copied.
echo Copy the disk mannually start with Disk#9.
echo .

:finished
