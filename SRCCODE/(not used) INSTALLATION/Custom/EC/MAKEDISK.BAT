@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk6\localdb.2 goto disk6
if exist ..\disk\disk5\localdb.2 goto disk5
if exist ..\disk\disk4\localdb.2 goto disk4
if exist ..\disk\disk3\localdb.2 goto disk3
if exist ..\disk\disk2\localdb.2 goto disk2
if exist ..\disk\disk4\localdb.z goto disk4
if exist ..\disk\disk3\localdb.z goto disk3
if exist ..\disk\disk2\localdb.z goto disk2

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

@echo off
echo .
echo insert disk1
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk1\*.* a:

if exist ..\disk\disk1\localdb.2 goto makedisk_done
if exist ..\disk\disk1\localdb.z goto makedisk_done

echo insert disk2
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk2\*.* a:

if exist ..\disk\disk2\localdb.2 goto makedisk_done
if exist ..\disk\disk2\localdb.z goto makedisk_done

echo insert disk3
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk3\*.* a:

if exist ..\disk\disk3\localdb.2 goto makedisk_done
if exist ..\disk\disk3\localdb.z goto makedisk_done

echo insert disk4
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk4\*.* a:

if exist ..\disk\disk4\localdb.2 goto makedisk_done
if exist ..\disk\disk4\localdb.z goto makedisk_done

echo insert disk5
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk5\*.* a:

if exist ..\disk\disk5\localdb.2 goto makedisk_done
if exist ..\disk\disk5\localdb.z goto makedisk_done

echo insert disk6
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk6\*.* a:

if not exist ..\disk\disk6\localdb.2 goto need_more_disk
if not exist ..\disk\disk6\localdb.z goto need_more_disk

:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .

goto finished

:need_more_disk
echo .
echo One or More Distribution Disk needs to be copied.
echo Copy the disk mannually start with Disk#7.
echo .

:finished


