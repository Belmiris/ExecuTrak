@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk7\rtm.4 goto disk7
if exist ..\disk\disk6\rtm.4 goto disk6
if exist ..\disk\disk5\rtm.4 goto disk5
if exist ..\disk\disk4\rtm.4 goto disk4
if exist ..\disk\disk3\rtm.4 goto disk3
if exist ..\disk\disk2\rtm.4 goto disk2

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
goto start

:disk7
echo Total Distribution Diskette Required = 7

:start
echo .
echo insert disk#1
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk1\*.* a:

if exist ..\disk\disk1\factdb.z goto makedisk_done

echo insert disk#2
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk2\*.* a:

if exist ..\disk\disk2\factdb.z goto makedisk_done

echo insert disk#3
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk3\*.* a:

rem if not exist ..\disk\disk4\??client.4 goto makedisk_done
if exist ..\disk\disk3\factdb.z goto makedisk_done

echo insert disk#4
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk4\*.* a:

if exist ..\disk\disk4\factdb.z goto makedisk_done

echo insert disk#5
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk5\*.* a:

if exist ..\disk\disk5\factdb.z goto makedisk_done

echo insert disk#6
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk6\*.* a:

if exist ..\disk\disk6\factdb.z goto makedisk_done

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
