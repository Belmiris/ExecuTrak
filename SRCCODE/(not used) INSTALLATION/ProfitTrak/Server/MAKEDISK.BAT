@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk6\??client.6 goto disk6
if exist ..\disk\disk5\??client.5 goto disk5
if exist ..\disk\disk4\??client.4 goto disk4
if exist ..\disk\disk3\??client.3 goto disk3
if exist ..\disk\disk2\??client.2 goto disk2

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


:start
echo .
echo insert disk#1
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk1\*.* a:
if not exist ..\disk\disk2\??client.2 goto makedisk_done

echo insert disk#2
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk2\*.* a:
if not exist ..\disk\disk3\??client.3 goto makedisk_done

echo insert disk#3
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk3\*.* a:

if not exist ..\disk\disk4\??client.4 goto makedisk_done

echo insert disk#4
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk4\*.* a:

if not exist ..\disk\disk5\??client.5 goto makedisk_done

echo insert disk#5
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk5\*.* a:

if not exist ..\disk\disk6\??client.6 goto makedisk_done

echo insert disk#6
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk6\*.* a:

if exist ..\disk\disk7\??client.7 goto need_more_disk

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
