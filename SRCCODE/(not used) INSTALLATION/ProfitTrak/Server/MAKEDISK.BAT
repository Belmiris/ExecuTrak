@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk9\crystal.3 goto disk9
if exist ..\disk\disk8\crystal.3 goto disk8
if exist ..\disk\disk7\crystal.3 goto disk7
if exist ..\disk\disk6\crystal.3 goto disk6

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
goto start

:disk8
echo Total Distribution Diskette Required = 8
goto start

:disk9
echo Total Distribution Diskette Required = 9
goto start

:start
echo .
echo insert disk#1
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk1\*.* a:

echo insert disk#2
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk2\*.* a:

echo insert disk#3
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk3\*.* a:

if exist ..\disk\disk3\crystal.3 goto makedisk_done

echo insert disk#4
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk4\*.* a:

if exist ..\disk\disk4\crystal.3 goto makedisk_done

echo insert disk#5
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk5\*.* a:

if exist ..\disk\disk5\crystal.3 goto makedisk_done

echo insert disk#6
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk6\*.* a:

if exist ..\disk\disk6\crystal.3 goto makedisk_done

echo insert disk#7
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk7\*.* a:

if exist ..\disk\disk7\crystal.3 goto makedisk_done

echo insert disk#8
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk8\*.* a:

if exist ..\disk\disk8\crystal.3 goto makedisk_done

echo insert disk#9
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk9\*.* a:

if exist ..\disk\disk9\crystal.3 goto makedisk_done

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
