@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk14\??client.14 goto disk14
if exist ..\disk\disk13\??client.13 goto disk13
if exist ..\disk\disk12\??client.12 goto disk12
if exist ..\disk\disk11\??client.11 goto disk11
if exist ..\disk\disk10\??client.10 goto disk10
if exist ..\disk\disk9\??client.9 goto disk9
if exist ..\disk\disk8\??client.8 goto disk8
if exist ..\disk\disk7\??client.7 goto disk7
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

:disk7
echo Total Distribution Diskette Required = 7
goto start

:disk8
echo Total Distribution Diskette Required = 8
goto start

:disk9
echo Total Distribution Diskette Required = 9
goto start

:disk10
echo Total Distribution Diskette Required = 10
goto start

:disk11
echo Total Distribution Diskette Required = 11
goto start

:disk12
echo Total Distribution Diskette Required = 12
goto start

:disk13
echo Total Distribution Diskette Required = 13
goto start

:disk14
echo Total Distribution Diskette Required = 14
goto start

:start
echo .
echo insert disk#1
pause

echo Delete all Files in A drive
del a:\*.*
xcopy ..\disk\disk1\*.* a:

if exist ..\disk\disk1\*.z goto makedisk_done

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

if not exist ..\disk\disk7\??client.7 goto makedisk_done

echo insert disk#7
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk7\*.* a:
if not exist ..\disk\disk8\??client.8 goto makedisk_done

echo insert disk#8
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk8\*.* a:
if not exist ..\disk\disk9\??client.9 goto makedisk_done

echo insert disk#9
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk9\*.* a:
if not exist ..\disk\disk10\??client.10 goto makedisk_done

echo insert disk#10
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk10\*.* a:
if not exist ..\disk\disk11\??client.11 goto makedisk_done

echo insert disk#11
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk11\*.* a:
if not exist ..\disk\disk12\??client.12 goto makedisk_done

echo insert disk#12
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk12\*.* a:
if not exist ..\disk\disk13\??client.13 goto makedisk_done

echo insert disk#13
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk13\*.* a:
if not exist ..\disk\disk14\??client.14 goto makedisk_done

echo insert disk#14
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk14\*.* a:

if exist ..\disk\disk15\??client.15 goto need_more_disk

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
