@echo off
echo .
echo Making Distribution Disk

if exist ..\disk\disk20\??client.20 goto disk20
if exist ..\disk\disk19\??client.19 goto disk19
if exist ..\disk\disk18\??client.18 goto disk18
if exist ..\disk\disk17\??client.17 goto disk17

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

:disk15
echo Total Distribution Diskette Required = 15
goto start

:disk16
echo Total Distribution Diskette Required = 16
goto start

:disk17
echo Total Distribution Diskette Required = 17
goto start

:disk18
echo Total Distribution Diskette Required = 18
goto start

:disk19
echo Total Distribution Diskette Required = 19
goto start

:disk20
echo Total Distribution Diskette Required = 20
goto start

:disk21
echo Total Distribution Diskette Required = 21
goto start

:disk22
echo Total Distribution Diskette Required = 22
goto start

:disk23
echo Total Distribution Diskette Required = 23
goto start

:disk24
echo Total Distribution Diskette Required = 24
goto start

:disk25
echo Total Distribution Diskette Required = 25
goto start

:disk26
echo Total Distribution Diskette Required = 26
goto start

:disk27
echo Total Distribution Diskette Required = 27
goto start

:disk28
echo Total Distribution Diskette Required = 28
goto start

:disk29
echo Total Distribution Diskette Required = 29
goto start

:disk30
echo Total Distribution Diskette Required = 30
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
if not exist ..\disk\disk15\??client.15 goto makedisk_done

echo insert disk#15
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk15\*.* a:
if not exist ..\disk\disk16\??client.16 goto makedisk_done

echo insert disk#16
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk16\*.* a:
if not exist ..\disk\disk17\??client.17 goto makedisk_done

echo insert disk#17
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk17\*.* a:
if not exist ..\disk\disk18\??client.18 goto makedisk_done

echo insert disk#18
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk18\*.* a:
if not exist ..\disk\disk19\??client.19 goto makedisk_done

echo insert disk#19
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk19\*.* a:
if not exist ..\disk\disk20\??client.20 goto makedisk_done

echo insert disk#20
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk20\*.* a:

if not exist ..\disk\disk21\??client.21 goto makedisk_done

echo insert disk#21
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk21\*.* a:

if not exist ..\disk\disk22\??client.22 goto makedisk_done

echo insert disk#22
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk22\*.* a:

if not exist ..\disk\disk23\??client.23 goto makedisk_done

echo insert disk#23
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk23\*.* a:

if not exist ..\disk\disk24\??client.24 goto makedisk_done

echo insert disk#24
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk24\*.* a:

if not exist ..\disk\disk25\??client.25 goto makedisk_done

echo insert disk#25
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk25\*.* a:

if not exist ..\disk\disk26\??client.26 goto makedisk_done

echo insert disk#26
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk26\*.* a:

if not exist ..\disk\disk27\??client.27 goto makedisk_done

echo insert disk#27
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk27\*.* a:

if not exist ..\disk\disk28\??client.28 goto makedisk_done

echo insert disk#28
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk28\*.* a:

if not exist ..\disk\disk29\??client.29 goto makedisk_done

echo insert disk#29
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk29\*.* a:

if not exist ..\disk\disk30\??client.30 goto makedisk_done

echo insert disk#30
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk30\*.* a:

if not exist ..\disk\disk31\??client.31 goto makedisk_done

echo insert disk#31
pause

echo Delete all Files in A drive
del a:\*.* 
xcopy ..\disk\disk31\*.* a:

if exist ..\disk\disk32\??client.32 goto need_more_disk
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
