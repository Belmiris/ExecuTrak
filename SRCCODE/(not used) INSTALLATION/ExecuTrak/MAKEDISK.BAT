@echo off
echo .
echo Making Distribution Disk

if exist disk\disk20\client.20 goto disk20
if exist disk\disk19\client.19 goto disk19
if exist disk\disk18\client.18 goto disk18
if exist disk\disk17\client.17 goto disk17
if exist disk\disk16\client.16 goto disk16
if exist disk\disk15\client.15 goto disk15
if exist disk\disk14\client.14 goto disk14
if exist disk\disk13\client.13 goto disk13
if exist disk\disk12\client.12 goto disk12
if exist disk\disk11\client.11 goto disk11
if exist disk\disk10\client.10 goto disk10
if exist disk\disk9\client.9 goto disk9
if exist disk\disk8\client.8 goto disk8
if exist disk\disk7\client.7 goto disk7
if exist disk\disk6\client.6 goto disk6
if exist disk\disk5\client.5 goto disk5
if exist disk\disk4\client.4 goto disk4
if exist disk\disk3\client.3 goto disk3
if exist disk\disk2\client.2 goto disk2

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


:start

echo .
echo insert disk#1
pause

deltree /Y a:\*.*
xcopy disk\disk1\*.* a:

if not exist disk\disk2\client.2 goto makedisk_done

echo insert disk#2
pause

deltree /Y a:\*.* 
xcopy disk\disk2\*.* a:

if not exist disk\disk3\client.3 goto makedisk_done

echo insert disk#3
pause

deltree /Y a:\*.* 
xcopy disk\disk3\*.* a:

if not exist disk\disk4\client.4 goto makedisk_done

echo insert disk#4
pause

deltree /Y a:\*.* 
xcopy disk\disk4\*.* a:

if not exist disk\disk5\client.5 goto makedisk_done

echo insert disk#5
pause

deltree /Y a:\*.* 
xcopy disk\disk5\*.* a:

if not exist disk\disk6\client.6 goto makedisk_done

echo insert disk#6
pause

deltree /Y a:\*.* 
xcopy disk\disk6\*.* a:

if not exist disk\disk7\client.7 goto makedisk_done

echo insert disk#7
pause

deltree /Y a:\*.* 
xcopy disk\disk7\*.* a:

if not exist disk\disk8\client.8 goto makedisk_done

echo insert disk#8
pause

deltree /Y a:\*.* 
xcopy disk\disk8\*.* a:

if not exist disk\disk9\client.9 goto makedisk_done

echo insert disk#9
pause

deltree /Y a:\*.* 
xcopy disk\disk9\*.* a:

if not exist disk\disk10\client.10 goto makedisk_done

echo insert disk#10
pause

deltree /Y a:\*.* 
xcopy disk\disk10\*.* a:

if not exist disk\disk11\client.11 goto makedisk_done

echo insert disk#11
pause

deltree /Y a:\*.* 
xcopy disk\disk11\*.* a:

if not exist disk\disk12\client.12 goto makedisk_done

echo insert disk#12
pause

deltree /Y a:\*.* 
xcopy disk\disk12\*.* a:

if not exist disk\disk13\client.13 goto makedisk_done

echo insert disk#13
pause

deltree /Y a:\*.* 
xcopy disk\disk13\*.* a:

if not exist disk\disk14\client.14 goto makedisk_done

echo insert disk#14
pause

deltree /Y a:\*.* 
xcopy disk\disk14\*.* a:

if not exist disk\disk15\client.15 goto makedisk_done

echo insert disk#15
pause

deltree /Y a:\*.* 
xcopy disk\disk15\*.* a:

if not exist disk\disk16\client.16 goto makedisk_done

echo insert disk#16
pause

deltree /Y a:\*.* 
xcopy disk\disk16\*.* a:

if not exist disk\disk17\client.17 goto makedisk_done

echo insert disk#17
pause

deltree /Y a:\*.* 
xcopy disk\disk17\*.* a:

if not exist disk\disk18\client.18 goto makedisk_done

echo insert disk#18
pause

deltree /Y a:\*.* 
xcopy disk\disk18\*.* a:

if not exist disk\disk19\client.19 goto makedisk_done

echo insert disk#19
pause

deltree /Y a:\*.* 
xcopy disk\disk19\*.* a:

if not exist disk\disk20\client.20 goto makedisk_done

echo insert disk#20
pause

deltree /Y a:\*.* 
xcopy disk\disk20\*.* a:



:makedisk_done
echo .
echo Distribution Disk making has completed.
echo .
