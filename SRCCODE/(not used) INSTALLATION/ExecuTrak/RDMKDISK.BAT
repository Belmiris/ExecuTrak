@echo off
echo insert disk1
pause

rem  copy \exectrak\setup\*.* a:
copy G:\home\group\program\release\exectrak\setup\*.* a:
copy zfiles\factole.z a:
copy zfiles\factmenu.z a:
copy zfiles\setup.z a:
copy setup.ins a:
copy setup.pkg a:
call makeid.bat 1

echo insert disk2
pause

copy zfiles\sybin.z a:
call makeid.bat 2

echo insert disk3
pause

copy zfiles\client.1 a:
call makeid.bat 3

echo insert disk4
pause

copy zfiles\client.2 a:
call makeid.bat 4

echo insert disk5
pause

copy zfiles\client.3 a:
call makeid.bat 5

echo insert disk6
pause

copy zfiles\client.4 a:
call makeid.bat 6

