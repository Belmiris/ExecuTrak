rem Assume: current directory is \release\exectrak\setup

@echo off
cls

call makedir

echo .
echo ******1 copy server setup files
call copysetf
echo .
echo ******2 copy setup.z
xcopy .\servers\exectrak\zfiles\setup.z .\disk\disk1

echo .
echo ******3 split apclient.z
call fsplit .\servers\exectrak\zfiles\client.z .\servers\exectrak\zfiles 60

move .\servers\exectrak\zfiles\client.1 .\disk\disk1
move .\servers\exectrak\zfiles\client.2 .\disk\disk2
move .\servers\exectrak\zfiles\client.3 .\disk\disk3
move .\servers\exectrak\zfiles\client.4 .\disk\disk4
move .\servers\exectrak\zfiles\client.5 .\disk\disk5
move .\servers\exectrak\zfiles\client.6 .\disk\disk6
move .\servers\exectrak\zfiles\client.7 .\disk\disk7
if not exist .\servers\exectrak\zfiles\client.8 goto finished
move .\servers\exectrak\zfiles\client.8 .\disk\disk8
if not exist .\servers\exectrak\zfiles\client.9 goto finished
move .\servers\exectrak\zfiles\client.9 .\disk\disk9
if not exist .\servers\exectrak\zfiles\client.10 goto finished
move .\servers\exectrak\zfiles\client.10 .\disk\disk10
if not exist .\servers\exectrak\zfiles\client.11 goto finished
move .\servers\exectrak\zfiles\client.11 .\disk\disk11
if not exist .\servers\exectrak\zfiles\client.12 goto finished
move .\servers\exectrak\zfiles\client.12 .\disk\disk12
if not exist .\servers\exectrak\zfiles\client.13 goto finished
move .\servers\exectrak\zfiles\client.13 .\disk\disk13
if not exist .\servers\exectrak\zfiles\client.14 goto finished
move .\servers\exectrak\zfiles\client.14 .\disk\disk14
if not exist .\servers\exectrak\zfiles\client.15 goto finished
move .\servers\exectrak\zfiles\client.15 .\disk\disk15
if not exist .\servers\exectrak\zfiles\client.10 goto finished
move .\servers\exectrak\zfiles\client.16 .\disk\disk16
if not exist .\servers\exectrak\zfiles\client.17 goto finished
move .\servers\exectrak\zfiles\client.17 .\disk\disk17
if not exist .\servers\exectrak\zfiles\client.18 goto finished
move .\servers\exectrak\zfiles\client.18 .\disk\disk18
if not exist .\servers\exectrak\zfiles\client.19 goto finished
move .\servers\exectrak\zfiles\client.19 .\disk\disk19
if not exist .\servers\exectrak\zfiles\client.20 goto finished
move .\servers\exectrak\zfiles\client.20 .\disk\disk20
if exist .\servers\exectrak\zfiles\client.21 goto moredisk


:finished
echo .
echo split .z files is finished.  Run makedisk.bat to copy distribution disk.
goto quit


:moredisk
echo .
echo one or more zfiles have not been copy to .\disk\disk?? directory.
echo .
echo press CTRL-C to break the process.
echo .
echo TO DO:
echo        0. check to see how many zfiles (with numbered extension)
echo           left in current directory.
echo        1. make a new directory called disk?? under .\disk directory.
echo        2. copy disk id and client.?? to .\disk\disk?? directory.
echo        3. change directory to \release\exectrak\setup
echo        4. run makedisk.bat to copy distribution disk


:quit

