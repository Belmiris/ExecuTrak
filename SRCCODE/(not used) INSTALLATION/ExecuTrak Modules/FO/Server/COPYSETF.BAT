@echo off

if "%1" == "" goto parm_missing

echo .
choice /c:yn /t:y,7 Do you want to Compile the Server Setup Script
if errorlevel 2 goto skipcompile

compile setup
echo .
echo .
echo Create a new packlist -- setup.lst
echo 1; > setup.lst
echo ..\disk\%1client.z >> setup.lst

packlist setup.lst
echo .


:skipcompile
copy \template\setup\_inst32i.ex_ ..\disk\disk1
copy \template\setup\_isdel.exe ..\disk\disk1
copy \template\setup\_setup.dll ..\disk\disk1
copy \template\setup\_setup.lib ..\disk\disk1
copy \template\setup\setup.exe ..\disk\disk1
copy setup.ins ..\disk\disk1
copy setup.pkg ..\disk\disk1
copy setup.dbg ..\disk\disk1
copy setup.bmp ..\disk\disk1
echo .
echo Finished copying all setup files to DISK#1

goto finished

:parm_missing
echo .
echo missing parameter.  format: copysetf [AppName].
echo AppName - AP, AR, GL, and etc.
echo .

:finished

