@echo off

rem if "%1" == "" goto parm_missing

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

