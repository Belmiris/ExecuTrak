rem this batch file is called by EXBUILD.BAT
rem
echo build disk log for %1 version %2
echo .
if exist \project6\%1\setup\nul goto standalone
cd \project6\%1\%1client\setup
goto build

:standalone
cd \project6\%1\setup
cd
call buildlog.bat %1 %2
echo .
cd\project6
