echo off

rem this batch file will be use when you want to make a distribution disk.
rem 
rem parm1 = module name.
rem parm2 = module version.
rem parm3 = Beta version.

:begin
if "%1" == "" goto module_missing
if "%2" == "" goto module_missing

if "%2" == "b" goto module_missing
if "%2" == "B" goto module_missing

echo off
cls

if exist \project6\%1\setup\nul goto standalone

call buildbmp.bat %1 %2

:compile_file

echo client + server
echo compile files
echo .
cd \project6\%1\%1server\setup
call cmpfiles.bat b %1 %2 %3

echo .
echo done
goto done

:standalone

echo standalone
echo compile files
echo .
cd \project6\%1\setup
call cmpfiles.bat %2

echo .
echo done
goto done

:module_missing
echo Paramater is missing
echo .
echo EXMAKE module_name module_version [Beta]
echo Beta keyword is optional that will bypass the copying files
echo .


:done

cd \project6
