echo off

rem this batch file will be use when you want to copy the setup from local hard drive
rem to diskette.

rem parm1 = module name.

if "%1" == "" goto module_missing

echo off
cls

echo .
echo Copying %1 distribution disk.
echo .

if exist \project6\%1\setup\nul goto standalone

cd \project6\%1\%1server\setup
goto makedisk

:standalone
cd \project6\%1\setup

:makedisk
call makedisk.bat
echo .

echo done
goto done


:module_missing
echo Paramater is missing
echo .
echo MAKEDISK module_name
echo .


:done

cd \project6
