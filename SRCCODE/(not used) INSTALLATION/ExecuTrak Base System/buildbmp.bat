@echo off
rem BUILDBMP.BAT


if "%1" == "" goto module_missing

choice /c:yn /t:y,7 Do you want to create the SETUP.BMP?

if errorlevel 2 goto skip_create


rem SETUP COMMAND LINE ARGUMENTS

set path1=\release6\exectrak\setup\clients\exectrak
set mod_name="Main Module"
set path2=\release6\exectrak\setup\servers\exectrak
set tm=TRUE

rem Create SETUP.BMP

echo .
echo BUILD SETUP.BMP
echo Please wait for program to create the SETUP.BMP
echo DO NOT PRESS ANY KEY BEFORE THE PROGRAM IS DONE!!!

i:\program\factmenu\bin\setupbmp.exe PATH=%path1% MODULE=%mod_name% VERSION=%1 TM=%tm% PATH2=%path2%
pause

goto done


:skip_create

echo .
echo Create SETUP.BMP has been skippped!
echo .
goto done


:module_missing
echo .
echo Paramater is missing
echo .
echo BUILDBMP module_version
echo e.g. BUILDBMP 8.06
echo .


:done

set path1=
set mod_name=
set co_name=
set tm=

