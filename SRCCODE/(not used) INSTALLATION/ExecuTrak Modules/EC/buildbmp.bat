@echo off
rem BUILDBMP.BAT


if "%1" == "" goto module_missing

choice /c:yn /t:y,7 Do you want to create the SETUP.BMP?

if errorlevel 2 goto skip_create

rem Create SETUP.BMP

echo .
echo BUILD SETUP.BMP
echo Please wait for program to create the SETUP.BMP
echo DO NOT PRESS ANY KEY BEFORE THE PROGRAM IS DONE!!!

g:\program\factmenu\bin\setupbmp.exe PATH="\PROJECT6\EC\SETUP" MODULE="Electronic Commerce" VERSION=%1 TM=FALSE
pause

goto done


:skip_create

echo .
echo Create SETUP.BMP has been skippped!
echo .
goto done


:mod_name_missing

echo .
echo Batch File Error.  Module Name is missing!
echo .
goto done


:module_missing
echo .
echo Paramater is missing
echo .
echo BUILDBMP module_version
echo .


:done

