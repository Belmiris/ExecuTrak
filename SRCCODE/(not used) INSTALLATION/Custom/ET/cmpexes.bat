rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo .
choice /c:yn /t:y,7 Do you want to COPY FILES FROM RELEASE directory
if errorlevel 2 goto skipcopyfiles

echo .
echo Copy files from release directory
call copyfile


:skipcopyfiles

echo .
choice /c:yn /t:y,7 Do you want to CONTINUE the DISK BUILD PROCESS
if errorlevel 2 goto quitbuild


echo Delete z files
del ..\disk\factbin.z
rem del ..\disk\factdb.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat %1

echo .
echo Compress all files in appfiles directory to factbin.z
icomp ..\appfiles\*.* ..\disk\factbin.z -i

echo .
echo Compress Local Database
icomp I:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

rem echo .
rem echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
rem echo and then RUN the SPLITZ.BAT

echo .
echo RUN the SPLITZ.BAT
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .
goto finished


:quitbuild
echo .
echo Disk Build Process has been cancelled.
echo .


:finished
