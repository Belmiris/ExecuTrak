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

echo NO split factbin.z
copy ..\disk\factbin.z ..\disk\disk1

echo .
choice /c:yn /t:y,7 Do you want to Compile the Setup Script
if errorlevel 2 goto skipcompile
compile setup
echo .
packlist setup.lst
echo .
call copysetf.bat
echo .

:skipcompile

echo .
echo insert disk#1
pause
copy ..\disk\disk1\factbin.z a:
copy ..\disk\disk1\setup.ins a:
copy ..\disk\disk1\setup.pkg a:

echo .
echo Distribution Disk making has completed.
echo .

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
