rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat SIMONS %1

echo .
echo Compress all files in appfiles directory to factbin.z
icomp G:\PROGRAM\RELEASE\Custom\Simons\*.* ..\disk\factbin.z
icomp g:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN
icomp G:\PROGRAM\RELEASE\Custom\Simons\BIN\*.* ..\disk\factbin.z BIN

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
choice /c:yn /t:y,7 Do you want to Copy Files to Disk#1
if errorlevel 2 goto finished

echo .
fsecho insert disk#1
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


:finished
