echo off
rem This batch file will be used to re-make distribution disk.
rem XXMAKE [module_name]
rem if module_name is not supplied then re-make all module distribution disk.

rem parm1 = module name or 'SKIP' for all module - skip copying EXE files.
rem parm2 = (optional) string 'SKIP' to skip copying files from G: to local

cls

if NOT "%1" == "" goto one_module


:all_module
echo .
choice /c:yn /t:y,7 Are you sure you want to re-make all modules setup.

if errorlevel 1 goto remake_all

echo cancel re-make setup
echo .
goto done


:remake_all

if "%1" == "" goto copy_exe

for %%i in (ap ar fd fm fo gf gl oe pr rp rs sm tg tx ws) do call exmake %%i SKIP
goto done


:copy_exe

for %%i in (ap ar fd fm fo gf gl oe pr rp rs sm tg tx ws) do call exmake %%i
goto done


:one_module
if "%2" == "" goto one_copy_exe

call exmake %1 SKIP
goto done


:one_copy_exe

call exmake %1
goto done


:done
cd \project6
echo .
echo done
