echo off
rem This batch file will be used to re-compile the setup.rul file
rem in both server and client setup directory.
rem EXCOMP [module_name]
rem if module_name is not supplied then compile all module

cls

if NOT "%1" == "" goto one_module


:all_module
echo .
choice /c:yn /t:y,7 Are you sure you want to compile all modules' setup.rul.

if errorlevel 1 goto compile_all

echo cancel compile setup.rul
echo .
goto done


:compile_all
for %%i in (ap ar fd fm fo gf gl oe pr rp rs sm tg tx ws) do call comp_one %%i
cd \projec~1
echo .
echo done
goto done


:one_module
call comp_one %1
cd \projec~1
echo .
echo done
goto done


:done
