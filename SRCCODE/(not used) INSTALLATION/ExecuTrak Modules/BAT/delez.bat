echo off
rem This batch file will be used to delete all Z FILES in DISK directory.
rem DELEZ [module_name]
rem if module_name is not supplied then delete all module Z FILES.

cls

if NOT "%1" == "" goto one_module


:all_module
echo .
choice /c:yn /t:y,7 Are you sure you want to delete all modules Z FILES.

if errorlevel 1 goto delete_all

echo cancel delete Z FILES
echo .
goto done


:delete_all

for %%i in (ap ar fd fm fo gf gl oe pr rp rs sm tg tx ws) do call del_one_z %%i
goto done


:one_module
call del_one_z %1
goto done


:done
cd \project6
echo .
echo done
