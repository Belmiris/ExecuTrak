echo off
rem This batch file will be used to update the disk build log database
rem EXBUILD module_version [module_name] 
rem if module_name is not supplied then compile all module

cls

if "%1" == "" goto parm_missing
if NOT "%2" == "" goto update_one


:all_module
echo .
choice /c:yn /t:y,7 Are you sure you want to update the disk build log for all modules

if errorlevel 1 goto update_all

echo cancel updating disk build log
echo .
goto done


:update_all
for %%i in (ap ar cm fd fm fo gf gl oe pr rp rs sm tg tx ws ec et cstormnt execvisn rptvalue) do call bild_one %%i %1
cd \projecct6
echo .
echo done
goto done


:update_one
call bild_one %2 %1
cd \project6
echo .
echo done
goto done

:parm_missing
echo .
echo Parm is missing or is not valid.
echo Usage: EXBUILD module_version [module_name]
echo .


:done
