@echo off

if "%1"=="" goto parm1_is_missing
if "%2"=="" goto parm2_is_missing

icomp %1 %2
echo .
echo Update %1 to %2 has completed.
echo .
goto finished

:parm1_is_missing
echo .
echo [input path\]filename not specified, update z file failed.
echo .
goto finished

:parm2_is_missing
echo .
echo [output path\]zfilename not specified, update z file failed.
echo .

:finished

