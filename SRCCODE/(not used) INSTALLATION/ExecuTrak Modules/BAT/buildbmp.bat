@echo off
rem BUILDBMP.BAT


if "%1" == "" goto module_missing
if "%2" == "" goto module_missing

choice /c:yn /t:y,7 Do you want to create the SETUP.BMP?

if errorlevel 2 goto skip_create


rem SETUP COMMAND LINE ARGUMENTS

rem DEFAULT COMMAND LINE ARGUMENTS

set path1=\project6\%1\%1server\setup
set mod_name=
set tm=TRUE
set path2=\project6\%1\%1client\setup

if "%1" == "ap" goto set_ap_mod
if "%1" == "cm" goto set_cm_mod
if "%1" == "ar" goto set_ar_mod
if "%1" == "ec" goto set_ec_mod
if "%1" == "et" goto set_et_mod
if "%1" == "fd" goto set_fd_mod
if "%1" == "fm" goto set_fm_mod
if "%1" == "fo" goto set_fo_mod
if "%1" == "gf" goto set_gf_mod
if "%1" == "gl" goto set_gl_mod
if "%1" == "oe" goto set_oe_mod
if "%1" == "ps" goto set_ps_mod
if "%1" == "pr" goto set_pr_mod
if "%1" == "rp" goto set_rp_mod
if "%1" == "rs" goto set_rs_mod
if "%1" == "sm" goto set_sm_mod
if "%1" == "tg" goto set_tg_mod
if "%1" == "tx" goto set_tx_mod
if "%1" == "ws" goto set_ws_mod
if "%1" == "sp" goto set_sp_mod

goto main


:set_ap_mod
set mod_name="Account Payable"
goto main

:set_cm_mod
set mod_name="Credit Management System"
goto main


:set_ar_mod
set mod_name="Account Receivable"
goto main


:set_fd_mod
set mod_name="Fuel Dispatch"
goto main


:set_fm_mod
set mod_name="Fuel Management"
goto main


:set_fo_mod
set mod_name="Heating System"
goto main


:set_gf_mod
set mod_name="Advanced Financials"
goto main


:set_gl_mod
set mod_name="General Ledger"
goto main


:set_oe_mod
set mod_name="Order Entry"
goto main


:set_pr_mod
set mod_name="Payroll"
goto main


:set_ps_mod
set mod_name="Petro South Custom"
set tm=FALSE
goto main


:set_rp_mod
set mod_name="Advanced C-Store"
goto main


:set_rs_mod
set mod_name="Retail Sales"
goto main


:set_sm_mod
set mod_name="Service & Maintenance"
goto main


:set_tg_mod
set mod_name="Remote G/L Journal Transaction"
goto main


:set_tx_mod
set mod_name="Tax Control"
goto main


:set_ws_mod
set mod_name="WholeSale"
goto main

:set_sp_mod
set mod_name="Factor Service Pack"
goto main

:main

if "%mod_name" == "" goto mod_name_missing

rem Create SETUP.BMP

echo .
echo BUILD SETUP.BMP
echo Please wait for program to create the SETUP.BMP
echo DO NOT PRESS ANY KEY BEFORE THE PROGRAM IS DONE!!!

i:\program\factmenu\bin\setupbmp.exe PATH=%path1% MODULE=%mod_name% VERSION=%2 TM=%tm% PATH2=%path2%
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
echo BUILDBMP module_name module_version
echo .


:done

set path1=
set mod_name=
set tm=
set path2=

