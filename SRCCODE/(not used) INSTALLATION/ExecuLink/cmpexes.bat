rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat RPPCA %1

echo .
echo Compress all files in g:\program\release\execlink
echo and cpylocal.exe to factbin.z
icomp g:\program\release\execlink\*.* ..\disk\factbin.z -i
icomp g:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo Z File sequence (FACTBIN.Z, SHARED.Z, FACTDLL.Z, FACTOLE.Z, CUSTCTL.Z
echo                   , RTM.Z, FACTDB.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.

echo .
choice /c:yn /t:y,7 Do you want to run the splitz.bat.

if errorlevel 2 goto finished

echo .
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
