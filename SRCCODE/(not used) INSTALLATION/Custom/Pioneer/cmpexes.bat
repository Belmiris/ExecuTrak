rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z

echo .
echo Compress Application files

echo .
echo call DISKBILD.BAT
call buildlog.bat PIONEER %1

echo .
echo Compress all files to factbin.z
icomp g:\program\release\custom\pioneer\*.* ..\disk\factbin.z -i
icomp g:\program\release\exectrak\common\cpylocal.exe ..\disk\factbin.z BIN

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo Z File sequence (FACTBIN.Z, SHARED.Z, FACTDLL.Z, FACTOLE.Z, CUSTCTL.Z
echo                   , FACTDB.Z, CRYSTAL.Z, RTM.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
