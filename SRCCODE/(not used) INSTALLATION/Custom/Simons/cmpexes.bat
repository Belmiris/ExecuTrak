rem Compile ONLY the EXE files
@echo off

if "%1" == "" goto parm_missing

echo Delete z files
del ..\disk\factbin.z
rem del ..\disk\factdb.z

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

echo .
echo Compress Local Database
icomp g:\program\release\exectrak\local_db\*.* ..\disk\factdb.z

echo .
echo Run the PACKCALC and edit the SPLITZ.BAT/SETUP.LST
echo .
echo Z File sequence (FACTBIN.Z, TLC.Z ,FACTOLE.Z, CUSTCTL.Z,
echo                  FACTDLL.Z, RTM.Z, SHARED.Z, FACTDB.Z)

echo .
echo Then run the SPLITZ.BAT to split the Z file.

echo .
echo RUN the SPLITZ.BAT
call splitz

goto finished

:parm_missing
echo .
echo missing or invalid parameter.  format: cmpfiles [Module Version]
echo .

:finished
