echo .
del .\clients\exectrak\zfiles\factdll.z

echo .
echo 2d -- Compress all files in Dll directory to factdll.z
icomp G:\program\release\exectrak\dll\*.* .\clients\exectrak\zfiles\factdll.z
