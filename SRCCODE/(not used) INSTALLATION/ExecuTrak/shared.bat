echo .
del .\clients\exectrak\zfiles\shared.z

echo .
echo 2g -- Compress all files in Shared directory to shared.z
icomp G:\program\release\exectrak\shared\*.* .\clients\exectrak\zfiles\shared.z
