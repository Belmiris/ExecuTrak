echo .
del .\clients\exectrak\zfiles\custctl.z

echo .
echo 2c -- Compress all files in Custctl directory to custctl.z
icomp G:\program\release\exectrak\custctl\*.* .\clients\exectrak\zfiles\custctl.z
