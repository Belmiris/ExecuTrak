echo .
del .\clients\exectrak\zfiles\crystal.z

echo .
echo 2b -- Compress all files in Crystal directory to crystal.z
icomp G:\program\release\exectrak\crystal\*.* .\clients\exectrak\zfiles\crystal.z
