echo .
del .\clients\exectrak\zfiles\rtm.z

echo .
echo 2g -- Compress all files in Rtm directory to rtm.z
icomp G:\program\release\exectrak\rtm\*.* .\clients\exectrak\zfiles\rtm.z
