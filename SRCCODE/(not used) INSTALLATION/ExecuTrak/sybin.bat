echo .
del .\clients\exectrak\zfiles\sybin.z

echo 2a -- Compress all files in bin directory to sybin.z
icomp G:\program\release\exectrak\sy\bin\*.* .\clients\exectrak\zfiles\sybin.z
icomp G:\program\release\exectrak\common\*.exe .\clients\exectrak\zfiles\sybin.z
