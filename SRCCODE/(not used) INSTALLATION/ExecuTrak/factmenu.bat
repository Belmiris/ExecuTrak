echo .
del .\clients\exectrak\zfiles\factmenu.z

echo .
echo 2h -- Compress all files in parent directory to factmenu.z
icomp G:\program\release\exectrak\sy\*.* .\clients\exectrak\zfiles\factmenu.z
icomp G:\program\release\exectrak\common\*.hlp .\clients\exectrak\zfiles\factmenu.z
