echo .
del .\clients\exectrak\zfiles\localdb.z

echo .
echo 2e -- Compress Local Database
icomp G:\program\release\exectrak\local_db\*.* .\clients\exectrak\zfiles\localdb.z
