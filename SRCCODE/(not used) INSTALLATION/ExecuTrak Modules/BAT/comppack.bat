echo off
cls
echo %1-SERVER COMPILE AND PACKLIST 
cd \Project6\%1\%1server\setup
compile setup.rul
echo .
packlist setup.lst
echo .
echo %1-CLIENT COMPILE AND PACKLIST 
cd \Project6\%1\%1client\setup
compile setup.rul
echo .
packlist setup.lst
echo .
cd \Project6
echo %1 - FINISHED
