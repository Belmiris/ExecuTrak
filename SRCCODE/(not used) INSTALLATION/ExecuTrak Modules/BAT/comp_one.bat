rem this batch file is called by EXCOMP.BAT
rem
echo compile %1 server setup.rul
echo .
cd \projec~1\%1\%1server\setup
cd
compile setup.rul
echo .
echo compile %1 client setup.rul
echo .
cd \projec~1\%1\%1client\setup
cd
compile setup.rul
echo .
cd\projec~1
