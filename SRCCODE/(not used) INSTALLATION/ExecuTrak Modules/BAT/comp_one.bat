rem this batch file is called by EXCOMP.BAT
rem
echo compile %1 server setup.rul
echo .
cd \Project6\%1\%1server\setup
cd
compile setup.rul
echo .
cd\Project6
