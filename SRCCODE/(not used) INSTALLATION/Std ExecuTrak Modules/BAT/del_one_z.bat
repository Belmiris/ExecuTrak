rem this batch file is called by DELEZ.BAT
rem
echo delete %1 server Z Files
echo .
cd \project6\%1\%1server\disk
cd
del *.z
echo .
echo delete %1 client Z Files
echo .
cd \project6\%1\%1client\disk\disk1
cd
del *.z
echo .
cd\project6
