@echo off
call updatez .\clients\exectrak\setup.ins .\servers\exectrak\zfiles\setup.z
call updatez .\clients\exectrak\setup.pkg .\servers\exectrak\zfiles\setup.z
echo y | copy .\servers\exectrak\zfiles\setup.z .\disk\disk1
