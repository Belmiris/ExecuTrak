@echo off
deltree /Y ..\disk\disk1
md ..\disk\disk1
copy disk.id ..\disk\disk1\disk1.id

deltree /Y ..\disk\disk2
md ..\disk\disk2
copy disk.id ..\disk\disk2\disk2.id

