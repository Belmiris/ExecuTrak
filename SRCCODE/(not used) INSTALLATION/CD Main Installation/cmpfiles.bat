@echo off

echo .

compile setup.rul

echo y | copy setup.ins ..\disk

echo .
echo Setup Rule Compilation Done!
