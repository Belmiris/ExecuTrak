@echo off

echo .

echo Copy FACTOR.IQK from g:\program\release\objects\iqk to ..\disk

Copy g:\program\release\objects\iqk\FACTOR.IQK ..\disk

compile setup.rul

echo y | copy setup.ins ..\disk

echo .
echo Setup Rule Compilation Done!
