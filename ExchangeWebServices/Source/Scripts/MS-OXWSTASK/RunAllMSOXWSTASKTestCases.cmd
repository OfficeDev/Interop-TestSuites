@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSTASK\TestSuite\bin\Debug\MS-OXWSTASK_TestSuite.dll /runconfig:..\..\MS-OXWSTASK\MS-OXWSTASK.testsettings 
pause