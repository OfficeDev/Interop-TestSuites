@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ASTASK\TestSuite\bin\Debug\MS-ASTASK_TestSuite.dll /runconfig:..\..\MS-ASTASK\MS-ASTASK.testsettings 
pause