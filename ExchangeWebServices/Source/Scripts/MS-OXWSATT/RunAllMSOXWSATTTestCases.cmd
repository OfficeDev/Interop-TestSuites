@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSATT\TestSuite\bin\Debug\MS-OXWSATT_TestSuite.dll /runconfig:..\..\MS-OXWSATT\MS-OXWSATT.testsettings 
pause