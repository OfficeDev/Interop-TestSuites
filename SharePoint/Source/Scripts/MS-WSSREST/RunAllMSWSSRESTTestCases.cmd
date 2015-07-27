@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-WSSREST\TestSuite\bin\Debug\MS-WSSREST_TestSuite.dll /runconfig:..\..\MS-WSSREST\MS-WSSREST.testsettings 
pause