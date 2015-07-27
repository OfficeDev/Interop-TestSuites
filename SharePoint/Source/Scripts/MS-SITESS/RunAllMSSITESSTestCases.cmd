@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-SITESS\TestSuite\bin\Debug\MS-SITESS_TestSuite.dll /runconfig:..\..\MS-SITESS\MS-SITESS.testsettings 
pause