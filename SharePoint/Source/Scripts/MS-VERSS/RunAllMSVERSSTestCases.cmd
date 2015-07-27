@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-VERSS\TestSuite\bin\Debug\MS-VERSS_TestSuite.dll /runconfig:..\..\MS-VERSS\MS-VERSS.testsettings 
pause