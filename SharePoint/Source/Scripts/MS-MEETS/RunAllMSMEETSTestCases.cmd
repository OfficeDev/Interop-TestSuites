@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-MEETS\TestSuite\bin\Debug\MS-MEETS_TestSuite.dll /runconfig:..\..\MS-MEETS\MS-MEETS.testsettings 
pause