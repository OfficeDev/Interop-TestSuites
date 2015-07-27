@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXCSTOR\TestSuite\bin\Debug\MS-OXCSTOR_TestSuite.dll /runconfig:..\..\MS-OXCSTOR\MS-OXCSTOR.testsettings 
pause