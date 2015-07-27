@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-SHDACCWS\TestSuite\bin\Debug\MS-SHDACCWS_TestSuite.dll /runconfig:..\..\MS-SHDACCWS\MS-SHDACCWS.testsettings 
pause