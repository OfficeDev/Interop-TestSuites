@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-AUTHWS\TestSuite\bin\Debug\MS-AUTHWS_TestSuite.dll /runconfig:..\..\MS-AUTHWS\MS-AUTHWS.testsettings 
pause