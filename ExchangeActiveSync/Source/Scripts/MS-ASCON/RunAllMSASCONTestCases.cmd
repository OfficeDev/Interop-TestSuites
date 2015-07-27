@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ASCON\TestSuite\bin\Debug\MS-ASCON_TestSuite.dll /runconfig:..\..\MS-ASCON\MS-ASCON.testsettings 
pause