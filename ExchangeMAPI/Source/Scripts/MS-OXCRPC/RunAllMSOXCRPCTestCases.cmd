@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXCRPC\TestSuite\bin\Debug\MS-OXCRPC_TestSuite.dll /runconfig:..\..\MS-OXCRPC\MS-OXCRPC.testsettings 
pause