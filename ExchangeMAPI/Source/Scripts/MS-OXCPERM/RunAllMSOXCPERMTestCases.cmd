@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXCPERM\TestSuite\bin\Debug\MS-OXCPERM_TestSuite.dll /runconfig:..\..\MS-OXCPERM\MS-OXCPERM.testsettings 
pause