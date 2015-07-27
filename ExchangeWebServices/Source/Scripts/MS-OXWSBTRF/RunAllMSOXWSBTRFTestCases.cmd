@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSBTRF\TestSuite\bin\Debug\MS-OXWSBTRF_TestSuite.dll /runconfig:..\..\MS-OXWSBTRF\MS-OXWSBTRF.testsettings 
pause