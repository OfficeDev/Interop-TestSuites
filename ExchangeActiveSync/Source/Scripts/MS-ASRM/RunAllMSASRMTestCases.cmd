@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ASRM\TestSuite\bin\Debug\MS-ASRM_TestSuite.dll /runconfig:..\..\MS-ASRM\MS-ASRM.testsettings 
pause