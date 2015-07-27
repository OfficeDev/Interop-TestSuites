@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXCNOTIF\TestSuite\bin\Debug\MS-OXCNOTIF_TestSuite.dll /runconfig:..\..\MS-OXCNOTIF\MS-OXCNOTIF.testsettings 
pause