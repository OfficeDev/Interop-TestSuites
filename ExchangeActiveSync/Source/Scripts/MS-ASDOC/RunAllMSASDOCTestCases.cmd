@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ASDOC\TestSuite\bin\Debug\MS-ASDOC_TestSuite.dll /runconfig:..\..\MS-ASDOC\MS-ASDOC.testsettings 
pause