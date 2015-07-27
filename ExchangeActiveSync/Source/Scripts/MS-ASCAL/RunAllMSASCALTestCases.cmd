@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ASCAL\TestSuite\bin\Debug\MS-ASCAL_TestSuite.dll /runconfig:..\..\MS-ASCAL\MS-ASCAL.testsettings 
pause