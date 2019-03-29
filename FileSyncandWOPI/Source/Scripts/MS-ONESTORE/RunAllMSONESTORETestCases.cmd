@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-ONESTORE\TestSuite\bin\Debug\MS-ONESTORE_TestSuite.dll /runconfig:..\..\MS-ONESTORE\MS-ONESTORE.testsettings 
pause