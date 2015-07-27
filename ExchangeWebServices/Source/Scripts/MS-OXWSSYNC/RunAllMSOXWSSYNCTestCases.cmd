@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSSYNC\TestSuite\bin\Debug\MS-OXWSSYNC_TestSuite.dll /runconfig:..\..\MS-OXWSSYNC\MS-OXWSSYNC.testsettings 
pause