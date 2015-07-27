@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /runconfig:..\..\MS-OXWSCONT\MS-OXWSCONT.testsettings 
pause