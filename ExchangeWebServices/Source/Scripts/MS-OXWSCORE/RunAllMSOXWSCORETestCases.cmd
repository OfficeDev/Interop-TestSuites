@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXWSCORE\TestSuite\bin\Debug\MS-OXWSCORE_TestSuite.dll /runconfig:..\..\MS-OXWSCORE\MS-OXWSCORE.testsettings 
pause