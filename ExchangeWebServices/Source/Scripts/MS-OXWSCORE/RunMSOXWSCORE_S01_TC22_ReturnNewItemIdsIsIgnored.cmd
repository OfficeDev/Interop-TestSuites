@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCORE.S01_ManageBaseItems.MSOXWSCORE_S01_TC22_ReturnNewItemIdsIsIgnored /testcontainer:..\..\MS-OXWSCORE\TestSuite\bin\Debug\MS-OXWSCORE_TestSuite.dll /runconfig:..\..\MS-OXWSCORE\MS-OXWSCORE.testsettings /unique
pause