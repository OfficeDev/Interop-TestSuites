@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCORE.S03_ManageDistributionListsItems.MSOXWSCORE_S03_TC17_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean /testcontainer:..\..\MS-OXWSCORE\TestSuite\bin\Debug\MS-OXWSCORE_TestSuite.dll /runconfig:..\..\MS-OXWSCORE\MS-OXWSCORE.testsettings /unique
pause