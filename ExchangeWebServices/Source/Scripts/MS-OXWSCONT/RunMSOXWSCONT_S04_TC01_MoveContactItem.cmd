@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCONT.S04_MoveContactItem.MSOXWSCONT_S04_TC01_MoveContactItem /testcontainer:..\..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /runconfig:..\..\MS-OXWSCONT\MS-OXWSCONT.testsettings /unique
pause