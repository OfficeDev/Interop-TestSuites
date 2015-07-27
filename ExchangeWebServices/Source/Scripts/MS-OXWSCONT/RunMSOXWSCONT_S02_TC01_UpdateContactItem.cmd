@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCONT.S02_UpdateContactItem.MSOXWSCONT_S02_TC01_UpdateContactItem /testcontainer:..\..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /runconfig:..\..\MS-OXWSCONT\MS-OXWSCONT.testsettings /unique
pause