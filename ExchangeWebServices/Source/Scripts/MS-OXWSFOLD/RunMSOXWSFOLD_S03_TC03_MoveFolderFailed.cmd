@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSFOLD.S03_MoveFolder.MSOXWSFOLD_S03_TC03_MoveFolderFailed /testcontainer:..\..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /runconfig:..\..\MS-OXWSFOLD\MS-OXWSFOLD.testsettings /unique
pause