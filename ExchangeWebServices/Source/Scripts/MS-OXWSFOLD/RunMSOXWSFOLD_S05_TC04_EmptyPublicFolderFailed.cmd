@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSFOLD.S05_EmptyFolder.MSOXWSFOLD_S05_TC04_EmptyPublicFolderFailed /testcontainer:..\..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /runconfig:..\..\MS-OXWSFOLD\MS-OXWSFOLD.testsettings /unique
pause