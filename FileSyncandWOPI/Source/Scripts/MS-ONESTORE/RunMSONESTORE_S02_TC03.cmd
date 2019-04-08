@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ONESTORE.S02_OneNoteRevisionStore.MSONESTORE_S02_TC03_VerifyguildFile /testcontainer:..\..\MS-ONESTORE\TestSuite\bin\Debug\MS-ONESTORE_TestSuite.dll /runconfig:..\..\MS-ONESTORE\MS-ONESTORE.testsettings /unique
pause