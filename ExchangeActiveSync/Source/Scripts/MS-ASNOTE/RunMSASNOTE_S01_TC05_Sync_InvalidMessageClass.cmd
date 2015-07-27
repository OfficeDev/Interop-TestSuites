@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASNOTE.S01_SyncCommand.MSASNOTE_S01_TC05_Sync_InvalidMessageClass /testcontainer:..\..\MS-ASNOTE\TestSuite\bin\Debug\MS-ASNOTE_TestSuite.dll /runconfig:..\..\MS-ASNOTE\MS-ASNOTE.testsettings /unique
pause