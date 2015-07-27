@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASNOTE.S01_SyncCommand.MSASNOTE_S01_TC07_Sync_ChangeNote_Categories /testcontainer:..\..\MS-ASNOTE\TestSuite\bin\Debug\MS-ASNOTE_TestSuite.dll /runconfig:..\..\MS-ASNOTE\MS-ASNOTE.testsettings /unique
pause