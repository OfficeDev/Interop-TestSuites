@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASNOTE.S02_SearchCommand.MSASNOTE_S02_TC01_Search_GetZeroOrMoreNotes /testcontainer:..\..\MS-ASNOTE\TestSuite\bin\Debug\MS-ASNOTE_TestSuite.dll /runconfig:..\..\MS-ASNOTE\MS-ASNOTE.testsettings /unique
pause