@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASNOTE.S03_ItemOperationsCommand.MSASNOTE_S03_TC01_ItemOperations_GetZeroOrMoreNotes /testcontainer:..\..\MS-ASNOTE\TestSuite\bin\Debug\MS-ASNOTE_TestSuite.dll /runconfig:..\..\MS-ASNOTE\MS-ASNOTE.testsettings /unique
pause