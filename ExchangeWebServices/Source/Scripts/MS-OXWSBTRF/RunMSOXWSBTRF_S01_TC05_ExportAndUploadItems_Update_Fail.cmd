@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSBTRF.S01_ExportAndUploadItems.MSOXWSBTRF_S01_TC05_ExportAndUploadItems_Update_Fail /testcontainer:..\..\MS-OXWSBTRF\TestSuite\bin\Debug\MS-OXWSBTRF_TestSuite.dll /runconfig:..\..\MS-OXWSBTRF\MS-OXWSBTRF.testsettings /unique
pause