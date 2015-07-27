@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSATT.S01_AttachmentProcessing.MSOXWSATT_S01_TC06_ProcessPostItemTypeItemAttachment /testcontainer:..\..\MS-OXWSATT\TestSuite\bin\Debug\MS-OXWSATT_TestSuite.dll /runconfig:..\..\MS-OXWSATT\MS-OXWSATT.testsettings /unique
pause