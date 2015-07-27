@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMSG.S07_OptionalElementsValidation.MSOXWSMSG_S07_TC02_VerifyMessageWithoutOptionalElements /testcontainer:..\..\MS-OXWSMSG\TestSuite\bin\Debug\MS-OXWSMSG_TestSuite.dll /runconfig:..\..\MS-OXWSMSG\MS-OXWSMSG.testsettings /unique
pause