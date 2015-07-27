@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_SHDACCWS.S01_VerifyIsSingleClient.MSSHDACCWS_S01_TC04_FileNotExistOnServer /testcontainer:..\..\MS-SHDACCWS\TestSuite\bin\Debug\MS-SHDACCWS_TestSuite.dll /runconfig:..\..\MS-SHDACCWS\MS-SHDACCWS.testsettings /unique
pause