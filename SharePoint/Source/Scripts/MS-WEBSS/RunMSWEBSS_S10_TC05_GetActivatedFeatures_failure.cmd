@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WEBSS.S10_OperationsOnActivatedFeatures.MSWEBSS_S10_TC05_GetActivatedFeatures_failure /testcontainer:..\..\MS-WEBSS\TestSuite\bin\Debug\MS-WEBSS_TestSuite.dll /runconfig:..\..\MS-WEBSS\MS-WEBSS.testsettings /unique
pause