@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASPROV.S01_AcknowledgePolicySettings.MSASPROV_S01_TC02_WithoutAcknowledgingSecurityPolicySettings /testcontainer:..\..\MS-ASPROV\TestSuite\bin\Debug\MS-ASPROV_TestSuite.dll /runconfig:..\..\MS-ASPROV\MS-ASPROV.testsettings /unique
pause