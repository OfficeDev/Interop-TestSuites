@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASPROV.S03_ProvisionNegative.MSASPROV_S03_TC03_VerifyProvisionStatus2 /testcontainer:..\..\MS-ASPROV\TestSuite\bin\Debug\MS-ASPROV_TestSuite.dll /runconfig:..\..\MS-ASPROV\MS-ASPROV.testsettings /unique
pause