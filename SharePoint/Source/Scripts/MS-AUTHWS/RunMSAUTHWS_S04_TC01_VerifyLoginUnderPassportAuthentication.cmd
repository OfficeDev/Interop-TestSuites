@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_AUTHWS.S04_LoginApplicationUnderPassportAuthentication.MSAUTHWS_S04_TC01_VerifyLoginUnderPassportAuthentication /testcontainer:..\..\MS-AUTHWS\TestSuite\bin\Debug\MS-AUTHWS_TestSuite.dll /runconfig:..\..\MS-AUTHWS\MS-AUTHWS.testsettings /unique
pause