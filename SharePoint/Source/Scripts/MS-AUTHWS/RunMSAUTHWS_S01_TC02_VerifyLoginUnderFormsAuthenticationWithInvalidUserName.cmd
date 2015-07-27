@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_AUTHWS.S01_LoginApplicationUnderFormsAuthentication.MSAUTHWS_S01_TC02_VerifyLoginUnderFormsAuthenticationWithInvalidUserName /testcontainer:..\..\MS-AUTHWS\TestSuite\bin\Debug\MS-AUTHWS_TestSuite.dll /runconfig:..\..\MS-AUTHWS\MS-AUTHWS.testsettings /unique
pause