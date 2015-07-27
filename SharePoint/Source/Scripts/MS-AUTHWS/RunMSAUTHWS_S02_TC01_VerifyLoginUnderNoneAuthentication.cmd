@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_AUTHWS.S02_LoginApplicationUnderNoneAuthentication.MSAUTHWS_S02_TC01_VerifyLoginUnderNoneAuthentication /testcontainer:..\..\MS-AUTHWS\TestSuite\bin\Debug\MS-AUTHWS_TestSuite.dll /runconfig:..\..\MS-AUTHWS\MS-AUTHWS.testsettings /unique
pause