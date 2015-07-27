@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_AUTHWS.S03_LoginApplicationUnderWindowsAuthentication.MSAUTHWS_S03_TC01_VerifyLoginUnderWindowsAuthentication /testcontainer:..\..\MS-AUTHWS\TestSuite\bin\Debug\MS-AUTHWS_TestSuite.dll /runconfig:..\..\MS-AUTHWS\MS-AUTHWS.testsettings /unique
pause