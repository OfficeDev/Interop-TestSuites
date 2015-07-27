@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WOPI.MS_WOPI_S02_Coauth.TestCase_S02_TC05_JoinCoauthoringSession_Success /testcontainer:..\..\MS-WOPI\TestSuite\bin\Debug\MS-WOPI_TestSuite.dll /runconfig:..\..\MS-WOPI\MS-WOPI.testsettings /unique
pause