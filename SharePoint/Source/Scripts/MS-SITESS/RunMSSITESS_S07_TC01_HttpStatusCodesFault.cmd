@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_SITESS.S07_HTTPStatusCode.MSSITESS_S07_TC01_HttpStatusCodesFault /testcontainer:..\..\MS-SITESS\TestSuite\bin\Debug\MS-SITESS_TestSuite.dll /runconfig:..\..\MS-SITESS\MS-SITESS.testsettings /unique
pause