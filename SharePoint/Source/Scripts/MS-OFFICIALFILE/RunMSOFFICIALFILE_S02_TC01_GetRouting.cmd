@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OFFICIALFILE.S02_GetRoutingInfo.MSOFFICIALFILE_S02_TC01_GetRouting /testcontainer:..\..\MS-OFFICIALFILE\TestSuite\bin\Debug\MS-OFFICIALFILE_TestSuite.dll /runconfig:..\..\MS-OFFICIALFILE\MS-OFFICIALFILE.testsettings /unique
pause