@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OFFICIALFILE.S04_GetHoldsInfo.MSOFFICIALFILE_S04_TC01_GetHoldsInfo /testcontainer:..\..\MS-OFFICIALFILE\TestSuite\bin\Debug\MS-OFFICIALFILE_TestSuite.dll /runconfig:..\..\MS-OFFICIALFILE\MS-OFFICIALFILE.testsettings /unique
pause