@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASCNTC.S03_Search.MSASCNTC_S03_TC01_Search_RetrieveContact /testcontainer:..\..\MS-ASCNTC\TestSuite\bin\Debug\MS-ASCNTC_TestSuite.dll /runconfig:..\..\MS-ASCNTC\MS-ASCNTC.testsettings /unique
pause