@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WSSREST.S02_RetrieveCSDLDocument.MSWSSREST_S02_TC01_RetrieveACSDLDocument /testcontainer:..\..\MS-WSSREST\TestSuite\bin\Debug\MS-WSSREST_TestSuite.dll /runconfig:..\..\MS-WSSREST\MS-WSSREST.testsettings /unique
pause