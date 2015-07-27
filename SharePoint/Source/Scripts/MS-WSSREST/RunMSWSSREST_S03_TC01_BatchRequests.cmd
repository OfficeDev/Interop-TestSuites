@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WSSREST.S03_BatchRequests.MSWSSREST_S03_TC01_BatchRequests /testcontainer:..\..\MS-WSSREST\TestSuite\bin\Debug\MS-WSSREST_TestSuite.dll /runconfig:..\..\MS-WSSREST\MS-WSSREST.testsettings /unique
pause