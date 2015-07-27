@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXNSPI.S02_ObtainDetailsInfo.MSOXNSPI_S02_TC04_UpdateStatAbsolutePosition /testcontainer:..\..\MS-OXNSPI\TestSuite\bin\Debug\MS-OXNSPI_TestSuite.dll /runconfig:..\..\MS-OXNSPI\MS-OXNSPI.testsettings /unique
pause