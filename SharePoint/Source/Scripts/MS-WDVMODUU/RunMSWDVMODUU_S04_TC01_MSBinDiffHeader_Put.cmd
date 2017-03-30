@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WDVMODUU.S04_MSBinDiffHeader.MSWDVMODUU_S04_TC01_MSBinDiffHeader_Put /testcontainer:..\..\MS-WDVMODUU\TestSuite\bin\Debug\MS-WDVMODUU_TestSuite.dll /runconfig:..\..\MS-WDVMODUU\MS-WDVMODUU.testsettings /unique
pause