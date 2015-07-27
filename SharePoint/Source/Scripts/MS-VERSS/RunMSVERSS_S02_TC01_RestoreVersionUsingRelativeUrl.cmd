@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_VERSS.S02_RestoreVersion.MSVERSS_S02_TC01_RestoreVersionUsingRelativeUrl /testcontainer:..\..\MS-VERSS\TestSuite\bin\Debug\MS-VERSS_TestSuite.dll /runconfig:..\..\MS-VERSS\MS-VERSS.testsettings /unique
pause