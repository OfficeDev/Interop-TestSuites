@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_VERSS.S01_DeleteVersion.MSVERSS_S01_TC01_DeleteVersionUsingRelativeUrl /testcontainer:..\..\MS-VERSS\TestSuite\bin\Debug\MS-VERSS_TestSuite.dll /runconfig:..\..\MS-VERSS\MS-VERSS.testsettings /unique
pause