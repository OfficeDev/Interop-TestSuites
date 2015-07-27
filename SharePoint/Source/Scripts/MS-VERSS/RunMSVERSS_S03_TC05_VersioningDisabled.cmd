@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_VERSS.S03_ErrorConditions.MSVERSS_S03_TC05_VersioningDisabled /testcontainer:..\..\MS-VERSS\TestSuite\bin\Debug\MS-VERSS_TestSuite.dll /runconfig:..\..\MS-VERSS\MS-VERSS.testsettings /unique
pause