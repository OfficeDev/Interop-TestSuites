@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_SITESS.S01_MigrateSite.MSSITESS_S01_TC07_ImportingFailureOverwriteFailure /testcontainer:..\..\MS-SITESS\TestSuite\bin\Debug\MS-SITESS_TestSuite.dll /runconfig:..\..\MS-SITESS\MS-SITESS.testsettings /unique
pause