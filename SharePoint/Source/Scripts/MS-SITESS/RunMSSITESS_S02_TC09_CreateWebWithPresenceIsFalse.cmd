@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_SITESS.S02_ManageSubSite.MSSITESS_S02_TC09_CreateWebWithPresenceIsFalse /testcontainer:..\..\MS-SITESS\TestSuite\bin\Debug\MS-SITESS_TestSuite.dll /runconfig:..\..\MS-SITESS\MS-SITESS.testsettings /unique
pause