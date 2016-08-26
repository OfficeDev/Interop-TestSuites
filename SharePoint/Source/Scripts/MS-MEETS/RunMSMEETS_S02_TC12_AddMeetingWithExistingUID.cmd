@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_MEETS.S02_Meeting.MSMEETS_S02_TC12_AddMeetingWithExistingUID /testcontainer:..\..\MS-MEETS\TestSuite\bin\Debug\MS-MEETS_TestSuite.dll /runconfig:..\..\MS-MEETS\MS-MEETS.testsettings /unique
pause