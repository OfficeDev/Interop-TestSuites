@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_MEETS.S04_RecurringMeeting.MSMEETS_S04_TC01_RecurringMeetingOperations /testcontainer:..\..\MS-MEETS\TestSuite\bin\Debug\MS-MEETS_TestSuite.dll /runconfig:..\..\MS-MEETS\MS-MEETS.testsettings /unique
pause