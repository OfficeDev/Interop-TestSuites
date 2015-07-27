@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_MEETS.S03_MeetingFromICal.MSMEETS_S03_TC11_SetAttendeeResponseWithoutOptionalParameters /testcontainer:..\..\MS-MEETS\TestSuite\bin\Debug\MS-MEETS_TestSuite.dll /runconfig:..\..\MS-MEETS\MS-MEETS.testsettings /unique
pause