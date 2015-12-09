@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASAIRS.S06_MeetingResponseCommand.MSASAIRS_S06_TC01_MeetingResponseWithPlainTextBody /testcontainer:..\..\MS-ASAIRS\TestSuite\bin\Debug\MS-ASAIRS_TestSuite.dll /runconfig:..\..\MS-ASAIRS\MS-ASAIRS.testsettings /unique
pause