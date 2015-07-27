@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASCAL.S02_MeetingElement.MSASCAL_S02_TC02_MeetingDeclined /testcontainer:..\..\MS-ASCAL\TestSuite\bin\Debug\MS-ASCAL_TestSuite.dll /runconfig:..\..\MS-ASCAL\MS-ASCAL.testsettings /unique
pause