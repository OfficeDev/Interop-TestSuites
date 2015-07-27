@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMSG.S04_MoveEmailMessage.MSOXWSMSG_S04_TC01_MoveMessage /testcontainer:..\..\MS-OXWSMSG\TestSuite\bin\Debug\MS-OXWSMSG_TestSuite.dll /runconfig:..\..\MS-OXWSMSG\MS-OXWSMSG.testsettings /unique
pause