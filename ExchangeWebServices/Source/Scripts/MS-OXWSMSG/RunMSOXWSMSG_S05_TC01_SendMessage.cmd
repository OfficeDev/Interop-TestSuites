@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMSG.S05_SendEmailMessage.MSOXWSMSG_S05_TC01_SendMessage /testcontainer:..\..\MS-OXWSMSG\TestSuite\bin\Debug\MS-OXWSMSG_TestSuite.dll /runconfig:..\..\MS-OXWSMSG\MS-OXWSMSG.testsettings /unique
pause