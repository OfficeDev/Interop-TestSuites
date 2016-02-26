@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMSG.S01_CreateGetDeleteEmailMessage.MSOXWSMSG_S01_TC08_GetMessageWithBodyTypeAndAdditionProperties /testcontainer:..\..\MS-OXWSMSG\TestSuite\bin\Debug\MS-OXWSMSG_TestSuite.dll /runconfig:..\..\MS-OXWSMSG\MS-OXWSMSG.testsettings /unique
pause