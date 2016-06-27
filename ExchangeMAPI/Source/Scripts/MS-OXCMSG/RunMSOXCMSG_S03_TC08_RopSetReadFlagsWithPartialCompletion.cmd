@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCMSG.S03_SetMessageFlags.MSOXCMSG_S03_TC08_RopSetReadFlagsWithPartialCompletion /testcontainer:..\..\MS-OXCMSG\TestSuite\bin\Debug\MS-OXCMSG_TestSuite.dll /runconfig:..\..\MS-OXCMSG\MS-OXCMSG.testsettings /unique
pause