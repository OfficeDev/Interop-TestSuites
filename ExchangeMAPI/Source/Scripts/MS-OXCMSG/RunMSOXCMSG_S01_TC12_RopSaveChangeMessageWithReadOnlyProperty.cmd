@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCMSG.S01_CreateAndSaveMessage.MSOXCMSG_S01_TC12_RopSaveChangeMessageWithReadOnlyProperty /testcontainer:..\..\MS-OXCMSG\TestSuite\bin\Debug\MS-OXCMSG_TestSuite.dll /runconfig:..\..\MS-OXCMSG\MS-OXCMSG.testsettings /unique
pause