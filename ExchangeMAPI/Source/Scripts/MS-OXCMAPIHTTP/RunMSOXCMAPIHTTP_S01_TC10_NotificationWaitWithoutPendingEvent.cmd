@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP.S01_RequestTypesForMailboxServerEndpoint.MSOXCMAPIHTTP_S01_TC10_NotificationWaitWithoutPendingEvent /testcontainer:..\..\MS-OXCMAPIHTTP\TestSuite\bin\Debug\MS-OXCMAPIHTTP_TestSuite.dll /runconfig:..\..\MS-OXCMAPIHTTP\MS-OXCMAPIHTTP.testsettings /unique
pause