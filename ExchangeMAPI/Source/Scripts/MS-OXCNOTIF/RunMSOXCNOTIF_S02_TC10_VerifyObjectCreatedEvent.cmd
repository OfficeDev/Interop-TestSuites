@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCNOTIF.S02_SubscribeAndReceiveNotifications.MSOXCNOTIF_S02_TC10_VerifyObjectCreatedEvent /testcontainer:..\..\MS-OXCNOTIF\TestSuite\bin\Debug\MS-OXCNOTIF_TestSuite.dll /runconfig:..\..\MS-OXCNOTIF\MS-OXCNOTIF.testsettings /unique
pause