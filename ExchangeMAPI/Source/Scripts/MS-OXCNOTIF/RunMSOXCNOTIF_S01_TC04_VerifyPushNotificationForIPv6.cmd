@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCNOTIF.S01_ReceivePendingNotifications.MSOXCNOTIF_S01_TC04_VerifyPushNotificationForIPv6 /testcontainer:..\..\MS-OXCNOTIF\TestSuite\bin\Debug\MS-OXCNOTIF_TestSuite.dll /runconfig:..\..\MS-OXCNOTIF\MS-OXCNOTIF.testsettings /unique
pause