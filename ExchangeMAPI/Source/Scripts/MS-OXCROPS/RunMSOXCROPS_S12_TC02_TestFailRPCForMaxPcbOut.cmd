@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCROPS.S12_NotificationROPs.MSOXCROPS_S12_TC02_TestFailRPCForMaxPcbOut /testcontainer:..\..\MS-OXCROPS\TestSuite\bin\Debug\MS-OXCROPS_TestSuite.dll /runconfig:..\..\MS-OXCROPS\MS-OXCROPS.testsettings /unique
pause