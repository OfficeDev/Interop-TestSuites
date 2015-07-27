@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCSTOR.S03_SyncUpReadAndUnreadInformation.MSOXCSTOR_S03_TC01_ReadPerUserInformationForPublicLogon /testcontainer:..\..\MS-OXCSTOR\TestSuite\bin\Debug\MS-OXCSTOR_TestSuite.dll /runconfig:..\..\MS-OXCSTOR\MS-OXCSTOR.testsettings /unique
pause