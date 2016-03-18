@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSFOLD.S01_CreateDeleteFolder.MSOXWSFOLD_S01_TC20_VerifyStorageQuotaSet /testcontainer:..\..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /runconfig:..\..\MS-OXWSFOLD\MS-OXWSFOLD.testsettings /unique
pause