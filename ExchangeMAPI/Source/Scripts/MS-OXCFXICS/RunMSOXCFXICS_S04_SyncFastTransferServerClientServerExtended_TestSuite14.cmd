@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCFXICS.S04_SyncFastTransferServerClientServerExtended_TestSuite.MSOXCFXICS_S04_SyncFastTransferServerClientServerExtended_TestSuite14 /testcontainer:..\..\MS-OXCFXICS\TestSuite\bin\Debug\MS-OXCFXICS_TestSuite.dll /runconfig:..\..\MS-OXCFXICS\MS-OXCFXICS.testsettings /unique
pause