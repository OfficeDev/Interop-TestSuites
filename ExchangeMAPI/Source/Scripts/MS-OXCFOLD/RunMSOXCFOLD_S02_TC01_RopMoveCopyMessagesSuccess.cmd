@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCFOLD.S02_MessageRopOperations.MSOXCFOLD_S02_TC01_RopMoveCopyMessagesSuccess /testcontainer:..\..\MS-OXCFOLD\TestSuite\bin\Debug\MS-OXCFOLD_TestSuite.dll /runconfig:..\..\MS-OXCFOLD\MS-OXCFOLD.testsettings /unique
pause