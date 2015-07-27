@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCFOLD.S04_OperateOnPublicFolder.MSOXCFOLD_S04_TC06_RopCopyPublicFolder /testcontainer:..\..\MS-OXCFOLD\TestSuite\bin\Debug\MS-OXCFOLD_TestSuite.dll /runconfig:..\..\MS-OXCFOLD\MS-OXCFOLD.testsettings /unique
pause