@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCFOLD.S01_FolderRopOperations.MSOXCFOLD_S01_TC09_RopDeleteFolderToSoftDelete /testcontainer:..\..\MS-OXCFOLD\TestSuite\bin\Debug\MS-OXCFOLD_TestSuite.dll /runconfig:..\..\MS-OXCFOLD\MS-OXCFOLD.testsettings /unique
pause