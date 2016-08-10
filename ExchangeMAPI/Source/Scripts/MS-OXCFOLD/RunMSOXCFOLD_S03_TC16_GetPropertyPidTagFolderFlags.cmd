@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCFOLD.S03_FolderInformation.MSOXCFOLD_S03_TC16_GetPropertyPidTagFolderFlags /testcontainer:..\..\MS-OXCFOLD\TestSuite\bin\Debug\MS-OXCFOLD_TestSuite.dll /runconfig:..\..\MS-OXCFOLD\MS-OXCFOLD.testsettings /unique
pause