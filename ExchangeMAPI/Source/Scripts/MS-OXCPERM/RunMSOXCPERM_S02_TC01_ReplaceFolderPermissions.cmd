@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCPERM.S02_ModifyFolderPermissions.MSOXCPERM_S02_TC01_ReplaceFolderPermissions /testcontainer:..\..\MS-OXCPERM\TestSuite\bin\Debug\MS-OXCPERM_TestSuite.dll /runconfig:..\..\MS-OXCPERM\MS-OXCPERM.testsettings /unique
pause