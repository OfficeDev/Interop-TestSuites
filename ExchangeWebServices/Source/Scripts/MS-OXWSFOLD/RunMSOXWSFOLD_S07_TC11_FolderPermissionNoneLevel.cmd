@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSFOLD.S07_FolderPermission.MSOXWSFOLD_S07_TC11_FolderPermissionNoneLevel /testcontainer:..\..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /runconfig:..\..\MS-OXWSFOLD\MS-OXWSFOLD.testsettings /unique
pause