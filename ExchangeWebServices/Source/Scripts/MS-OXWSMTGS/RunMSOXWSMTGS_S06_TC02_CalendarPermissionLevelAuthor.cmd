@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMTGS.S06_CalendarFolderPermission.MSOXWSMTGS_S06_TC02_CalendarPermissionLevelAuthor /testcontainer:..\..\MS-OXWSMTGS\TestSuite\bin\Debug\MS-OXWSMTGS_TestSuite.dll /runconfig:..\..\MS-OXWSMTGS\MS-OXWSMTGS.testsettings /unique
pause