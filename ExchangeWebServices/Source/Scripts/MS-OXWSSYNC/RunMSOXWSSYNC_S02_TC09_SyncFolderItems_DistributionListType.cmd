@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSSYNC.S02_SyncFolderItems.MSOXWSSYNC_S02_TC09_SyncFolderItems_DistributionListType /testcontainer:..\..\MS-OXWSSYNC\TestSuite\bin\Debug\MS-OXWSSYNC_TestSuite.dll /runconfig:..\..\MS-OXWSSYNC\MS-OXWSSYNC.testsettings /unique
pause