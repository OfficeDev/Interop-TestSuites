@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASTASK.S01_SyncCommand.MSASTASK_S01_TC06_CreateTaskItemRecursWeekly /testcontainer:..\..\MS-ASTASK\TestSuite\bin\Debug\MS-ASTASK_TestSuite.dll /runconfig:..\..\MS-ASTASK\MS-ASTASK.testsettings /unique
pause