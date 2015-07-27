@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSTASK.S01_CreateGetDeleteTaskItem.MSOXWSTASK_S01_TC12_VerifyTaskStatusType /testcontainer:..\..\MS-OXWSTASK\TestSuite\bin\Debug\MS-OXWSTASK_TestSuite.dll /runconfig:..\..\MS-OXWSTASK\MS-OXWSTASK.testsettings /unique
pause