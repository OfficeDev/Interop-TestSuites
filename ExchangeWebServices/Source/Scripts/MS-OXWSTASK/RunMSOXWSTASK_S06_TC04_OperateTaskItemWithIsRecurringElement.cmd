@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSTASK.S06_OperateTaskItemWithOptionalElements.MSOXWSTASK_S06_TC04_OperateTaskItemWithIsRecurringElement /testcontainer:..\..\MS-OXWSTASK\TestSuite\bin\Debug\MS-OXWSTASK_TestSuite.dll /runconfig:..\..\MS-OXWSTASK\MS-OXWSTASK.testsettings /unique
pause