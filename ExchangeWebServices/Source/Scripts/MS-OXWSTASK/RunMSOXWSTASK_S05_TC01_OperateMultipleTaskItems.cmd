@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSTASK.S05_OperateMultipleTaskItems.MSOXWSTASK_S05_TC01_OperateMultipleTaskItems /testcontainer:..\..\MS-OXWSTASK\TestSuite\bin\Debug\MS-OXWSTASK_TestSuite.dll /runconfig:..\..\MS-OXWSTASK\MS-OXWSTASK.testsettings /unique
pause