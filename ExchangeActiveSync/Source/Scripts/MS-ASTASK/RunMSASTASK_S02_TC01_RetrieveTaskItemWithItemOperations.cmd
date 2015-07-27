@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASTASK.S02_ItemOperationsCommand.MSASTASK_S02_TC01_RetrieveTaskItemWithItemOperations /testcontainer:..\..\MS-ASTASK\TestSuite\bin\Debug\MS-ASTASK_TestSuite.dll /runconfig:..\..\MS-ASTASK\MS-ASTASK.testsettings /unique
pause