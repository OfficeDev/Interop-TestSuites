@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASTASK.S03_SearchCommand.MSASTASK_S03_TC02_RetrieveInvalidTaskItemWithSearch /testcontainer:..\..\MS-ASTASK\TestSuite\bin\Debug\MS-ASTASK_TestSuite.dll /runconfig:..\..\MS-ASTASK\MS-ASTASK.testsettings /unique
pause