@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WWSP.S03_AlterToDo.MSWWSP_S03_TC01_AlterToDo_Success /testcontainer:..\..\MS-WWSP\TestSuite\bin\Debug\MS-WWSP_TestSuite.dll /runconfig:..\..\MS-WWSP\MS-WWSP.testsettings /unique
pause