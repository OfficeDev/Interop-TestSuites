@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCSTOR.S02_PublicFoldersLogon.MSOXCSTOR_S02_TC06_TestOperationsOnProperties /testcontainer:..\..\MS-OXCSTOR\TestSuite\bin\Debug\MS-OXCSTOR_TestSuite.dll /runconfig:..\..\MS-OXCSTOR\MS-OXCSTOR.testsettings /unique
pause