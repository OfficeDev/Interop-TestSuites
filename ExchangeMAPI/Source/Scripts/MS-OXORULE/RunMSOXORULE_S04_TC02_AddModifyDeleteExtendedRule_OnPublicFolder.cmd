@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXORULE.S04_ProcessRulesOnPublicFolder.MSOXORULE_S04_TC02_AddModifyDeleteExtendedRule_OnPublicFolder /testcontainer:..\..\MS-OXORULE\TestSuite\bin\Debug\MS-OXORULE_TestSuite.dll /runconfig:..\..\MS-OXORULE\MS-OXORULE.testsettings /unique
pause