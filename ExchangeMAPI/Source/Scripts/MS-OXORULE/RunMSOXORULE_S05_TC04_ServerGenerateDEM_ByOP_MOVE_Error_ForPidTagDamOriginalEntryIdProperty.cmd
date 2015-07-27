@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXORULE.S05_GenerateDAMAndDEM.MSOXORULE_S05_TC04_ServerGenerateDEM_ByOP_MOVE_Error_ForPidTagDamOriginalEntryIdProperty /testcontainer:..\..\MS-OXORULE\TestSuite\bin\Debug\MS-OXORULE_TestSuite.dll /runconfig:..\..\MS-OXORULE\MS-OXORULE.testsettings /unique
pause