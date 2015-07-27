@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_SITESS.S05_ExportWorkflowTemplate.MSSITESS_S05_TC03_ExportWorkflowTemplateInvalidTemplate /testcontainer:..\..\MS-SITESS\TestSuite\bin\Debug\MS-SITESS_TestSuite.dll /runconfig:..\..\MS-SITESS\MS-SITESS.testsettings /unique
pause