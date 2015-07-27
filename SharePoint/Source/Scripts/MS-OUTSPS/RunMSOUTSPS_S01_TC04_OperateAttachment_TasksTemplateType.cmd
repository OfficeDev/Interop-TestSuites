@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OUTSPS.S01_OperateAttachment.MSOUTSPS_S01_TC04_OperateAttachment_TasksTemplateType /testcontainer:..\..\MS-OUTSPS\TestSuite\bin\Debug\MS-OUTSPS_TestSuite.dll /runconfig:..\..\MS-OUTSPS\MS-OUTSPS.testsettings /unique
pause