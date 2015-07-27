@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASRM.S04_SmartForward.MSASRM_S04_TC09_SmartForward_Protected_OriginalMessage_NotProtected /testcontainer:..\..\MS-ASRM\TestSuite\bin\Debug\MS-ASRM_TestSuite.dll /runconfig:..\..\MS-ASRM\MS-ASRM.testsettings /unique
pause