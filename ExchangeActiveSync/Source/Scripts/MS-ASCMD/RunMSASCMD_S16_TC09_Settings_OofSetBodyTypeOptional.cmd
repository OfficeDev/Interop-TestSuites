@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASCMD.S16_Settings.MSASCMD_S16_TC09_Settings_OofSetBodyTypeOptional /testcontainer:..\..\MS-ASCMD\TestSuite\bin\Debug\MS-ASCMD_TestSuite.dll /runconfig:..\..\MS-ASCMD\MS-ASCMD.testsettings /unique
pause