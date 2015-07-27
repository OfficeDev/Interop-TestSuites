@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCPERM.S03_NegativeOrErrorValidation.MSOXCPERM_S03_TC01_SetUnexpectedPermissionFlags /testcontainer:..\..\MS-OXCPERM\TestSuite\bin\Debug\MS-OXCPERM_TestSuite.dll /runconfig:..\..\MS-OXCPERM\MS-OXCPERM.testsettings /unique
pause