@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WOPI.MS_WOPI_S03_SchemaLock.TestCase_S03_TC05_ConvertToExclusiveLock_ExitCoauthSessionAsConvertToExclusiveFailed /testcontainer:..\..\MS-WOPI\TestSuite\bin\Debug\MS-WOPI_TestSuite.dll /runconfig:..\..\MS-WOPI\MS-WOPI.testsettings /unique
pause