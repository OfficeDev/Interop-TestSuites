@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP.S02_RequestTypesForAddressBookServerEndpoint.MSOXCMAPIHTTP_S02_TC19_TestFlagWithPtypErrorCode /testcontainer:..\..\MS-OXCMAPIHTTP\TestSuite\bin\Debug\MS-OXCMAPIHTTP_TestSuite.dll /runconfig:..\..\MS-OXCMAPIHTTP\MS-OXCMAPIHTTP.testsettings /unique
pause