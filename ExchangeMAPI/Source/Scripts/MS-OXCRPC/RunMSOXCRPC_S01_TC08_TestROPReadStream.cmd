@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCRPC.S01_SynchronousCall.MSOXCRPC_S01_TC08_TestROPReadStream /testcontainer:..\..\MS-OXCRPC\TestSuite\bin\Debug\MS-OXCRPC_TestSuite.dll /runconfig:..\..\MS-OXCRPC\MS-OXCRPC.testsettings /unique
pause