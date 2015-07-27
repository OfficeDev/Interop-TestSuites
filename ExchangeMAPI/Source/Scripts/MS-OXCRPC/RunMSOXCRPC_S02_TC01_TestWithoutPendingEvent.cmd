@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCRPC.S02_AsynchronousCall.MSOXCRPC_S02_TC01_TestWithoutPendingEvent /testcontainer:..\..\MS-OXCRPC\TestSuite\bin\Debug\MS-OXCRPC_TestSuite.dll /runconfig:..\..\MS-OXCRPC\MS-OXCRPC.testsettings /unique
pause