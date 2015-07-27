@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASPROV.S02_RemoteWipe.MSASPROV_S02_TC01_RemoteWipe /testcontainer:..\..\MS-ASPROV\TestSuite\bin\Debug\MS-ASPROV_TestSuite.dll /runconfig:..\..\MS-ASPROV\MS-ASPROV.testsettings /unique
pause