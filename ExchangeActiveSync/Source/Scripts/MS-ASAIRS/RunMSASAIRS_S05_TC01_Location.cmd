@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ASAIRS.S05_Location.MSASAIRS_S05_TC01_Location /testcontainer:..\..\MS-ASAIRS\TestSuite\bin\Debug\MS-ASAIRS_TestSuite.dll /runconfig:..\..\MS-ASAIRS\MS-ASAIRS.testsettings /unique
pause