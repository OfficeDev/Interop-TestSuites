@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_ONESTORE.S01_TransmissionByFSSHTTP.MSONESTORE_S01_TC01_QueryOneFileContainsFileData /testcontainer:..\..\MS-ONESTORE\TestSuite\bin\Debug\MS-ONESTORE_TestSuite.dll /runconfig:..\..\MS-ONESTORE\MS-ONESTORE.testsettings /unique
pause