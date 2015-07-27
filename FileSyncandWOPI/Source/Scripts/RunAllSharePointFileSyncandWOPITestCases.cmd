@echo off
pushd %~dp0

"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\MS-FSSHTTP-FSSHTTPB\TestSuite\bin\Debug\MS-FSSHTTP-FSSHTTPB_TestSuite.dll /testcontainer:..\MS-WOPI\TestSuite\bin\Debug\MS-WOPI_TestSuite.dll /runconfig:..\SharePointFileSyncandWOPIProtocolTestSuites.testsettings

popd
pause