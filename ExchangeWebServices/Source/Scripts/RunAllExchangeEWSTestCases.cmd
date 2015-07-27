@echo off
pushd %~dp0

"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\MS-OXWSATT\TestSuite\bin\Debug\MS-OXWSATT_TestSuite.dll /testcontainer:..\MS-OXWSBTRF\TestSuite\bin\Debug\MS-OXWSBTRF_TestSuite.dll /testcontainer:..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /testcontainer:..\MS-OXWSCORE\TestSuite\bin\Debug\MS-OXWSCORE_TestSuite.dll /testcontainer:..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /testcontainer:..\MS-OXWSMSG\TestSuite\bin\Debug\MS-OXWSMSG_TestSuite.dll /testcontainer:..\MS-OXWSMTGS\TestSuite\bin\Debug\MS-OXWSMTGS_TestSuite.dll /testcontainer:..\MS-OXWSSYNC\TestSuite\bin\Debug\MS-OXWSSYNC_TestSuite.dll /testcontainer:..\MS-OXWSTASK\TestSuite\bin\Debug\MS-OXWSTASK_TestSuite.dll /runconfig:..\ExchangeServerEWSProtocolTestSuites.testsettings

popd
pause