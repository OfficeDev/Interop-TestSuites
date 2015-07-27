@echo off
pushd %~dp0

"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\MS-ASAIRS\TestSuite\bin\Debug\MS-ASAIRS_TestSuite.dll /testcontainer:..\MS-ASCAL\TestSuite\bin\Debug\MS-ASCAL_TestSuite.dll /testcontainer:..\MS-ASCMD\TestSuite\bin\Debug\MS-ASCMD_TestSuite.dll /testcontainer:..\MS-ASCNTC\TestSuite\bin\Debug\MS-ASCNTC_TestSuite.dll /testcontainer:..\MS-ASCON\TestSuite\bin\Debug\MS-ASCON_TestSuite.dll /testcontainer:..\MS-ASDOC\TestSuite\bin\Debug\MS-ASDOC_TestSuite.dll /testcontainer:..\MS-ASEMAIL\TestSuite\bin\Debug\MS-ASEMAIL_TestSuite.dll /testcontainer:..\MS-ASHTTP\TestSuite\bin\Debug\MS-ASHTTP_TestSuite.dll /testcontainer:..\MS-ASNOTE\TestSuite\bin\Debug\MS-ASNOTE_TestSuite.dll /testcontainer:..\MS-ASPROV\TestSuite\bin\Debug\MS-ASPROV_TestSuite.dll /testcontainer:..\MS-ASRM\TestSuite\bin\Debug\MS-ASRM_TestSuite.dll /testcontainer:..\MS-ASTASK\TestSuite\bin\Debug\MS-ASTASK_TestSuite.dll /runconfig:..\ExchangeServerEASProtocolTestSuites.testsettings

popd
pause