@echo off
pushd %~dp0

"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\MS-OXCFOLD\TestSuite\bin\Debug\MS-OXCFOLD_TestSuite.dll /testcontainer:..\MS-OXCFXICS\TestSuite\bin\Debug\MS-OXCFXICS_TestSuite.dll /testcontainer:..\MS-OXCMAPIHTTP\TestSuite\bin\Debug\MS-OXCMAPIHTTP_TestSuite.dll /testcontainer:..\MS-OXCMSG\TestSuite\bin\Debug\MS-OXCMSG_TestSuite.dll /testcontainer:..\MS-OXCNOTIF\TestSuite\bin\Debug\MS-OXCNOTIF_TestSuite.dll /testcontainer:..\MS-OXCPERM\TestSuite\bin\Debug\MS-OXCPERM_TestSuite.dll /testcontainer:..\MS-OXCPRPT\TestSuite\bin\Debug\MS-OXCPRPT_TestSuite.dll /testcontainer:..\MS-OXCROPS\TestSuite\bin\Debug\MS-OXCROPS_TestSuite.dll /testcontainer:..\MS-OXCRPC\TestSuite\bin\Debug\MS-OXCRPC_TestSuite.dll /testcontainer:..\MS-OXCSTOR\TestSuite\bin\Debug\MS-OXCSTOR_TestSuite.dll /testcontainer:..\MS-OXCTABL\TestSuite\bin\Debug\MS-OXCTABL_TestSuite.dll /testcontainer:..\MS-OXNSPI\TestSuite\bin\Debug\MS-OXNSPI_TestSuite.dll /testcontainer:..\MS-OXORULE\TestSuite\bin\Debug\MS-OXORULE_TestSuite.dll /runconfig:..\ExchangeMAPIProtocolTestSuites.testsettings

popd
pause