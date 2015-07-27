@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCTABL.S02_RowRops_ParseQueryRowsAfterSortTable_TestSuite.MSOXCTABL_S02_RowRops_ParseQueryRowsAfterSortTable_TestSuite1 /testcontainer:..\..\MS-OXCTABL\TestSuite\bin\Debug\MS-OXCTABL_TestSuite.dll /runconfig:..\..\MS-OXCTABL\MS-OXCTABL.testsettings /unique
pause