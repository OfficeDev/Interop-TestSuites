@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OUTSPS.S02_OperateListItems.MSOUTSPS_S02_TC01_OperationListItemsForAppointment /testcontainer:..\..\MS-OUTSPS\TestSuite\bin\Debug\MS-OUTSPS_TestSuite.dll /runconfig:..\..\MS-OUTSPS\MS-OUTSPS.testsettings /unique
pause