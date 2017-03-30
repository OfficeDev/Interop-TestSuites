@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_WSSREST.S01_ManageListItem.MSWSSREST_S01_TC01_CreateAndUpdateAttachmentInListItem /testcontainer:..\..\MS-WSSREST\TestSuite\bin\Debug\MS-WSSREST_TestSuite.dll /runconfig:..\..\MS-WSSREST\MS-WSSREST.testsettings /unique
pause