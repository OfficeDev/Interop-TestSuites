@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMTGS.S03_CopyCalendarRelatedItem.MSOXWSMTGS_S03_TC05_CopyItemErrorCalendarCannotUseIdForRecurringMasterId /testcontainer:..\..\MS-OXWSMTGS\TestSuite\bin\Debug\MS-OXWSMTGS_TestSuite.dll /runconfig:..\..\MS-OXWSMTGS\MS-OXWSMTGS.testsettings /unique
pause