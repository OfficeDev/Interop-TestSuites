@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMTGS.S01_CreateGetDeleteCalendarRelatedItem.MSOXWSMTGS_S01_TC27_CreateSingleCalendarItemWithReminders /testcontainer:..\..\MS-OXWSMTGS\TestSuite\bin\Debug\MS-OXWSMTGS_TestSuite.dll /runconfig:..\..\MS-OXWSMTGS\MS-OXWSMTGS.testsettings /unique
pause