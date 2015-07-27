@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSMTGS.S05_OperateMultipleCalendarRelatedItems.MSOXWSMTGS_S05_TC04_MoveMultipleCalendarItems /testcontainer:..\..\MS-OXWSMTGS\TestSuite\bin\Debug\MS-OXWSMTGS_TestSuite.dll /runconfig:..\..\MS-OXWSMTGS\MS-OXWSMTGS.testsettings /unique
pause