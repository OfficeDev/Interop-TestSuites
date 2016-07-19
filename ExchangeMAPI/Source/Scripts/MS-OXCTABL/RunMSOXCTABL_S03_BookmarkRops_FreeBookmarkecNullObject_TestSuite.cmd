@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCTABL.S03_BookmarkRops_FreeBookmarkecNullObject_TestSuite.MSOXCTABL_S03_BookmarkRops_FreeBookmarkecNullObject_TestSuite /testcontainer:..\..\MS-OXCTABL\TestSuite\bin\Debug\MS-OXCTABL_TestSuite.dll /runconfig:..\..\MS-OXCTABL\MS-OXCTABL.testsettings /unique
pause