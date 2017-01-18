@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB.MS_FSSHTTP_FSSHTTPB_S15_CreateFile.TestCase_S15_TC04_UploadContents_ExclusiveLockSuccess /testcontainer:..\..\MS-FSSHTTP-FSSHTTPB\TestSuite\bin\Debug\MS-FSSHTTP-FSSHTTPB_TestSuite.dll /runconfig:..\..\MS-FSSHTTP-FSSHTTPB\MS-FSSHTTP-FSSHTTPB.testsettings /unique
pause