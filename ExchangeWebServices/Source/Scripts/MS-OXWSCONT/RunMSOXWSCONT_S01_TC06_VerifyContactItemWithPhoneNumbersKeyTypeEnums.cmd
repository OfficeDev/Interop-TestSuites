@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCONT.S01_CreateGetDeleteContactItem.MSOXWSCONT_S01_TC06_VerifyContactItemWithPhoneNumbersKeyTypeEnums /testcontainer:..\..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /runconfig:..\..\MS-OXWSCONT\MS-OXWSCONT.testsettings /unique
pause