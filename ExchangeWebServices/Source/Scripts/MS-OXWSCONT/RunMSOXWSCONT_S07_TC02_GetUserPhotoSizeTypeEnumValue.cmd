@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSCONT.S07_SetGetUserPhoto.MSOXWSCONT_S07_TC02_GetUserPhotoSizeTypeEnumValue  /testcontainer:..\..\MS-OXWSCONT\TestSuite\bin\Debug\MS-OXWSCONT_TestSuite.dll /runconfig:..\..\MS-OXWSCONT\MS-OXWSCONT.testsettings /unique
pause