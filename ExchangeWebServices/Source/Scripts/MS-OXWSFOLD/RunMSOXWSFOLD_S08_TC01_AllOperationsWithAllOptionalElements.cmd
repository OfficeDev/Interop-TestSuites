@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXWSFOLD.S08_OptionalElements.MSOXWSFOLD_S08_TC01_AllOperationsWithAllOptionalElements /testcontainer:..\..\MS-OXWSFOLD\TestSuite\bin\Debug\MS-OXWSFOLD_TestSuite.dll /runconfig:..\..\MS-OXWSFOLD\MS-OXWSFOLD.testsettings /unique
pause