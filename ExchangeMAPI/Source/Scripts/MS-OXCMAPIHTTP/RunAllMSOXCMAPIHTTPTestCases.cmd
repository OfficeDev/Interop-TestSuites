@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /testcontainer:..\..\MS-OXCMAPIHTTP\TestSuite\bin\Debug\MS-OXCMAPIHTTP_TestSuite.dll /runconfig:..\..\MS-OXCMAPIHTTP\MS-OXCMAPIHTTP.testsettings 
pause