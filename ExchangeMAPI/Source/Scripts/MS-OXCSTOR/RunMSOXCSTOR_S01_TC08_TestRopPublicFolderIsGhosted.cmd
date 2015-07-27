@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCSTOR.S01_PrivateMailboxLogon.MSOXCSTOR_S01_TC08_TestRopPublicFolderIsGhosted /testcontainer:..\..\MS-OXCSTOR\TestSuite\bin\Debug\MS-OXCSTOR_TestSuite.dll /runconfig:..\..\MS-OXCSTOR\MS-OXCSTOR.testsettings /unique
pause