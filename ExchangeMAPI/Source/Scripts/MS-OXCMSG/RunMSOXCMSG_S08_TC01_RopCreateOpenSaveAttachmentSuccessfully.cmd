@echo off
pushd %~dp0
"%VS120COMNTOOLS%..\IDE\mstest" /test:Microsoft.Protocols.TestSuites.MS_OXCMSG.S08_RopAttachment.MSOXCMSG_S08_TC01_RopCreateOpenSaveAttachmentSuccessfully /testcontainer:..\..\MS-OXCMSG\TestSuite\bin\Debug\MS-OXCMSG_TestSuite.dll /runconfig:..\..\MS-OXCMSG\MS-OXCMSG.testsettings /unique
pause