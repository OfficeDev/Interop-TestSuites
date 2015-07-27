//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using Microsoft.Protocols.TestTools;
    
    /// <summary>
    /// The SUT control adapter interface which is used in the test suite to carry out various operations related with SUT settings.
    /// </summary>
    public interface IMS_OXCRPCSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// A method used to disable asynchronous RPC notifications on the server.
        /// </summary>
        [MethodHelp(@"Disable asynchronous RPC on the server." + "\r\n" +
                    @"If the SUT is a Microsoft Exchange server, create a key named ""Async Rpc Notify Enabled"" of REG_DWORD type in ""HKLM\SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersSystem"" and set key value to 0, then open service manager and restart the ""Microsoft Exchange Information Store"" service.")]
        void DisableAsyncRPCNotification();

        /// <summary>
        /// A method used to enable asynchronous RPC notifications on the server.
        /// </summary>
        [MethodHelp(@"Enable asynchronous RPC on the server." + "\r\n" +
                    @"If the SUT is Microsoft Exchange server, set the ""Async Rpc Notify Enabled"" register key to 1, the regkey path is ""HKLM\SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersSystem"", then open service manager and restart the ""Microsoft Exchange Information Store"" service.")]
        void EnableAsyncRPCNotification();

        /// <summary>
        /// A method used to send an email to the user (defined by the ""AdminUserName"" property in the ptfconfig file).
        /// </summary>
        /// <returns>A Boolean value indicates whether create mail successfully.</returns>
        [MethodHelp(@"Send an email from the AdminUser mailbox to itself. The value of AdminUser is defined in the AdminUserName property in the MS-OXCRPC_TestSuite.deployment.ptfconfig file. " +
                    @"The body of the email can be blank. " +
                    " TRUE means an email was sent and received by AdminUser successfully." +
                    " FALSE means the email was not sent successfully.")]
        bool CreateMailItem();

        /// <summary>
        /// This operation is to get the server's operating system version and operating system service pack information.
        /// The return value should be a String value that includes the server's operating system version and operating system service pack information.
        /// The Format of the version should be \"X.X.XXXX.X.X \". 
        /// The first number is major version number of the operating system of the server.
        /// The second number is minor version number of the operating system of the server.
        /// The third number is build number of the operating system of the server.
        /// The fourth number is major version number of the latest operating system service pack that is installed on server.
        /// The fifth number is minor version number of the latest operating system service pack that is installed on server.
        /// </summary>
        /// <returns>The server's operating system version and operating system service pack information.</returns>
        [MethodHelp("Get the server's operating system version and service pack information.\r\n" +
            "The return value should be a string value that includes the server's operating system version and service pack information.\r\n" +
            "The Format of the version should be \"X.X.XXXX.X.X \". The first number is the major version number of the server's operating system.\r\n" +
            "The second number is the minor version number of the server's operating system.\r\n " +
            "The third number is the build number of the server's operating system.\r\n " +
            "The fourth number is the major version number of the latest operating system service pack that is installed on server.\r\n " +
            "The fifth number is the minor version number of the latest operating system service pack that is installed on server.\r\n" +
            "The Windows platform can use the \"Get-WmiObject Win32_OperatingSystem\" PowerShell command to obtain the server's operating system version and service pack information.")]
        string GetOSVersions();
    }
}