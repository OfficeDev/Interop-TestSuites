//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT control adapter interface which is used in the test suite to carry out various operations related with SUT settings.
    /// </summary>
    public interface IMS_OXCMAPIHTTPSUTControlAdapter : IAdapter
    {
        /// <summary>
        /// A method used to send an email to the specified user account.
        /// </summary>
        /// <returns>A Boolean value indicates whether send mail successfully.</returns>
        [MethodHelp(@"Send an email from the AdminUser mailbox to itself. The value of AdminUser is defined in the AdminUserName property in the MS-OXCMAPIHTTP_TestSuite.deployment.ptfconfig file. " +
                    @"The body of the email can be blank. " +
                    " TRUE means an email was sent and received by AdminUser successfully." +
                    " FALSE means the email was not sent successfully.")]
        bool SendMailItem();
    }
}
