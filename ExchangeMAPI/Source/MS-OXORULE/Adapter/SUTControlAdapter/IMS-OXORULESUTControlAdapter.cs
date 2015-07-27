//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The SUT adapter interface which is used by test cases in the test suite to send an email to the recipient.
    /// </summary>
    public interface IMS_OXORULESUTControlAdapter : IAdapter
    {
        /// <summary>
        /// Send an email message to the recipient.
        /// </summary>
        /// <param name="senderUserName">The sender's name.</param>
        /// <param name="senderPassword">The sender's password.</param>
        /// <param name="recipientUserName">The recipient's name.</param>
        /// <param name="subject">The email's subject.</param>
        [MethodHelp(@"Send an email from one user (senderUserName,senderPassword) to another user (recipientUserName) with the subject(subject).")]
        void SendMailToRecipient(string senderUserName, string senderPassword, string recipientUserName, string subject);
    }
}