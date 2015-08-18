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