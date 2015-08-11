namespace Microsoft.Protocols.TestSuites.MS_ASRM
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The interface of MS-ASRMAdapter class.
    /// </summary>
    public interface IMS_ASRMAdapter : IAdapter
    {
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }

        /// <summary>
        /// Sync data from the server.
        /// </summary>
        /// <param name="syncRequest">The request for Sync command.</param>
        /// <returns>The Sync result which is returned from server.</returns>
        SyncStore Sync(SyncRequest syncRequest);

        /// <summary>
        /// Find an e-mail with specific subject.
        /// </summary>
        /// <param name="request">The request for Sync command.</param>
        /// <param name="subject">The subject of the e-mail to find.</param>
        /// <param name="isRetryNeeded">A boolean value specifies whether need retry.</param>
        /// <returns>The Sync result.</returns>
        Sync SyncEmail(SyncRequest request, string subject, bool isRetryNeeded);

        /// <summary>
        /// Fetch all information about exchange object.
        /// </summary>
        /// <param name="itemOperationsRequest">The request for ItemOperations command.</param>
        /// <returns>The ItemOperations result which is returned from server.</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest);

        /// <summary>
        /// Search items on server.
        /// </summary>
        /// <param name="searchRequest">The request for Search command.</param>
        /// <returns>The Search result which is returned from server.</returns>
        SearchStore Search(SearchRequest searchRequest);

        /// <summary>
        /// Gets the RightsManagementInformation by Settings command.
        /// </summary>
        /// <returns>The Settings response which is returned from the server.</returns>
        SettingsResponse Settings();

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="sendMailRequest">The request for SendMail command.</param>
        /// <returns>The SendMail response which is returned from server.</returns>
        SendMailResponse SendMail(SendMailRequest sendMailRequest);

        /// <summary>
        /// Reply to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartReplyRequest">The request for SmartReply command.</param>
        /// <returns>The SmartReply response which is returned from server.</returns>
        SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest);

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartForwardRequest">The request for SmartForward command.</param>
        /// <returns>The SmartForward response which is returned from server.</returns>
        SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest);

        /// <summary>
        /// Synchronize the collection hierarchy.
        /// </summary>
        /// <param name="request">The request for FolderSync command.</param>
        /// <returns>The FolderSync response which is returned from server.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest request);

        /// <summary>
        /// Change user to call ActiveSync command.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}