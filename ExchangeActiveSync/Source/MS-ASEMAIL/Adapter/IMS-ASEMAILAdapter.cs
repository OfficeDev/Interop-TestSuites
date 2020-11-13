namespace Microsoft.Protocols.TestSuites.MS_ASEMAIL
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASEMAIL.
    /// </summary>
    public interface IMS_ASEMAILAdapter : IAdapter
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
        /// <param name="syncRequest">The request for sync operation.</param>
        /// <returns>The sync result which is returned from server.</returns>
        SyncStore Sync(SyncRequest syncRequest);

        /// <summary>
        /// Sync data from the server with an invalid sync request which contains additional element.
        /// </summary>
        /// <param name="syncRequest">The request for sync operation.</param>
        /// <param name="addElement">Additional element insert into normal sync request.</param>
        /// <param name="insertTag">Insert tag shows where the additional element should inserted.</param>
        /// <returns>The sync result which is returned from server.</returns>
        SendStringResponse InvalidSync(SyncRequest syncRequest, string addElement, string insertTag);

        /// <summary>
        /// Search items on server.
        /// </summary>
        /// <param name="searchRequest">The request for search operation.</param>
        /// <returns>The search response which is returned from the server.</returns>
        SearchResponse Search(SearchRequest searchRequest);

        /// <summary>
        /// Find items on server.
        /// </summary>
        /// <param name="findRequest">The request for find operation.</param>
        /// <returns>The find response which is returned from the server.</returns>
        FindResponse Find(FindRequest findRequest);

        /// <summary>
        /// Search data on the server with an invalid Search request which contains an E-mail Class element.
        /// </summary>
        /// <param name="searchRequest">The request for search operation.</param>
        /// <param name="emailClassElement">The email class element.</param>
        /// <returns>The search response which is returned from server.</returns>
        SendStringResponse InvalidSearch(SearchRequest searchRequest, string emailClassElement);

        /// <summary>
        /// MeetingResponse for accepting or declining a meeting request.
        /// </summary>
        /// <param name="meetingResponseRequest">The request for meeting.</param>
        /// <returns>The meeting response which is returned from server.</returns>
        MeetingResponseResponse MeetingResponse(MeetingResponseRequest meetingResponseRequest);

        /// <summary>
        /// Fetch all information about exchange object.
        /// </summary>
        /// <param name="itemOperationsRequest">The request for itemOperations.</param>
        /// <returns>The ItemOperations result which is returned from server.</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest);

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="sendMailRequest">The request for SendMail operation.</param>
        /// <returns>The SendMail response which is returned from the server.</returns>
        SendMailResponse SendMail(SendMailRequest sendMailRequest);

        /// <summary>
        /// Reply to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartReplyRequest">The request for SmartReply operation.</param>
        /// <returns>The SmartReply response which is returned from the server.</returns>
        SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest);

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartForwardRequest">The request for SmartForward operation.</param>
        /// <returns>The SmartForward response which is returned from the server.</returns>
        SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest);

        /// <summary>
        /// Synchronizes the collection hierarchy.
        /// </summary>
        /// <param name="folderSyncRequest">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>The FolderSync response which is returned from the server.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest);
 
        /// <summary>
        /// Change user to call active sync operation.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}