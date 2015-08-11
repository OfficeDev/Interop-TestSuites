namespace Microsoft.Protocols.TestSuites.MS_ASCON
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASCON.
    /// </summary>
    public interface IMS_ASCONAdapter : IAdapter
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
        /// Change the user authentication.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);

        /// <summary>
        /// Synchronizes changes in a collection between the client and the server.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <returns>The SyncStore result which is returned from server.</returns>
        SyncStore Sync(SyncRequest syncRequest);

        /// <summary>
        /// Find an email with specific subject.
        /// </summary>
        /// <param name="syncRequest">A SyncRequest object that contains the request information.</param>
        /// <param name="subject">The subject of the email to find.</param>
        /// <param name="isRetryNeeded">A boolean whether need retry.</param>
        /// <returns>The email with specific subject.</returns>
        Sync SyncEmail(SyncRequest syncRequest, string subject, bool isRetryNeeded);

        /// <summary>
        /// Find entries address book, mailbox, or document library.
        /// </summary>
        /// <param name="searchRequest">A SearchRequest object that contains the request information.</param>
        /// <param name="expectSuccess">Whether the Search command is expected to be successful.</param>
        /// <param name="itemsCount">The count of the items expected to be found.</param>
        /// <returns>The SearchStore result which is returned from server.</returns>
        SearchStore Search(SearchRequest searchRequest, bool expectSuccess, int itemsCount);

        /// <summary>
        /// Acts as a container for the Fetch element, the EmptyFolderContents element, and the Move element to provide batched online handling of these operations against the server.
        /// </summary>
        /// <param name="itemOperationsRequest">An ItemOperationsRequest object that contains the request information.</param>
        /// <returns>ItemOperations command response.</returns>
        ItemOperationsResponse ItemOperations(ItemOperationsRequest itemOperationsRequest);

        /// <summary>
        /// Synchronizes the collection hierarchy. 
        /// </summary>
        /// <param name="folderSyncRequest">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>FolderSync command response.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest);

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="sendMailRequest">A SendMailRequest object that contains the request information.</param>
        /// <returns>SendMail command response.</returns>
        SendMailResponse SendMail(SendMailRequest sendMailRequest);

        /// <summary>
        /// Replies to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartReplyRequest">A SmartReplyRequest object that contains the request information.</param>
        /// <returns>SmartReply command response.</returns>
        SmartReplyResponse SmartReply(SmartReplyRequest smartReplyRequest);

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="smartForwardRequest">A SmartForwardRequest object that contains the request information.</param>
        /// <returns>SmartForward command response.</returns>
        SmartForwardResponse SmartForward(SmartForwardRequest smartForwardRequest);

        /// <summary>
        /// Moves an item or items from one folder on the server to another.
        /// </summary>
        /// <param name="moveItemsRequest">A MoveItemsRequest object that contains the request information.</param>
        /// <returns>MoveItems command response.</returns>
        MoveItemsResponse MoveItems(MoveItemsRequest moveItemsRequest);

        /// <summary>
        /// Gets an estimate of the number of items in a collection or folder on the server that have to be synchronized.
        /// </summary>
        /// <param name="getItemEstimateRequest">A GetItemEstimateRequest object that contains the request information.</param>
        /// <returns>GetItemEstimate command response.</returns>
        GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest getItemEstimateRequest);
    }
}