namespace Microsoft.Protocols.TestSuites.MS_ASCNTC
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASCNTC.
    /// </summary>
    public interface IMS_ASCNTCAdapter : IAdapter
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
        /// Synchronizes the changes in a collection between the client and the server by sending SyncRequest object.
        /// </summary>
        /// <param name="request">A SyncRequest object which contains the request information.</param>
        /// <returns>A SyncStore object.</returns>
        SyncStore Sync(SyncRequest request);

        /// <summary>
        /// Finds entries in an address book, mailbox or document library by sending SearchRequest object.
        /// </summary>
        /// <param name="request">A SearchRequest object which contains the request information.</param>
        /// <returns>A SearchStore object.</returns>
        SearchStore Search(SearchRequest request);

        /// <summary>
        /// Retrieves an item from the server by sending ItemOperationsRequest object.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object which contains the request information.</param>
        /// <param name="deliveryMethod">The delivery method specifies what kind of response is accepted.</param>
        /// <returns>An ItemOperationsStore object.</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod);

        /// <summary>
        /// Synchronizes the collection hierarchy.
        /// </summary>
        /// <param name="request">A FolderSyncRequest object which contains the request information.</param>
        /// <returns>A FolderSyncResponse object.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest request);

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="request">A SendMailRequest object which contains the request information.</param>
        /// <returns>A SendMailResponse object.</returns>
        SendMailResponse SendMail(SendMailRequest request);

        /// <summary>
        /// Changes user to call ActiveSync operation.
        /// </summary>
        /// <param name="userName">The name of the user.</param>
        /// <param name="userPassword">The password of the user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
    }
}