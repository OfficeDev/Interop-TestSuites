namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASAIRS.
    /// </summary>
    public interface IMS_ASAIRSAdapter : IAdapter
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
        /// Changes user to call active sync operation.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);

        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending SyncRequest object.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <returns>A SyncStore object.</returns>
        SyncStore Sync(SyncRequest request);

        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw Sync request.</param>
        /// <returns>A SendStringResponse object.</returns>
        SendStringResponse Sync(string request);

        /// <summary>
        /// Retrieves an item from the server by sending ItemOperationsRequest object.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object that contains the request information.</param>
        /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
        /// <returns>An ItemOperationsStore object.</returns>
        ItemOperationsStore ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod);

        /// <summary>
        /// Retrieves an item from the server by sending the raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw ItemOperations request.</param>
        /// <returns>A SendStringResponse object.</returns>
        SendStringResponse ItemOperations(string request);

        /// <summary>
        /// Finds entries in an address book, mailbox or document library by sending SearchRequest object.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>A SearchStore object.</returns>
        SearchStore Search(SearchRequest request);

        /// <summary>
        /// Finds entries in an address book, mailbox or document library by sending raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw Search request.</param>
        /// <returns>A SendStringResponse object.</returns>
        SendStringResponse Search(string request);

        /// <summary>
        /// Synchronizes the collection hierarchy.
        /// </summary>
        /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>A FolderSyncResponse object.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest request);

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="request">A SendMailRequest object that contains the request information.</param>
        /// <returns>A SendMailResponse object.</returns>
        SendMailResponse SendMail(SendMailRequest request);
    }
}