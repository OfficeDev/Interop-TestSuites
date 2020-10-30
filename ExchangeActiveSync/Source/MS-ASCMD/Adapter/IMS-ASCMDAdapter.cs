namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter interface which provides methods defined in MS-ASCMD.
    /// </summary>
    public interface IMS_ASCMDAdapter : IAdapter
    {
        #region Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        IXPathNavigable LastRawRequestXml { get; }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        IXPathNavigable LastRawResponseXml { get; }
        #endregion

        #region Protocol Operations
        /// <summary>
        /// Facilitates the discovery of core account configuration information by using the user's Simple Mail Transfer Protocol (SMTP) address as the primary input.
        /// </summary>
        /// <param name="request">An AutodiscoverRequest object that contains the request information.</param>
        /// <param name="contentType">Content Type that indicate the body's format.</param>
        /// <returns>Autodiscover command response.</returns>
        AutodiscoverResponse Autodiscover(AutodiscoverRequest request, ContentTypeEnum contentType);

        /// <summary>
        /// Synchronizes changes in a collection between the client and the server.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <param name="isResyncNeeded">A bool value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>Sync command response</returns>
        SyncResponse Sync(SyncRequest request, bool isResyncNeeded = true);

        /// <summary>
        /// Sends MIME-formatted email messages to the server.
        /// </summary>
        /// <param name="request">A SendMailRequest object that contains the request information.</param>
        /// <returns>SendMail command response.</returns>
        SendMailResponse SendMail(SendMailRequest request);

        /// <summary>
        /// Retrieves an email attachment from the server.
        /// </summary>
        /// <param name="request">A GetAttachmentRequest object that contains the request information.</param>
        /// <returns>GetAttachment command response.</returns>
        GetAttachmentResponse GetAttachment(GetAttachmentRequest request);

        /// <summary>
        /// Synchronizes the collection hierarchy.
        /// </summary>
        /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>FolderSync command response.</returns>
        FolderSyncResponse FolderSync(FolderSyncRequest request);

        /// <summary>
        /// Creates a new folder as a child folder of the specified parent folder.
        /// </summary>
        /// <param name="request">A FolderCreateRequest object that contains the request information.</param>
        /// <returns>FolderCreate command response.</returns>
        FolderCreateResponse FolderCreate(FolderCreateRequest request);

        /// <summary>
        /// Deletes a folder from the server.
        /// </summary>
        /// <param name="request">A FolderDeleteRequest object that contains the request information.</param>
        /// <returns>FolderDelete command response.</returns>
        FolderDeleteResponse FolderDelete(FolderDeleteRequest request);

        /// <summary>
        /// Moves a folder from one location to another on the server or renames a folder.
        /// </summary>
        /// <param name="request">A FolderUpdateRequest object that contains the request information.</param>
        /// <returns>FolderUpdate command response.</returns>
        FolderUpdateResponse FolderUpdate(FolderUpdateRequest request);

        /// <summary>
        /// Moves an item or items from one folder to another on the server.
        /// </summary>
        /// <param name="request">A MoveItemsRequest object that contains the request information.</param>
        /// <returns>MoveItems command response.</returns>
        MoveItemsResponse MoveItems(MoveItemsRequest request);

        /// <summary>
        /// Gets the list of email folders from the server
        /// </summary>
        /// <returns>GetHierarchy command response.</returns>
        GetHierarchyResponse GetHierarchy();

        /// <summary>
        /// Gets an estimated number of items in a collection or folder that has to be synchronized on the server.
        /// </summary>
        /// <param name="request">A GetItemEstimateRequest object that contains the request information.</param>
        /// <returns>GetItemEstimate command response.</returns>
        GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest request);

        /// <summary>
        /// Accepts, tentatively accepts, or declines a meeting request in the user's Inbox folder or Calendar folder.
        /// </summary>
        /// <param name="request">A MeetingResponseRequest object that contains the request information.</param>
        /// <returns>MeetingResponse command response.</returns>
        MeetingResponseResponse MeetingResponse(MeetingResponseRequest request);

        /// <summary>
        /// Searches entries in an address book, mailbox, or document library.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>Search command response.</returns>
        SearchResponse Search(SearchRequest request);

        /// <summary>
        /// Finds entries in an address book, mailbox, or document library.
        /// </summary>
        /// <param name="request">A FindRequest object that contains the request information.</param>
        /// <returns>Find command response.</returns>
        FindResponse Find(FindRequest request);

        /// <summary>
        /// Supports get and set operations on global properties and Out of Office (OOF) settings for the user, sends device information to the server, implements the device password/personal identification number (PIN) recovery, or retrieves a list of the user's email addresses.
        /// </summary>
        /// <param name="request">A SettingsRequest object that contains the request information.</param>
        /// <returns>Settings command response.</returns>
        SettingsResponse Settings(SettingsRequest request);

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="request">A SmartForwardRequest object that contains the request information.</param>
        /// <returns>SmartForward command response.</returns>
        SmartForwardResponse SmartForward(SmartForwardRequest request);

        /// <summary>
        /// Replies to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="request">A SmartReplyRequest object that contains the request information.</param>
        /// <returns>SmartReply command response.</returns>
        SmartReplyResponse SmartReply(SmartReplyRequest request);

        /// <summary>
        /// Requests that the server monitor specified folders for changes that would require the client to resynchronize.
        /// </summary>
        /// <param name="request">A PingRequest object that contains the request information.</param>
        /// <returns>Ping command response.</returns>
        PingResponse Ping(PingRequest request);

        /// <summary>
        /// Acts as a container for the Fetch element, the EmptyFolderContents element, and the Move element to provide batched online handling of these operations against the server.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object that contains the request information.</param>
        /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
        /// <returns>ItemOperations command response.</returns>
        ItemOperationsResponse ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod);

        /// <summary>
        /// Enables client devices to request the administrator's security policy settings on the server.
        /// </summary>
        /// <param name="request">A ProvisionRequest object that contains the request information.</param>
        /// <returns>Provision command response.</returns>
        ProvisionResponse Provision(ProvisionRequest request);

        /// <summary>
        /// Resolves a list of supplied recipients, retrieves their free/busy information, or retrieves their S/MIME certificates so that clients can send encrypted S/MIME email messages.
        /// </summary>
        /// <param name="request">A ResolveRecipientsRequest object that contains the request information.</param>
        /// <returns>ResolveRecipients command response.</returns>
        ResolveRecipientsResponse ResolveRecipients(ResolveRecipientsRequest request);

        /// <summary>
        /// Validates a certificate that has been received via an S/MIME mail.
        /// </summary>
        /// <param name="request">A ValidateCertRequest object that contains the request information.</param>
        /// <returns>ValidateCert command response.</returns>
        ValidateCertResponse ValidateCert(ValidateCertRequest request);
        #endregion

        #region Helper Methods
        /// <summary>
        /// Sends a plain text request.
        /// </summary>
        /// <param name="cmdName">The name of the command to send.</param>
        /// <param name="parameters">The command parameters.</param>
        /// <param name="request">The plain text request.</param>
        /// <returns>The plain text response.</returns>
        SendStringResponse SendStringRequest(CommandName cmdName, IDictionary<CmdParameterName, object> parameters, string request);

        /// <summary>
        /// Changes http request header encoding type.
        /// </summary>
        /// <param name="headerEncodingType">The header encoding type.</param>
        void ChangeHeaderEncodingType(QueryValueType headerEncodingType);

        /// <summary>
        /// Changes device id.
        /// </summary>
        /// <param name="newDeviceId">The new device id.</param>
        void ChangeDeviceID(string newDeviceId);

        /// <summary>
        /// Changes the specified PolicyKey.
        /// </summary>
        /// <param name="appliedPolicyKey">The Policy Key to apply.</param>
        void ChangePolicyKey(string appliedPolicyKey);

        /// <summary>
        /// Changes device type.
        /// </summary>
        /// <param name="newDeviceType">The value of the new device type.</param>
        void ChangeDeviceType(string newDeviceType);

        /// <summary>
        /// Changes user to call ActiveSync operation.
        /// </summary>
        /// <param name="userName">The user's name.</param>
        /// <param name="userPassword">The user's password.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        void SwitchUser(string userName, string userPassword, string userDomain);
        #endregion
    }
}