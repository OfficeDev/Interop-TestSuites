namespace Microsoft.Protocols.TestSuites.MS_ASCMD
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Threading;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASCMD.
    /// </summary>
    public partial class MS_ASCMDAdapter : ManagedAdapterBase, IMS_ASCMDAdapter
    {
        #region Variables
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;
        #endregion

        #region IMS_ASCMDAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-ASCMD";

            // Merge the common configuration
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                AcceptLanguage = "en-us",
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }
        #endregion

        #region IMS-ASCMDAdapter Members
        /// <summary>
        /// Facilitates the discovery of core account configuration information by using the user's Simple Mail Transfer Protocol (SMTP) address as the primary input
        /// </summary>
        /// <param name="request">An AutodiscoverRequest object that contains the request information.</param>
        /// <param name="contentType">Content Type that indicates the body's format</param>
        /// <returns>Autodiscover command response</returns>
        public AutodiscoverResponse Autodiscover(AutodiscoverRequest request, ContentTypeEnum contentType)
        {
            AutodiscoverResponse response = this.activeSyncClient.Autodiscover(request, contentType);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Autodiscover, response);
            this.VerifyAutodiscoverCommand(response);
            return response;
        }

        /// <summary>
        /// Synchronizes changes in a collection between the client and the server.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <param name="isResyncNeeded">A bool value indicate whether need to re-sync when the response contains MoreAvailable.</param>
        /// <returns>Sync command response</returns>
        public SyncResponse Sync(SyncRequest request, bool isResyncNeeded = true)
        {
            SyncResponse response = this.activeSyncClient.Sync(request, isResyncNeeded);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Sync, response);
            this.VerifySyncCommand(response);
            return response;
        }

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="request">A SendMailRequest object that contains the request information.</param>
        /// <returns>SendMail command response</returns>
        public SendMailResponse SendMail(SendMailRequest request)
        {
            SendMailResponse response = this.activeSyncClient.SendMail(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.SendMail, response);
            this.VerifySendMailCommand(response);
            return response;
        }

        /// <summary>
        /// Retrieves an e-mail attachment from the server.
        /// </summary>
        /// <param name="request">A GetAttachmentRequest object that contains the request information.</param>
        /// <returns>GetAttachment command response</returns>
        public GetAttachmentResponse GetAttachment(GetAttachmentRequest request)
        {
            GetAttachmentResponse response = this.activeSyncClient.GetAttachment(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.GetAttachment, response);
            return response;
        }

        /// <summary>
        /// Synchronizes the collection hierarchy 
        /// </summary>
        /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>FolderSync command response</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest request)
        {
            FolderSyncResponse response = this.activeSyncClient.FolderSync(request);
            this.VerifyTransportRequirements();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                this.VerifyWBXMLCapture(CommandName.FolderSync, response);
                this.VerifyFolderSyncCommand(response);
            }

            return response;
        }

        /// <summary>
        /// Creates a new folder as a child folder of the specified parent folder. 
        /// </summary>
        /// <param name="request">A FolderCreateRequest object that contains the request information.</param>
        /// <returns>FolderCreate command response</returns>
        public FolderCreateResponse FolderCreate(FolderCreateRequest request)
        {
            FolderCreateResponse response = this.activeSyncClient.FolderCreate(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.FolderCreate, response);
            this.VerifyFolderCreateCommand(response);
            return response;
        }

        /// <summary>
        /// Deletes a folder from the server.
        /// </summary>
        /// <param name="request">A FolderDeleteRequest object that contains the request information.</param>
        /// <returns>FolderDelete command response</returns>
        public FolderDeleteResponse FolderDelete(FolderDeleteRequest request)
        {
            FolderDeleteResponse response = this.activeSyncClient.FolderDelete(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.FolderDelete, response);
            this.VerifyFolderDeleteCommand(response);
            return response;
        }

        /// <summary>
        /// Moves a folder from one location to another on the server or renames a folder.
        /// </summary>
        /// <param name="request">A FolderUpdateRequest object that contains the request information.</param>
        /// <returns>FolderUpdate command response</returns>
        public FolderUpdateResponse FolderUpdate(FolderUpdateRequest request)
        {
            FolderUpdateResponse response = this.activeSyncClient.FolderUpdate(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.FolderUpdate, response);
            this.VerifyFolderUpdateCommand(response);
            return response;
        }

        /// <summary>
        /// Moves an item or items from one folder to another on the server..
        /// </summary>
        /// <param name="request">A MoveItemsRequest object that contains the request information.</param>
        /// <returns>MoveItems command response</returns>
        public MoveItemsResponse MoveItems(MoveItemsRequest request)
        {
            MoveItemsResponse response = this.activeSyncClient.MoveItems(request);
            this.VerifyTransportRequirements();
            if (response.ResponseData.Response != null)
            {
                this.VerifyWBXMLCapture(CommandName.MoveItems, response);
                this.VerifyMoveItemsCommand(response);
            }

            return response;
        }

        /// <summary>
        /// Gets the list of email folders from the server
        /// </summary>
        /// <returns>GetHierarchy command response.</returns>
        public GetHierarchyResponse GetHierarchy()
        {
            GetHierarchyResponse response = this.activeSyncClient.GetHierarchy();
            this.VerifyGetHierarchyCommand(response);
            return response;
        }

        /// <summary>
        /// Gets an estimated number of items in a collection or folder on the server that has to be synchronized.
        /// </summary>
        /// <param name="request">A GetItemEstimateRequest object that contains the request information.</param>
        /// <returns>GetItemEstimate command response</returns>
        public GetItemEstimateResponse GetItemEstimate(GetItemEstimateRequest request)
        {
            GetItemEstimateResponse response = this.activeSyncClient.GetItemEstimate(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.GetItemEstimate, response);
            this.VerifyGetItemEstimateCommand(response);
            return response;
        }

        /// <summary>
        /// Accepts, tentatively accepts, or declines a meeting request in the user's Inbox folder or Calendar folder.
        /// </summary>
        /// <param name="request">A MeetingResponseRequest object that contains the request information.</param>
        /// <returns>MeetingResponse command response</returns>
        public MeetingResponseResponse MeetingResponse(MeetingResponseRequest request)
        {
            MeetingResponseResponse response = this.activeSyncClient.MeetingResponse(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.MeetingResponse, response);
            this.VerifyMeetingResponseCommand(response);
            return response;
        }

        /// <summary>
        /// Finds entries in an address book, mailbox, or document library.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>Search command response.</returns>
        public SearchResponse Search(SearchRequest request)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            SearchResponse response = this.activeSyncClient.Search(request);

            while (counter < retryCount && response.ResponseData.Status.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.activeSyncClient.Search(request);
                counter++;
            }

            Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Search, response);
            this.VerifySearchCommand(response);
            return response;
        }

        /// <summary>
        /// Finds entries in an address book, mailbox, or document library.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>Search command response.</returns>
        public FindResponse Find(FindRequest request)
        {
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int retryCount = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            FindResponse response = this.activeSyncClient.Find(request);

            while (counter < retryCount && response.ResponseData.Status.Equals("10"))
            {
                Thread.Sleep(waitTime);
                response = this.activeSyncClient.Find(request);
                counter++;
            }
            
            Site.Log.Add(LogEntryKind.Debug, "Loop {0} times to get the search item", counter);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Find, response);
            this.VerifyFindCommand(response);
            return response;
        }

        /// <summary>
        /// Supports get and set operations on global properties and Out of Office (OOF) settings for the user, sends device information to the server, implements the device password/personal identification number (PIN) recovery, and retrieves a list of the user's e-mail addresses.
        /// </summary>
        /// <param name="request">A SettingsRequest object that contains the request information.</param>
        /// <returns>Settings command response</returns>
        public SettingsResponse Settings(SettingsRequest request)
        {
            SettingsResponse response = this.activeSyncClient.Settings(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Settings, response);
            this.VerifySettingsCommand(response);
            return response;
        }

        /// <summary>
        /// Forwards messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="request">A SmartForwardRequest object that contains the request information.</param>
        /// <returns>SmartForward command response</returns>
        public SmartForwardResponse SmartForward(SmartForwardRequest request)
        {
            SmartForwardResponse response = this.activeSyncClient.SmartForward(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.SmartForward, response);
            this.VerifySmartForwardCommand(response);
            return response;
        }

        /// <summary>
        /// Reply to messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="request">A SmartReplyRequest object that contains the request information.</param>
        /// <returns>SmartReply command response</returns>
        public SmartReplyResponse SmartReply(SmartReplyRequest request)
        {
            SmartReplyResponse response = this.activeSyncClient.SmartReply(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.SmartReply, response);
            this.VerifySmartReplyCommand(response);
            return response;
        }

        /// <summary>
        /// Requests that the server monitor specified folders for changes that would require the client to resynchronize.
        /// </summary>
        /// <param name="request">A PingRequest object that contains the request information.</param>
        /// <returns>Ping command response</returns>
        public PingResponse Ping(PingRequest request)
        {
            PingResponse response = this.activeSyncClient.Ping(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.Ping, response);
            this.VerifyPingCommand(response);
            return response;
        }

        /// <summary>
        /// Acts as a container for the Fetch element, the EmptyFolderContents element, and the Move element to provide batched online handling of these operations against the server.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object that contains the request information.</param>
        /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
        /// <returns>ItemOperations command response</returns>
        public ItemOperationsResponse ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(request, deliveryMethod);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.ItemOperations, response);
            this.VerifyItemOperationsCommand(response);
            return response;
        }

        /// <summary>
        /// Enables client devices to request the administrator's security policy settings on the server..
        /// </summary>
        /// <param name="request">A ProvisionRequest object that contains the request information.</param>
        /// <returns>Provision command response</returns>
        public ProvisionResponse Provision(ProvisionRequest request)
        {
            // When the value of the MS-ASProtocolVersion header is 14.0 or 12.1, the client MUST NOT send the setting:DeviceInformation element in any Provision command request.
            if (this.activeSyncClient.ActiveSyncProtocolVersion.Equals("140", StringComparison.OrdinalIgnoreCase) || this.activeSyncClient.ActiveSyncProtocolVersion.Equals("121", StringComparison.OrdinalIgnoreCase))
            {
                request.RequestData.DeviceInformation = null;
            }

            ProvisionResponse response = this.activeSyncClient.Provision(request);
            this.VerifyTransportRequirements();
            this.VerifyProvisionCommand(response);
            this.VerifyWBXMLCapture(CommandName.Provision, response);
            return response;
        }

        /// <summary>
        /// Resolves a list of supplied recipients, retrieves their free/busy information, or retrieves their S/MIME certificates so that clients can send encrypted S/MIME e-mail messages.
        /// </summary>
        /// <param name="request">A ResolveRecipientsRequest object that contains the request information.</param>
        /// <returns>ResolveRecipients command response</returns>
        public ResolveRecipientsResponse ResolveRecipients(ResolveRecipientsRequest request)
        {
            ResolveRecipientsResponse response = this.activeSyncClient.ResolveRecipients(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.ResolveRecipients, response);
            this.VerifyResolveRecipientsCommand(response);
            return response;
        }

        /// <summary>
        /// Validates a certificate that has been received via an S/MIME mail.
        /// </summary>
        /// <param name="request">A ValidateCertRequest object that contains the request information.</param>
        /// <returns>ValidateCert command response</returns>
        public ValidateCertResponse ValidateCert(ValidateCertRequest request)
        {
            ValidateCertResponse response = this.activeSyncClient.ValidateCert(request);
            this.VerifyTransportRequirements();
            this.VerifyWBXMLCapture(CommandName.ValidateCert, response);
            this.VerifyValidateCertCommand(response);
            return response;
        }

        /// <summary>
        /// Sends a plain text request.
        /// </summary>
        /// <param name="cmdName">The name of the command to send</param>
        /// <param name="parameters">The command parameters</param>
        /// <param name="request">The plain text request</param>
        /// <returns>The plain text response.</returns>
        public SendStringResponse SendStringRequest(CommandName cmdName, IDictionary<CmdParameterName, object> parameters, string request)
        {
            SendStringResponse response = this.activeSyncClient.SendStringRequest(cmdName, parameters, request);
            return response;
        }

        /// <summary>
        /// Changes device id.
        /// </summary>
        /// <param name="newDeviceId">The new device id.</param>
        public void ChangeDeviceID(string newDeviceId)
        {
            this.activeSyncClient.DeviceID = newDeviceId;
        }

        /// <summary>
        /// Changes the specified PolicyKey.
        /// </summary>
        /// <param name="appliedPolicyKey">The Policy Key to apply.</param>
        public void ChangePolicyKey(string appliedPolicyKey)
        {
            this.activeSyncClient.PolicyKey = appliedPolicyKey;
        }

        /// <summary>
        /// Changes http request header encoding type
        /// </summary>
        /// <param name="headerEncodingType">The header encoding type</param>
        public void ChangeHeaderEncodingType(QueryValueType headerEncodingType)
        {
            this.activeSyncClient.QueryValueType = headerEncodingType;
        }

        /// <summary>
        /// Changes device type.
        /// </summary>
        /// <param name="newDeviceType">The value of the new device type.</param>
        public void ChangeDeviceType(string newDeviceType)
        {
            this.activeSyncClient.DeviceType = newDeviceType;
        }

        /// <summary>
        /// Changes user to call ActiveSync operation.
        /// </summary>
        /// <param name="userName">The user's name.</param>
        /// <param name="userPassword">The user's password.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.activeSyncClient.UserName = userName;
            this.activeSyncClient.Password = userPassword;
            this.activeSyncClient.Domain = userDomain;
        }
        #endregion
    }
}