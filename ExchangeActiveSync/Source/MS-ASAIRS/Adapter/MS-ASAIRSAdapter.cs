namespace Microsoft.Protocols.TestSuites.MS_ASAIRS
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASAIRS.
    /// </summary>
    public partial class MS_ASAIRSAdapter : ManagedAdapterBase, IMS_ASAIRSAdapter
    {
        #region Variables
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;
        #endregion

        #region IMS_ASAIRSAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASAIRS";

            // Merge the common configuration
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("User1Name", testSite),
                Password = Common.GetConfigurationPropertyValue("User1Password", testSite)
            };
        }

        /// <summary>
        /// Changes user to call ActiveSync operation.
        /// </summary>
        /// <param name="userName">The name of a user.</param>
        /// <param name="userPassword">The password of a user.</param>
        /// <param name="userDomain">The domain which the user belongs to.</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.activeSyncClient.UserName = userName;
            this.activeSyncClient.Password = userPassword;
            this.activeSyncClient.Domain = userDomain;
        }

        #region IMS_ASAIRSAdapter Commands
        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending SyncRequest object.
        /// </summary>
        /// <param name="request">A SyncRequest object that contains the request information.</param>
        /// <returns>A SyncStore object.</returns>
        public SyncStore Sync(SyncRequest request)
        {
            SyncResponse response = this.activeSyncClient.Sync(request, true);
            Site.Assert.IsNotNull(response, "If the Sync command executes successfully, the response from server should not be null.");
            SyncStore syncStore = Common.LoadSyncResponse(response);
            this.VerifySyncResponse(response, syncStore);
            this.VerifyWBXMLCapture();

            return syncStore;
        }

        /// <summary>
        /// Synchronizes the changes in a collection between the client and the server by sending raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw Sync request.</param>
        /// <returns>A SendStringResponse object.</returns>
        public SendStringResponse Sync(string request)
        {
            return this.activeSyncClient.SendStringRequest(CommandName.Sync, null, request);
        }

        /// <summary>
        /// Retrieves an item from the server by sending ItemOperationsRequest object.
        /// </summary>
        /// <param name="request">An ItemOperationsRequest object which contains the request information.</param>
        /// <param name="deliveryMethod">Delivery method specifies what kind of response is accepted.</param>
        /// <returns>An ItemOperationsStore object.</returns>
        public ItemOperationsStore ItemOperations(ItemOperationsRequest request, DeliveryMethodForFetch deliveryMethod)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(request, deliveryMethod);
            Site.Assert.IsNotNull(response, "If the ItemOperations command executes successfully, the response from server should not be null.");

            ItemOperationsStore itemOperationsStore = Common.LoadItemOperationsResponse(response);
            this.VerifyItemOperationsResponse(response, itemOperationsStore);
            this.VerifyWBXMLCapture();

            return itemOperationsStore;
        }

        /// <summary>
        /// Retrieves an item from the server by sending the raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw ItemOperations request.</param>
        /// <returns>A SendStringResponse object.</returns>
        public SendStringResponse ItemOperations(string request)
        {
            return this.activeSyncClient.SendStringRequest(CommandName.ItemOperations, null, request);
        }

        /// <summary>
        /// Finds entries in an address book, mailbox or document library by sending SearchRequest object.
        /// </summary>
        /// <param name="request">A SearchRequest object that contains the request information.</param>
        /// <returns>A SearchStore object.</returns>
        public SearchStore Search(SearchRequest request)
        {
            SearchResponse response = this.activeSyncClient.Search(request, true);
            Site.Assert.IsNotNull(response, "If the Search command executes successfully, the response from server should not be null.");
            SearchStore searchStore = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            this.VerifySearchResponse(response, searchStore);
            this.VerifyWBXMLCapture();

            return searchStore;
        }

        /// <summary>
        /// Finds entries in an address book, mailbox or document library by sending raw string request.
        /// </summary>
        /// <param name="request">A string which contains the raw Search request.</param>
        /// <returns>A SendStringResponse object.</returns>
        public SendStringResponse Search(string request)
        {
            return this.activeSyncClient.SendStringRequest(CommandName.Search, null, request);
        }

        /// <summary>
        /// Synchronizes the collection hierarchy.
        /// </summary>
        /// <param name="request">A FolderSyncRequest object that contains the request information.</param>
        /// <returns>A FolderSyncResponse object.</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest request)
        {
            FolderSyncResponse response = this.activeSyncClient.FolderSync(request);
            Site.Assert.IsNotNull(response, "If the FolderSync command executes successfully, the response from server should not be null.");
            return response;
        }

        /// <summary>
        /// Sends MIME-formatted e-mail messages to the server.
        /// </summary>
        /// <param name="request">A SendMailRequest object that contains the request information.</param>
        /// <returns>A SendMailResponse object.</returns>
        public SendMailResponse SendMail(SendMailRequest request)
        {
            SendMailResponse response = this.activeSyncClient.SendMail(request);
            Site.Assert.IsNotNull(response, "If the SendMail command executes successfully, the response from server should not be null.");
            return response;
        }

        /// <summary>
        /// Accept, tentatively accept, or decline a meeting request in the user's Inbox folder or Calendar folder
        /// </summary>
        /// <param name="request">A MeetingResponseRequest object that contains the request information.</param>
        /// <returns>A MeetingResponseResponse object.</returns>
        public MeetingResponseResponse MeetingResponse(MeetingResponseRequest request)
        {
            MeetingResponseResponse response = this.activeSyncClient.MeetingResponse(request);
            Site.Assert.IsNotNull(response, "If the MeetingResponse command executes successfully, the response from server should not be null.");
            return response;
        }

        /// <summary>
        /// Forward messages without retrieving the full, original message from the server.
        /// </summary>
        /// <param name="request">A SmartForwardRequest object that contains the request information.</param>
        /// <returns>A SmartForwardResponse object.</returns>
        public SmartForwardResponse SmartForward(SmartForwardRequest request)
        {
            SmartForwardResponse response = this.activeSyncClient.SmartForward(request);
            Site.Assert.IsNotNull(response, "If the SmartForward command executes successfully, the response from server should not be null.");
            return response;
        }
        #endregion
    }
}