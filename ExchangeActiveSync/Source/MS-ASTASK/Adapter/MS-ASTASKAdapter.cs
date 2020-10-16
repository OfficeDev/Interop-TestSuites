namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;

    /// <summary>
    /// Adapter class of MS-ASTASK.
    /// </summary>
    public partial class MS_ASTASKAdapter : ManagedAdapterBase, IMS_ASTASKAdapter
    {
        #region Private field

        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;

        #endregion

        #region IMS_ASTASKAdapter Properties

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

        #region IMS_ASTASKAdapter Initialize method

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-ASTASK";

            // Merge configuration.
            Common.MergeConfiguration(testSite);

            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("UserName", this.Site),
                Password = Common.GetConfigurationPropertyValue("Password", this.Site)
            };
        }

        #endregion

        #region IMS_ASTASKAdapter Commands

        /// <summary>
        /// Sync data from the server.
        /// </summary>
        /// <param name="syncRequest">A Sync command request.</param>
        /// <returns>A Sync command response returned from the server.</returns>
        public SyncStore Sync(SyncRequest syncRequest)
        {
            SyncResponse response = this.activeSyncClient.Sync(syncRequest, true);
            Site.Assert.IsNotNull(response, "The Sync response should be returned.");
            this.VerifyTransport();
            this.VerifyWBXMLRequirements();

            SyncStore syncResponse = Common.LoadSyncResponse(response);

            foreach (Request.SyncCollection collection in syncRequest.RequestData.Collections)
            {
                if (collection.SyncKey != "0")
                {
                    this.VerifyMessageSyntax();
                    this.VerifySyncCommandResponse(syncResponse);
                }
            }

            return syncResponse;
        }

        /// <summary>
        /// Synchronize the collection hierarchy.
        /// </summary>
        /// <returns>A FolderSync command response returned from the server.</returns>
        public FolderSyncResponse FolderSync()
        {
            FolderSyncRequest request = Common.CreateFolderSyncRequest("0");
            FolderSyncResponse folderSyncResponse = this.activeSyncClient.FolderSync(request);
            Site.Assert.IsNotNull(folderSyncResponse, "The FolderSync response should be returned.");

            return folderSyncResponse;
        }

        /// <summary>
        /// Search data using the given keyword text.
        /// </summary>
        /// <param name="searchRequest">A Search command request.</param>
        /// <returns>A Search command response returned from the server.</returns>
        public SearchStore Search(SearchRequest searchRequest)
        {
            SearchResponse response = this.activeSyncClient.Search(searchRequest, true);
            Site.Assert.IsNotNull(response, "The Search response should be returned.");
            SearchStore searchResponse = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            this.VerifySearchCommandResponse(searchResponse);

            return searchResponse;
        }

        /// <summary>
        /// Fetch all the data about tasks.
        /// </summary>
        /// <param name="itemOperationsRequest">An ItemOperations command request.</param>
        /// <returns>An ItemOperations command response returned from the server.</returns>
        public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            Site.Assert.IsNotNull(response, "The ItemOperations response should be returned.");
            ItemOperationsStore itemOperationResponse = Common.LoadItemOperationsResponse(response);
            this.VerifyItemOperationsResponse(itemOperationResponse);

            return itemOperationResponse;
        }

        /// <summary>
        /// Send a string request and get a response from server.
        /// </summary>
        /// <param name="stringRequest">A string request for a certain command.</param>
        /// <param name="commandName">Commands choices.</param>
        /// <returns>A string response returned from the server.</returns>
        public SendStringResponse SendStringRequest(string stringRequest, CommandName commandName)
        {
            SendStringResponse response = this.activeSyncClient.SendStringRequest(commandName, null, stringRequest);
            Site.Assert.IsNotNull(response, "The string response should be returned.");

            return response;
        }

        #endregion
    }
}