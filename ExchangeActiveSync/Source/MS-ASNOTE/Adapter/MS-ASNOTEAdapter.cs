namespace Microsoft.Protocols.TestSuites.MS_ASNOTE
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASNOTE.
    /// </summary>
    public partial class MS_ASNOTEAdapter : ManagedAdapterBase, IMS_ASNOTEAdapter
    {
        #region private field
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;

        #endregion

        #region IMS_ASNOTEAdapter Properties
        /// <summary>
        /// Gets the XML request sent to protocol SUT
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.activeSyncClient.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the XML response received from protocol SUT
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.activeSyncClient.LastRawResponseXml; }
        }
        #endregion

        /// <summary>
        /// Sync data from the server
        /// </summary>
        /// <param name="syncRequest">Sync command request.</param>
        /// <param name="isResyncNeeded">A bool value indicates whether need to re-sync when the response contains MoreAvailable element.</param>
        /// <returns>The sync result which is returned from server</returns>
        public SyncStore Sync(SyncRequest syncRequest, bool isResyncNeeded)
        {
            SyncResponse response = this.activeSyncClient.Sync(syncRequest, isResyncNeeded);
            this.VerifySyncResponse(response);
            SyncStore result = Common.LoadSyncResponse(response);
            this.VerifyTransport();
            this.VerifySyncResult(result);
            this.VerifyWBXMLCapture();
            return result;
        }

        /// <summary>
        /// Synchronizes the collection hierarchy 
        /// </summary>
        /// <param name="folderSyncRequest">FolderSync command request.</param>
        /// <returns>The FolderSync response which is returned from the server</returns>
        public FolderSyncResponse FolderSync(FolderSyncRequest folderSyncRequest)
        {
            FolderSyncResponse response = this.activeSyncClient.FolderSync(folderSyncRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");
            return response;
        }

        /// <summary>
        /// Loop to get the results of the specific query request by Search command.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to search.</param>
        /// <param name="subject">The subject of the note to get.</param>
        /// <param name="isLoopNeeded">A boolean value specify whether need the loop</param>
        /// <param name="expectedCount">The expected number of the note to be found.</param>
        /// <returns>The results in response of Search command</returns>
        public SearchStore Search(string collectionId, string subject, bool isLoopNeeded, int expectedCount)
        {
            SearchRequest searchRequest = Common.CreateSearchRequest(subject, collectionId);
            SearchResponse response = this.activeSyncClient.Search(searchRequest, isLoopNeeded, expectedCount);
            this.VerifySearchResponse(response);
            SearchStore result = Common.LoadSearchResponse(response, Common.GetConfigurationPropertyValue("ActiveSyncProtocolVersion", this.Site));
            this.VerifyTransport();
            this.VerifySearchResult(result);
            this.VerifyWBXMLCapture();
            return result;
        }

        /// <summary>
        /// Fetch all information about exchange object
        /// </summary>
        /// <param name="itemOperationsRequest">ItemOperations command request.</param>
        /// <returns>The ItemOperations result which is returned from server</returns>
        public ItemOperationsStore ItemOperations(ItemOperationsRequest itemOperationsRequest)
        {
            ItemOperationsResponse response = this.activeSyncClient.ItemOperations(itemOperationsRequest, DeliveryMethodForFetch.Inline);
            this.VerifyItemOperationsResponse(response);
            ItemOperationsStore result = Common.LoadItemOperationsResponse(response);
            bool hasSchemaElement = AdapterHelper.ContainsSchemaElement(itemOperationsRequest);
            this.VerifyTransport();
            this.VerifyItemOperationResult(result, hasSchemaElement);
            this.VerifyWBXMLCapture();
            return result;
        }

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-ASNOTE";

            // Merge the common configuration
            Common.MergeConfiguration(testSite);
            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("UserName", testSite),
                Password = Common.GetConfigurationPropertyValue("UserPassword", testSite)
            };
        }
    }
}