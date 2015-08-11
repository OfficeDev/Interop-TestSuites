namespace Microsoft.Protocols.TestSuites.MS_ASDOC
{
    using System.Net;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-ASDOC. 
    /// </summary>
    public partial class MS_ASDOCAdapter : ManagedAdapterBase, IMS_ASDOCAdapter
    {
        /// <summary>
        /// The instance of ActiveSync client.
        /// </summary>
        private ActiveSyncClient activeSyncClient;

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

        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Transfer ITestSite into adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Site.DefaultProtocolDocShortName = "MS-ASDOC";

            // Get the name of common configuration file.
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);

            // Merge the common configuration
            Common.MergeGlobalConfig(commonConfigFileName, testSite);
            this.activeSyncClient = new ActiveSyncClient(testSite)
            {
                UserName = Common.GetConfigurationPropertyValue("UserName", testSite),
                Password = Common.GetConfigurationPropertyValue("UserPassword", testSite)
            };
        }

        /// <summary>
        /// Retrieves data from the server for one or more individual documents.
        /// </summary>
        /// <param name="itemOperationsRequest">ItemOperations command request.</param>
        /// <param name="deliverMethod">Deliver method parameter.</param>
        /// <returns>ItemOperations command response.</returns>
        public ItemOperationsResponse ItemOperations(ItemOperationsRequest itemOperationsRequest, DeliveryMethodForFetch deliverMethod)
        {
            ItemOperationsResponse itemOperationsResponse = this.activeSyncClient.ItemOperations(itemOperationsRequest, deliverMethod);
            Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, itemOperationsResponse.StatusCode, "The call should be successful.");
            this.VerifyTransport();
            this.VerifyItemOperations(itemOperationsResponse);
            this.VerifyWBXMLCapture();
            return itemOperationsResponse;
        }

        /// <summary>
        /// Finds entries in document library (using Universal Naming Convention paths).
        /// </summary>
        /// <param name="searchRequest">Search command request.</param>
        /// <returns>Search command response.</returns>
        public SearchResponse Search(SearchRequest searchRequest)
        {
            SearchResponse searchResponse = this.activeSyncClient.Search(searchRequest);
            Site.Assert.AreEqual<HttpStatusCode>(HttpStatusCode.OK, searchResponse.StatusCode, "The call should be successful.");
            this.VerifyTransport();
            this.VerifySearch(searchResponse);
            this.VerifyWBXMLCapture();
            return searchResponse;
        }
    }
}