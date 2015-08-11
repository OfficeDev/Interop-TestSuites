namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the method defined in interface IMS_OXWSSRCHAdapter.
    /// </summary>
    public class MS_OXWSSRCHAdapter : ManagedAdapterBase, IMS_OXWSSRCHAdapter
    {
        #region Fields
        /// <summary>
        /// Exchange Web Service instance.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSSRCHAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.exchangeServiceBinding.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT.
        /// </summary>
        public IXPathNavigable LastRawResponseXml
        {
            get { return this.exchangeServiceBinding.LastRawResponseXml; }
        }

        #endregion

        #region Initialize TestSuite
        /// <summary>
        /// Overrides IAdapter's Initialize() and sets default protocol short name of the testSite.
        /// </summary>
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);

            // Merge the common configuration into local configuration
            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);
            Common.MergeGlobalConfig(commonConfigFileName, testSite);

            // Get the parameters from configuration files.
            string userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            string password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string urlFormat = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            // initialize service.
            this.exchangeServiceBinding = new ExchangeServiceBinding(urlFormat, userName, password, domain, testSite);

            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSSRCHAdapter Operation

        /// <summary>
        /// Find item on the server.
        /// </summary>
        /// <param name="findItemRequest">Find item operation request type.</param>
        /// <returns>Find item operation response type.</returns>
        public FindItemResponseType FindItem(FindItemType findItemRequest)
        {
            FindItemResponseType findItemResponse = this.exchangeServiceBinding.FindItem(findItemRequest);
            return findItemResponse;
        }

        #endregion
    }
}