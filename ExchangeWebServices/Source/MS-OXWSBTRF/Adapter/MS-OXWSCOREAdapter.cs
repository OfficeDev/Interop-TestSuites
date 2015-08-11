namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the methods defined in interface IMS_OXWSCOREAdapter.
    /// </summary>
    public class MS_OXWSCOREAdapter : ManagedAdapterBase, IMS_OXWSCOREAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary> 
        private ExchangeServiceBinding exchangeServiceBinding;

        /// <summary>
        /// The user name used to access web service.
        /// </summary>
        private string userName;

        /// <summary>
        /// The password for userName used to access web service.
        /// </summary>
        private string password;

        /// <summary>
        /// The domain of server.
        /// </summary>
        private string domain;

        /// <summary>
        /// The endpoint url of Exchange Web Service.
        /// </summary>
        private string url;

        #endregion

        #region IMS_OXWSCOREAdapter Properties
        /// <summary>
        /// Gets the raw XML request sent to protocol SUT
        /// </summary>
        public IXPathNavigable LastRawRequestXml
        {
            get { return this.exchangeServiceBinding.LastRawRequestXml; }
        }

        /// <summary>
        /// Gets the raw XML response received from protocol SUT
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
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSBTRF";

            // Merge common configuration and SHOULD/MAY configuration files
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            this.password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSCOREAdapter Operations
        /// <summary>
        /// Creates items on the server.
        /// </summary>
        /// <param name="createItemRequest">Specify the request for CreateItem operation.</param>
        /// <returns>The response to this operation request.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            if (createItemRequest == null)
            {
                throw new ArgumentException("The CreateItem request should not be null.");
            }

            // Send the request and get the response.
            CreateItemResponseType response = this.exchangeServiceBinding.CreateItem(createItemRequest);
            return response;
        }

        /// <summary>
        /// Gets items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify the request for GetItem operation.</param>
        /// <returns>The response to this operation request.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            if (getItemRequest == null)
            {
                throw new ArgumentException("The GetItem request should not be null.");
            }

            // Send the request and get the response.
            GetItemResponseType response = this.exchangeServiceBinding.GetItem(getItemRequest);
            return response;
        }
        #endregion
    }
}