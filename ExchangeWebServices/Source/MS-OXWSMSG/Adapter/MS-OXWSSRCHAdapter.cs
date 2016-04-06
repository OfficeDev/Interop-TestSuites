namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Net;
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
            string userName = Common.GetConfigurationPropertyValue("Sender", testSite);
            string password = Common.GetConfigurationPropertyValue("SenderPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string urlFormat = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            // initialize service.
            this.exchangeServiceBinding = new ExchangeServiceBinding(urlFormat, userName, password, domain, testSite);

            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSSRCHAdapter Operation
        /// <summary>
        /// Switch the current user to the new user, with the identity of the new role to communicate with server.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        public void SwitchUser(string userName, string password, string domain)
        {
            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, password, domain);

            // Verify the credential of the exchange service binding.
            bool isVerified = false;
            Uri uri = new Uri(Common.GetConfigurationPropertyValue("ServiceUrl", this.Site));
            NetworkCredential credential = this.exchangeServiceBinding.Credentials.GetCredential(uri, "basic");
            if (credential.Domain == domain && credential.UserName == userName)
            {
                isVerified = true;
            }

            this.Site.Assert.IsTrue(isVerified, "Service binding should be successful");
        }

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