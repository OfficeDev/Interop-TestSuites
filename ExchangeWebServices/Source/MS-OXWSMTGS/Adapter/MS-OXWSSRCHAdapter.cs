namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the methods defined in interface IMS_OXWSSRCHAdapter. 
    /// </summary>
    public class MS_OXWSSRCHAdapter : ManagedAdapterBase, IMS_OXWSSRCHAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        /// <summary>
        /// The user name used to access web service.
        /// </summary>
        private string username;

        /// <summary>
        /// The password for username used to access web service.
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

        #region IMS_OXWSSRCHAdapter Properties
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
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            Common.MergeConfiguration(testSite);

            this.username = Common.GetConfigurationPropertyValue("OrganizerName", this.Site);
            this.password = Common.GetConfigurationPropertyValue("OrganizerPassword", this.Site);
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.username, this.password, this.domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }
        #endregion

        #region IMS_OXWSSRCHAdapter Operations
        /// <summary>
        /// Search specified items.
        /// </summary>
        /// <param name="findRequest">A request to the FindItem operation.</param>
        /// <returns>The response message returned by FindItem operation.</returns>
        public FindItemResponseType FindItem(FindItemType findRequest)
        {
            FindItemResponseType findResponse = this.exchangeServiceBinding.FindItem(findRequest);

            return findResponse;
        }

        /// <summary>
        /// Switch the current user to the new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user</param>
        /// <param name="userPassword">The password of a user</param>
        /// <param name="userDomain">The domain, in which a user is</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.username = userName;
            this.password = userPassword;
            this.domain = userDomain;

            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(this.username, this.password, this.domain);
        }
        #endregion
    }
}