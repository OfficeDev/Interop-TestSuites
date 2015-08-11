namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
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
        /// An instance of ExchangeServiceBinding
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

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
        /// Initialize some variables overridden.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite Class.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);

            string userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            string password = Common.GetConfigurationPropertyValue("User1Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }

        #endregion

        #region IMS_OXWSSRCHAdapter Operations
        /// <summary>
        /// Search specific folders.
        /// </summary>
        /// <param name="findRequest">Specify a request for a FindFolder operation</param>
        /// <returns>A response to FindFolder operation request</returns>
        public FindFolderResponseType FindFolder(FindFolderType findRequest)
        {
            FindFolderResponseType findResponse = this.exchangeServiceBinding.FindFolder(findRequest);

            return findResponse; 
        }
        
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
            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, userPassword, userDomain);
        }
        #endregion
    }
}