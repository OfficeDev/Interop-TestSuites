namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSFOLD.
    /// </summary>
    public partial class MS_OXWSFOLDAdapter : ManagedAdapterBase, IMS_OXWSFOLDAdapter
    {
        #region Fields

        /// <summary>
        /// Exchange Web Service instance.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSFOLDAdapter Properties
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
            // Initialize.
            base.Initialize(testSite);
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            string userName = Common.GetConfigurationPropertyValue("OrganizerName", testSite);
            string password = Common.GetConfigurationPropertyValue("OrganizerPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string urlFormat = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            // Initialize service.
            this.exchangeServiceBinding = new ExchangeServiceBinding(urlFormat, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSFOLDAdapter Operations

        /// <summary>
        /// Create a new folder within a specific folder.
        /// </summary>
        /// <param name="request">Request of CreateFolder operation.</param>
        /// <returns>Response of CreateFolder operation.</returns>
        public CreateFolderResponseType CreateFolder(CreateFolderType request)
        {
            // Send the request and get the response.
            CreateFolderResponseType response = this.exchangeServiceBinding.CreateFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            return response;
        }

        /// <summary>
        /// Delete a folder from mailbox.
        /// </summary>
        /// <param name="request">Request DeleteFolder operation.</param>
        /// <returns>Response of DeleteFolder operation.</returns>
        public DeleteFolderResponseType DeleteFolder(DeleteFolderType request)
        {
            // Send the request and get the response.
            DeleteFolderResponseType response = this.exchangeServiceBinding.DeleteFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            return response;
        }

        /// <summary>
        /// Get folders, Calendar folders, Contacts folders, Tasks folders, and search folders.
        /// </summary>
        /// <param name="request">Request of GetFolder operation.</param>
        /// <returns>Response of GetFolder operation.</returns>
        public GetFolderResponseType GetFolder(GetFolderType request)
        {
            // Send the request and get the response.
            GetFolderResponseType response = this.exchangeServiceBinding.GetFolder(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            return response;
        }

        /// <summary>
        /// Switch the current user to the new one, with the identity of the new user to communicate with Exchange Web Service.
        /// </summary>
        /// <param name="userName">The name of a user</param>
        /// <param name="userPassword">The password of a user</param>
        /// <param name="userDomain">The domain, in which a user is</param>
        public void SwitchUser(string userName, string userPassword, string userDomain)
        {
            this.Initialize(this.Site);
            this.exchangeServiceBinding.Credentials = new System.Net.NetworkCredential(userName, userPassword, userDomain);
        }
        #endregion
    }
}