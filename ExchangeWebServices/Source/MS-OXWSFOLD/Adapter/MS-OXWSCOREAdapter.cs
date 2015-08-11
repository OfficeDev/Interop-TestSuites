namespace Microsoft.Protocols.TestSuites.MS_OXWSFOLD
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the method defined in interface IMS_OXWSCOREAdapter.
    /// </summary>
    public class MS_OXWSCOREAdapter : ManagedAdapterBase, IMS_OXWSCOREAdapter
    {
        #region Fields
        /// <summary>
        /// Exchange Web Service instance.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSCOREAdapter Properties
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

        #region IMS_OXWSCOREAdapter Operations
        /// <summary>
        /// Creates item on the server.
        /// </summary>
        /// <param name="createItemRequest">Create item operation request type.</param>
        /// <returns>Create item operation response type.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(createItemRequest);
            return createItemResponse;
        }

        /// <summary>
        /// Update item on the server.
        /// </summary>
        /// <param name="updateItemRequest">Update item operation request type.</param>
        /// <returns>Update item operation response type.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest)
        {
            UpdateItemResponseType updateItemResponse = this.exchangeServiceBinding.UpdateItem(updateItemRequest);
            return updateItemResponse;
        }

        /// <summary>
        /// Get item on the server.
        /// </summary>
        /// <param name="getItemRequest">Get item operation request type.</param>
        /// <returns>Get item operation response type.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            GetItemResponseType gitItemResponse = this.exchangeServiceBinding.GetItem(getItemRequest);
            return gitItemResponse;
        }

        /// <summary>
        /// Delete item on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Delete item operation request type.</param>
        /// <returns>Delete item operation response type.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            DeleteItemResponseType deleteItemResponse = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);
            return deleteItemResponse;
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