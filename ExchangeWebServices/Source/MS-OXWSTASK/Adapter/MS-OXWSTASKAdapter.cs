namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System.Web.Services.Protocols;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSTASK.
    /// </summary>
    public partial class MS_OXWSTASKAdapter : ManagedAdapterBase, IMS_OXWSTASKAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSTASKAdapter Properties
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
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSTASK";

            // Execute the merge the configuration
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);
            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
            this.exchangeServiceBinding.SoapVersion = SoapProtocolVersion.Soap11;
        }

        #endregion

        #region IMS_OXWSTASKAdapter Operations
        /// <summary>
        /// Gets Task items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specifies a request to get Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            GetItemResponseType getItemResponse = this.exchangeServiceBinding.GetItem(getItemRequest);
            Site.Assert.IsNotNull(getItemResponse, "If the operation is successful, the response should not be null.");

            // Verify the get item operation.
            this.VerifyGetItemOperation(getItemResponse, this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return getItemResponse;
        }

        /// <summary>
        /// Copies Task items and puts the items in a different folder.
        /// </summary>
        /// <param name="copyItemRequest">Specifies a request to copy Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public CopyItemResponseType CopyItem(CopyItemType copyItemRequest)
        {
            CopyItemResponseType copyItemResponse = this.exchangeServiceBinding.CopyItem(copyItemRequest);
            Site.Assert.IsNotNull(copyItemResponse, "If the operation is successful, the response should not be null.");

            // Verify the copy item operation.
            this.VerifyCopyItemOperation(copyItemResponse, this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return copyItemResponse;
        }

        /// <summary>
        /// Creates Task items in the Exchange store
        /// </summary>
        /// <param name="createItemRequest">Specifies a request to create Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(createItemRequest);
            Site.Assert.IsNotNull(createItemResponse, "If the operation is successful, the response should not be null.");

            // Verify the create item operation.
            this.VerifyCreateItemOperation(createItemResponse, this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return createItemResponse;
        }

        /// <summary>
        /// Deletes Task items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specifies a request to delete Task item on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            DeleteItemResponseType deleteItemResponse = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);

            // Verify the delete item operation.
            this.VerifyDeleteItemOperation(this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return deleteItemResponse;
        }

        /// <summary>
        /// Moves Task items on the server.
        /// </summary>
        /// <param name="moveItemRequest">Specifies a request to move Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public MoveItemResponseType MoveItem(MoveItemType moveItemRequest)
        {
            MoveItemResponseType moveItemResponse = this.exchangeServiceBinding.MoveItem(moveItemRequest);
            Site.Assert.IsNotNull(moveItemResponse, "If the operation is successful, the response should not be null.");

            // Verify the move item operation.
            this.VerifyMoveItemOperation(moveItemResponse, this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return moveItemResponse;
        }

        /// <summary>
        /// Updates Task items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specifies a request to update Task items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest)
        {
            UpdateItemResponseType updateItemResponse = this.exchangeServiceBinding.UpdateItem(updateItemRequest);
            Site.Assert.IsNotNull(updateItemResponse, "If the operation is successful, the response should not be null.");

            // Verify the update item operation.
            this.VerifyUpdateItemOperation(updateItemResponse, this.exchangeServiceBinding.IsSchemaValidated);

            // Verify Soap version requirements.
            this.VerifySoapVersion();

            // Verify transport requirements.
            this.VerifyTransportType();

            return updateItemResponse;
        }

        #endregion
    }
}