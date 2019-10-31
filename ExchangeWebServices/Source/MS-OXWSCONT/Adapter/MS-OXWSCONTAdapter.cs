namespace Microsoft.Protocols.TestSuites.MS_OXWSCONT
{
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSCONT.
    /// </summary>
    public partial class MS_OXWSCONTAdapter : ManagedAdapterBase, IMS_OXWSCONTAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        #endregion

        #region IMS_OXWSCONTAdapter Properties
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
        public override void Initialize(ITestSite testSite)
        {
            // Initialize
            base.Initialize(testSite);

            testSite.DefaultProtocolDocShortName = "MS-OXWSCONT";

            Common.MergeConfiguration(testSite);

            string userName = Common.GetConfigurationPropertyValue("ContactUserName", testSite);
            string password = Common.GetConfigurationPropertyValue("ContactUserPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSCONTAdapter Operations
        /// <summary>
        /// Get contact item on the server.
        /// </summary>
        /// <param name="getItemRequest">The request of GetItem operation.</param>
        /// <returns>A response to GetItem operation request.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            GetItemResponseType getItemResponse = this.exchangeServiceBinding.GetItem(getItemRequest);

            Site.Assert.IsNotNull(getItemResponse, "If the operation is successful, the response should not be null.");

            #region Verify GetItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyGetContactItem(getItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return getItemResponse;
        }

        /// <summary>
        /// Delete contact item on the server.
        /// </summary>
        /// <param name="deleteItemRequest">The request of DeleteItem operation.</param>
        /// <returns>A response to DeleteItem operation request.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            DeleteItemResponseType deleteItemResponse = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);

            #region Verify DeleteItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyDeleteContactItem(this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return deleteItemResponse;
        }

        /// <summary>
        /// Create contact item on the server.
        /// </summary>
        /// <param name="createItemRequest">The request of CreateItem operation.</param>
        /// <returns>A response to CreateItem operation request.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(createItemRequest);

            #region Verify CreateItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyCreateContactItem(this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return createItemResponse;
        }

        /// <summary>
        /// Update contact item on the server.
        /// </summary>
        /// <param name="updateItemRequest">The request of UpdateItem operation.</param>
        /// <returns>A response to UpdateItem operation request.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest)
        {
            UpdateItemResponseType updateItemResponse = this.exchangeServiceBinding.UpdateItem(updateItemRequest);

            #region Verify UpdateItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyUpdateContactItem(this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return updateItemResponse;
        }

        /// <summary>
        /// Copy contact item on the server.
        /// </summary>
        /// <param name="copyItemRequest">The request of CopyItem operation.</param>
        /// <returns>A response to CopyItem operation request.</returns>
        public CopyItemResponseType CopyItem(CopyItemType copyItemRequest)
        {
            CopyItemResponseType copyItemRespose = this.exchangeServiceBinding.CopyItem(copyItemRequest);

            #region Verify CopyItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyCopyContactItem(copyItemRespose,this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return copyItemRespose;
        }

        /// <summary>
        /// Move contact item on the server.
        /// </summary>
        /// <param name="moveItemRequest">The request of MoveItem operation.</param>
        /// <returns>A response to MoveItem operation request.</returns>
        public MoveItemResponseType MoveItem(MoveItemType moveItemRequest)
        {
            MoveItemResponseType moveItemResoponse = this.exchangeServiceBinding.MoveItem(moveItemRequest);

            #region Verify MoveItem operation requirements

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyMoveContactItem(this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return moveItemResoponse;
        }

        /// <summary>
        /// Retrieves the profile image for a mailbox
        /// </summary>
        /// <param name="getUserPhotoRequest">The request of GetUserPhoto operation.</param>
        /// <returns>A response to GetUserPhoto operation request.</returns>
        public GetUserPhotoResponseMessageType GetUserPhoto(GetUserPhotoType getUserPhotoRequest)
        {
            GetUserPhotoResponseMessageType getUserPhotoResponse = this.exchangeServiceBinding.GetUserPhoto(getUserPhotoRequest);
            
            #region Verifiy GetUserPhoto opreation requirements
            this.VerifySoapVersion();
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyGetUserPhotoResponseMessageType(getUserPhotoResponse, this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return getUserPhotoResponse;
        }

        /// <summary>
        /// Add a photo to a user's account.
        /// </summary>
        /// <param name="setUserPhotoRequest">The request of SetUserPhoto operation.</param>
        /// <returns>A response to SetUserPhoto operation request.</returns>
        public SetUserPhotoResponseMessageType SetUserPhoto(SetUserPhotoType setUserPhotoRequest)
        {
            SetUserPhotoResponseMessageType setUserPhotoResponse = this.exchangeServiceBinding.SetUserPhoto(setUserPhotoRequest);

            #region Verifiy SetUserPhoto opreation requirements
            this.VerifySoapVersion();
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifySetUserPhotoResponseMessageType(setUserPhotoResponse, this.exchangeServiceBinding.IsSchemaValidated);
            #endregion

            return setUserPhotoResponse;
        }
        #endregion
    }
}