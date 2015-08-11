namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSMSG. 
    /// </summary>
    public partial class MS_OXWSMSGAdapter : ManagedAdapterBase, IMS_OXWSMSGAdapter
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

        #region IMS_OXWSMSGAdapter Properties
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
            Site.DefaultProtocolDocShortName = "MS-OXWSMSG";
            Common.MergeConfiguration(testSite);

            this.userName = Common.GetConfigurationPropertyValue("Sender", this.Site);
            this.password = Common.GetConfigurationPropertyValue("SenderPassword", this.Site);
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }
        #endregion

        #region IMS_OXWSMSGAdapter Operations
        /// <summary>
        /// Get message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to get message objects.</param>
        /// <returns>The response message returned by GetItem operation.</returns>
        public GetItemResponseType GetItem(GetItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'GetItem' should not be null.");
            }

            GetItemResponseType response = this.exchangeServiceBinding.GetItem(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyGetItemOperation(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Copy message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to copy message objects.</param>
        /// <returns>The response message returned by CopyItem operation.</returns>
        public CopyItemResponseType CopyItem(CopyItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'CopyItem' should not be null.");
            }

            CopyItemResponseType response = this.exchangeServiceBinding.CopyItem(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyCopyItemOperation(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Create message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to create message objects.</param>
        /// <returns>The response message returned by CreateItem operation.</returns>
        public CreateItemResponseType CreateItem(CreateItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'CreateItem' should not be null.");
            }

            CreateItemResponseType response = this.exchangeServiceBinding.CreateItem(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyCreateItemOperation(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Delete message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to delete message objects.</param>
        /// <returns>The response message returned by DeleteItem operation.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'DeleteItem' should not be null.");
            }

            DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(request);
            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyDeleteItemOperation(this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Move message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to move message objects.</param>
        /// /// <returns>The response message returned by MoveItem operation.</returns>
        public MoveItemResponseType MoveItem(MoveItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'MoveItem' should not be null.");
            }

            MoveItemResponseType response = this.exchangeServiceBinding.MoveItem(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyMoveItemOperation(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Update message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to update message objects.</param>
        /// /// <returns>The response message returned by UpdateItem operation.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'UpdateItem' should not be null.");
            }

            UpdateItemResponseType response = this.exchangeServiceBinding.UpdateItem(request);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyUpdateItemOperation(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Send message related Item elements on the server.
        /// </summary>
        /// <param name="request">Specify a request to send message objects.</param>
        /// /// <returns>The response message returned by SendItem operation.</returns>
        public SendItemResponseType SendItem(SendItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'SendItem' should not be null.");
            }

            SendItemResponseType response = this.exchangeServiceBinding.SendItem(request);
            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifySendItemOperation(this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }
        #endregion
    }
}