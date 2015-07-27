//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSMTGS
{
    using System;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class of IMS_OXWSMTGSAdapter.
    /// </summary>
    public partial class MS_OXWSMTGSAdapter : ManagedAdapterBase, IMS_OXWSMTGSAdapter
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

        #region IMS_OXWSMTGSAdapter Properties
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
            Site.DefaultProtocolDocShortName = "MS-OXWSMTGS";
            Common.MergeConfiguration(testSite);

            this.username = Common.GetConfigurationPropertyValue("OrganizerName", this.Site);
            this.password = Common.GetConfigurationPropertyValue("OrganizerPassword", this.Site);
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.username, this.password, this.domain, this.Site);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, this.Site);
        }
        #endregion

        #region IMS_OXWSMTGSAdapter Operations
        /// <summary>
        /// Get the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the GetItem operation.</param>
        /// <returns>The response message returned by GetItem operation.</returns>
        public GetItemResponseType GetItem(GetItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'GetItem' should not be null.");
            }

            GetItemResponseType getItemResponse = this.exchangeServiceBinding.GetItem(request);
            Site.Assert.IsNotNull(getItemResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyGetItemOperation(getItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return getItemResponse;
        }
        
        /// <summary>
        /// Delete the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the DeleteItem operation.</param>
        /// <returns>The response message returned by DeleteItem operation.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'DeleteItem' should not be null.");
            }

            DeleteItemResponseType deleteItemResponse = this.exchangeServiceBinding.DeleteItem(request);

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyDeleteItemOperation(this.exchangeServiceBinding.IsSchemaValidated);
            return deleteItemResponse;
        }

        /// <summary>
        /// Update the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the UpdateItem operation.</param>
        /// <returns>The response message returned by UpdateItem operation.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'UpdateItem' should not be null.");
            }

            UpdateItemResponseType updateItemResponse = this.exchangeServiceBinding.UpdateItem(request);
            Site.Assert.IsNotNull(updateItemResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyUpdateItemOperation(updateItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return updateItemResponse;
        }

        /// <summary>
        /// Move the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the MoveItem operation.</param>
        /// <returns>The response message returned by MoveItem operation.</returns>
        public MoveItemResponseType MoveItem(MoveItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'MoveItem' should not be null.");
            }

            MoveItemResponseType moveItemResponse = this.exchangeServiceBinding.MoveItem(request);
            Site.Assert.IsNotNull(moveItemResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyMoveItemOperation(moveItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return moveItemResponse;
        }
                
        /// <summary>
        /// Copy the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the CopyItem operation.</param>
        /// <returns>The response message returned by CopyItem operation.</returns>
        public CopyItemResponseType CopyItem(CopyItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'CopyItem' should not be null.");
            }

            CopyItemResponseType copyItemResponse = this.exchangeServiceBinding.CopyItem(request);
            Site.Assert.IsNotNull(copyItemResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyCopyItemOperation(copyItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return copyItemResponse;
        }

        /// <summary>
        /// Create the calendar related item elements.
        /// </summary>
        /// <param name="request">A request to the CreateItem operation.</param>
        /// <returns>The response message returned by CreateItem operation.</returns>
        public CreateItemResponseType CreateItem(CreateItemType request)
        {
            if (request == null)
            {
                throw new ArgumentException("The request of operation 'CreateItem' should not be null.");
            }

            CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(request);
            Site.Assert.IsNotNull(createItemResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyCreateItemOperation(createItemResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return createItemResponse;
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