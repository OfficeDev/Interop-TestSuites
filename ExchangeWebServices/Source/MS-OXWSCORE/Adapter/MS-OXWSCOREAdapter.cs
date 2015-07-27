//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.MS_OXWSITEMID;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSCORE.
    /// </summary>
    public partial class MS_OXWSCOREAdapter : ManagedAdapterBase, IMS_OXWSCOREAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        /// <summary>
        /// MS-OXWSITEMID adapter.
        /// </summary>
        private IMS_OXWSITEMIDAdapter itemAdapter;
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
        public override void Initialize(ITestSite testSite)
        {          
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSCORE";

            Common.MergeConfiguration(testSite);

            string userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            string password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
            this.itemAdapter = Site.GetAdapter<IMS_OXWSITEMIDAdapter>(); 
        }

        #endregion

        #region IMS_OXWSCOREAdapter Operations
        /// <summary>
        /// Copy items and puts the items in a different folder.
        /// </summary>
        /// <param name="copyItemRequest">Specify a request to copy items on the server.</param>
        /// <returns>A response to CopyItem operation request.</returns>
        public CopyItemResponseType CopyItem(CopyItemType copyItemRequest)
        {
            CopyItemResponseType response = this.exchangeServiceBinding.CopyItem(copyItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyCopyItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Create items in the Exchange store
        /// </summary>
        /// <param name="createItemRequest">Specify a request to create items on the server.</param>
        /// <returns>A response to CreateItem operation request.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            CreateItemResponseType response = this.exchangeServiceBinding.CreateItem(createItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyCreateItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyItemId(response);
            return response;
        }

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specify a request to delete item on the server.</param>
        /// <returns>A response to DeleteItem operation request.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyDeleteItemResoponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Get items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify a request to get items on the server.</param>
        /// <returns>A response to GetItem operation request.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            GetItemResponseType response = this.exchangeServiceBinding.GetItem(getItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyGetItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyItemId(response);
            return response;
        }

        /// <summary>
        /// Move items on the server.
        /// </summary>
        /// <param name="moveItemRequest">Specify a request to move items on the server.</param>
        /// <returns>A response to MoveItem operation request.</returns>
        public MoveItemResponseType MoveItem(MoveItemType moveItemRequest)
        {
            MoveItemResponseType response = this.exchangeServiceBinding.MoveItem(moveItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyMoveItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Send messages and post items on the server.
        /// </summary>
        /// <param name="sendItemRequest">Specify a request to send items on the server.</param>
        /// <returns>A response to SendItem operation request.</returns>
        public SendItemResponseType SendItem(SendItemType sendItemRequest)
        {
            SendItemResponseType response = this.exchangeServiceBinding.SendItem(sendItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifySendItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Update items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specify a request to update items on the server.</param>
        /// <returns>A response to UpdateItem operation request.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest)
        {
            UpdateItemResponseType response = this.exchangeServiceBinding.UpdateItem(updateItemRequest);
            Site.Assert.IsNotNull(response, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyUpdateItemResponse(response, this.exchangeServiceBinding.IsSchemaValidated);
            return response;
        }

        /// <summary>
        /// Mark all items in a folder as read.
        /// </summary>
        /// <param name="markAllItemAsReadRequest">Specify a request to mark all items as read.</param>
        /// <returns>A response to MarkAllItemsAsRead operation request.</returns>
        public MarkAllItemsAsReadResponseType MarkAllItemsAsRead(MarkAllItemsAsReadType markAllItemAsReadRequest)
        {
            MarkAllItemsAsReadResponseType markAllItemsAsReadResponse = this.exchangeServiceBinding.MarkAllItemsAsRead(markAllItemAsReadRequest);
            Site.Assert.IsNotNull(markAllItemsAsReadResponse, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyMarkAllItemsAsReadResponse(markAllItemsAsReadResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return markAllItemsAsReadResponse;
        }

        /// <summary>
        /// The MarkAsJunk operation marks an item as junk.
        /// </summary>
        /// <param name="markAsJunkRequest">Specify a request for a MarkAsJunk operation.</param>
        /// <returns>A response to MarkAsJunk operation request.</returns>
        public MarkAsJunkResponseType MarkAsJunk(MarkAsJunkType markAsJunkRequest)
        {
            MarkAsJunkResponseType markAsJunkReponse = this.exchangeServiceBinding.MarkAsJunk(markAsJunkRequest);
            Site.Assert.IsNotNull(markAsJunkReponse, "If the operation is successful, the response should not be null.");

            // SOAP version is set to 1.1, if a response can be received from server, then it means SOAP 1.1 is supported.
            this.VerifySoapVersion();

            // Verify transport type related requirement.
            this.VerifyTransportType();

            this.VerifyServerVersionInfo(this.exchangeServiceBinding.ServerVersionInfoValue, this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyMarkAsJunkResponse(markAsJunkReponse, this.exchangeServiceBinding.IsSchemaValidated);
            return markAsJunkReponse;
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

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <param name="headerValues">Specify the values for SOAP header.</param>
        public void ConfigureSOAPHeader(Dictionary<string, object> headerValues)
        {
            Common.ConfigureSOAPHeader(headerValues, this.exchangeServiceBinding);
        }

        #endregion
    }
}