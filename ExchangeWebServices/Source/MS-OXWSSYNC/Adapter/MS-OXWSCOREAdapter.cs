//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    using System;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the methods defined in interface IMS_OXWSCOREAdapter.
    /// </summary>
    public class MS_OXWSCOREAdapter : ManagedAdapterBase, IMS_OXWSCOREAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;

        /// <summary>
        /// The endpoint url of Exchange Web Service.
        /// </summary>
        private string url;

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
        /// <param name="testSite">Pass ITestSite to adapter, make adapter can use ITestSite's function</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSSYNC";

            // Merge configuration files.
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("User1Name", testSite);
            this.password = Common.GetConfigurationPropertyValue("User1Password", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSCOREAdapter Operations
        /// <summary>
        /// Create items on the server.
        /// </summary>
        /// <param name="createItemRequest">Specify a request to create items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            if (createItemRequest == null)
            {
                throw new ArgumentException("The CreateItem request should not be null.");
            }

            CreateItemResponseType response = this.exchangeServiceBinding.CreateItem(createItemRequest);
            return response;
        }

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Specify a request to delete item on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            if (deleteItemRequest == null)
            {
                throw new ArgumentException("The DeleteItem request should not be null.");
            }

            DeleteItemResponseType response = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);
            return response;
        }

        /// <summary>
        /// Get items on the server.
        /// </summary>
        /// <param name="getItemRequest">Specify a request to get items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public GetItemResponseType GetItem(GetItemType getItemRequest)
        {
            if (getItemRequest == null)
            {
                throw new ArgumentException("The GetItem request should not be null.");
            }

            GetItemResponseType response = this.exchangeServiceBinding.GetItem(getItemRequest);
            return response;
        }

        /// <summary>
        /// Update items on the server.
        /// </summary>
        /// <param name="updateItemRequest">Specify a request to update items on the server.</param>
        /// <returns>A response to this operation request.</returns>
        public UpdateItemResponseType UpdateItem(UpdateItemType updateItemRequest)
        {
            if (updateItemRequest == null)
            {
                throw new ArgumentException("The UpdateItem request should not be null.");
            }

            UpdateItemResponseType response = this.exchangeServiceBinding.UpdateItem(updateItemRequest);
            return response;
        }
        #endregion
    }
}