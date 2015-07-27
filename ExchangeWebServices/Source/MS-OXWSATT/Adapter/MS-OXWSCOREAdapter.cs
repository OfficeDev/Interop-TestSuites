//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
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
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
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
        /// Implements Microsoft.Protocols.TestTools.IAdapter.Initialize(Microsoft.Protocols.TestTools.ITestSite).
        /// </summary>
        /// <param name="testSite">The test site instance associated with the current adapter.</param>
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);

            string commonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", testSite);

            Common.MergeGlobalConfig(commonConfigFileName, testSite);

            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSCOREAdapter Operations
        /// <summary>
        /// Creates items on the server.
        /// </summary>
        /// <param name="createItemRequest">Request message of "CreateItem" operation.</param>
        /// <returns>Response message of "CreateItem" operation.</returns>
        public CreateItemResponseType CreateItem(CreateItemType createItemRequest)
        {
            CreateItemResponseType createItemResponse = this.exchangeServiceBinding.CreateItem(createItemRequest);
            return createItemResponse;
        }

        /// <summary>
        /// Deletes items on the server.
        /// </summary>
        /// <param name="deleteItemRequest">Request message of "DeleteItem" operation.</param>
        /// <returns>Response message of "DeleteItem" operation.</returns>
        public DeleteItemResponseType DeleteItem(DeleteItemType deleteItemRequest)
        {
            DeleteItemResponseType deleteItemResponse = this.exchangeServiceBinding.DeleteItem(deleteItemRequest);
            return deleteItemResponse;
        }

        #endregion
    }
}