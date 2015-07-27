//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSBTRF
{
    using System;
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSBTRF.
    /// </summary>
    public partial class MS_OXWSBTRFAdapter : ManagedAdapterBase, IMS_OXWSBTRFAdapter
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
        /// The endpoint url of Exchange Web Service
        /// </summary>
        private string url;
        #endregion

        #region IMS_OXWSBTRFAdapter Properties
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
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSBTRF";

            // Merge common configuration and SHOULD/MAY configuration files
            Common.MergeConfiguration(testSite);

            // Get the parameters from configuration files.
            this.userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            this.password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            this.domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            this.url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(this.url, this.userName, this.password, this.domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }
        #endregion

        #region IMS_OXWSBTRFAdapter Operations
        /// <summary>
        /// Exports items from a specified folder.
        /// </summary>
        /// <param name="exportItemsRequest">Specify the request for ExportItems operation.</param>
        /// <returns>The response to this operation request.</returns>
        public ExportItemsResponseType ExportItems(ExportItemsType exportItemsRequest)
        {
            if (exportItemsRequest == null)
            {
                throw new ArgumentException("The exportItems should not be null.");
            }

            // Send the request and get the response.
            ExportItemsResponseType response = this.exchangeServiceBinding.ExportItems(exportItemsRequest);
            this.VerifyTransportType();
            this.VerifySoapVersion();
            this.VerifyExportItemsResponseType(response, this.exchangeServiceBinding.IsSchemaValidated);

            return response;
        }

        /// <summary>
        /// Uploads the items into a specified folder.
        /// </summary>
        /// <param name="uploadItemsRequest">Specify the request for UploadItems operation.</param>
        /// <returns>The response to this operation request.</returns>
        public UploadItemsResponseType UploadItems(UploadItemsType uploadItemsRequest)
        {
            if (uploadItemsRequest == null)
            {
                throw new ArgumentException("The uploadItems should not be null.");
            }

            // Send the request and get the response.
            UploadItemsResponseType response = this.exchangeServiceBinding.UploadItems(uploadItemsRequest);
            this.VerifyTransportType();
            this.VerifySoapVersion();
            this.VerifyUploadItemsResponseType(response, this.exchangeServiceBinding.IsSchemaValidated);

            return response;
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