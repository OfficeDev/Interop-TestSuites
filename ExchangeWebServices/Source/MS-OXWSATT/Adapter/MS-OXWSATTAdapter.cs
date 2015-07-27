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
    using System.Collections.Generic;
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter class of MS-OXWSATT.
    /// </summary>
    public partial class MS_OXWSATTAdapter : ManagedAdapterBase, IMS_OXWSATTAdapter
    {
        #region Fields
        /// <summary>
        /// The exchange service binding.
        /// </summary>
        private ExchangeServiceBinding exchangeServiceBinding;
        #endregion

        #region IMS_OXWSATTAdapter Properties
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
        public override void Initialize(TestTools.ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-OXWSATT";
            Common.MergeConfiguration(testSite);

            string userName = Common.GetConfigurationPropertyValue("UserName", testSite);
            string password = Common.GetConfigurationPropertyValue("UserPassword", testSite);
            string domain = Common.GetConfigurationPropertyValue("Domain", testSite);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", testSite);

            this.exchangeServiceBinding = new ExchangeServiceBinding(url, userName, password, domain, testSite);
            Common.InitializeServiceBinding(this.exchangeServiceBinding, testSite);
        }

        #endregion

        #region IMS_OXWSATTAdapter Operations
        /// <summary>
        /// Creates an item or file attachment on an item in the server store. 
        /// </summary>
        /// <param name="createAttachmentRequest">A CreateAttachmentType complex type specifies a request message to attach an item or file to a specified item in the server database. </param>
        /// <returns>A CreateAttachmentResponseType complex type specifies the response message that is returned by the CreateAttachment operation. </returns>
        public CreateAttachmentResponseType CreateAttachment(CreateAttachmentType createAttachmentRequest)
        {
            CreateAttachmentResponseType createAttachmentResponse = this.exchangeServiceBinding.CreateAttachment(createAttachmentRequest);

            Site.Assert.IsNotNull(createAttachmentResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyCreateAttachmentResponse(createAttachmentResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return createAttachmentResponse;
        }

        /// <summary>
        /// Gets an attachment from an item in the server store.
        /// </summary>
        /// <param name="getAttachmentRequest">A GetAttachmentType complex type specifies a request message to get attached items and files on an item in the server database.</param>
        /// <returns>A GetAttachmentResponseType complex type specifies the response message that is returned by the GetAttachment operation.</returns>
        public GetAttachmentResponseType GetAttachment(GetAttachmentType getAttachmentRequest)
        {
            GetAttachmentResponseType getAttachmentResponse = this.exchangeServiceBinding.GetAttachment(getAttachmentRequest);

            Site.Assert.IsNotNull(getAttachmentResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyGetAttachmentResponse(getAttachmentResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return getAttachmentResponse;
        }

        /// <summary>
        /// Deletes an attachment from an item in the server store. 
        /// </summary>
        /// <param name="deleteAttachmentRequest">A DeleteAttachmentType complex type specifies a request message to delete an attachment on an item in the server database.</param>
        /// <returns>A DeleteAttachmentResponseType complex type specifies the response message that is returned by the DeleteAttachment operation.</returns>
        public DeleteAttachmentResponseType DeleteAttachment(DeleteAttachmentType deleteAttachmentRequest)
        {
            DeleteAttachmentResponseType deleteAttachmentResponse = this.exchangeServiceBinding.DeleteAttachment(deleteAttachmentRequest);

            Site.Assert.IsNotNull(deleteAttachmentResponse, "If the operation is successful, the response should not be null.");

            this.VerifySoapVersion();
            this.VerifyTransportType();
            this.VerifyServerVersionInfo(this.exchangeServiceBinding.IsSchemaValidated);
            this.VerifyDeleteAttachmentResponse(deleteAttachmentResponse, this.exchangeServiceBinding.IsSchemaValidated);
            return deleteAttachmentResponse;
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