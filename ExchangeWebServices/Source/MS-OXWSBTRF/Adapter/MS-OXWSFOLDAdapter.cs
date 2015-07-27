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
    using System.Xml.XPath;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The adapter class implements the method defined in interface IMS_OXWSFOLDAdapter.
    /// </summary>
    public class MS_OXWSFOLDAdapter : ManagedAdapterBase, IMS_OXWSFOLDAdapter
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

        #region IMS_OXWSFOLDAdapter Properties
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

        #region IMS_OXWSFOLDAdapter Operations
        /// <summary>
        /// Creates folders on the server.
        /// </summary>
        /// <param name="createFolderRequest">Specify the request for the CreateFolder operation.</param>
        /// <returns>The response to this operation request.</returns>
        public CreateFolderResponseType CreateFolder(CreateFolderType createFolderRequest)
        {
            if (createFolderRequest == null)
            {
                throw new ArgumentException("The CreateFolder request should not be null.");
            }

            // Send the request and get the response.
            CreateFolderResponseType response = this.exchangeServiceBinding.CreateFolder(createFolderRequest);
            return response;
        }
        #endregion
    }
}