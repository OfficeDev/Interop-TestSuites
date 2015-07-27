//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Net;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The base test suite class defines common initialization method and cleanup method for all the two scenarios.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// Adapter Instance.
        /// </summary>
        private IMS_OXCMAPIHTTPAdapter adapter;

        /// <summary>
        /// The instance of the IMS_OXCMAPIHTTPSUTControlAdapter.
        /// </summary>
        private IMS_OXCMAPIHTTPSUTControlAdapter sutControlAdapter;

        /// <summary>
        /// An instance of RopBufferHelper that is used to compose and parse ROP buffer.
        /// </summary>
        private RopBufferHelper ropBufferHelper;

        /// <summary>
        /// AdminUser name which can be used by client to access to the Exchange server via MAPIHTTP.
        /// </summary>
        private string adminUserName = string.Empty;

        /// <summary>
        /// The password for AdminUserName which can be used by client to access to the Exchange via MAPIHTTP.
        /// </summary>
        private string adminUserPassword = string.Empty;

        /// <summary>
        /// A string containing the distinguished name (DN) of the user configured by "AdminUserName".
        /// </summary>
        private string adminUserDN = string.Empty;
        #endregion Variables

        /// <summary>
        /// Gets the instance of IMS_OXCMAPIHTTPAdapter protocol adapter.
        /// </summary>
        protected IMS_OXCMAPIHTTPAdapter Adapter
        {
            get { return this.adapter; }
        }

        /// <summary>
        /// Gets the instance of IMS_OXCMAPIHTTPSUTControlAdapter SUT control adapter.
        /// </summary>
        protected IMS_OXCMAPIHTTPSUTControlAdapter SUTControlAdapter
        {
            get { return this.sutControlAdapter; }
        }

        /// <summary>
        /// Gets instance of RopBufferHelper which is used to compose and parse ROP buffer.
        /// </summary>
        protected RopBufferHelper RopBufferHelper
        {
            get { return this.ropBufferHelper; }
        }

        /// <summary>
        /// Gets AdminUser name which can be used by client to access to the Exchange server via MAPIHTTP.
        /// </summary>
        protected string AdminUserName
        {
            get { return this.adminUserName; }
        }

        /// <summary>
        /// Gets the password for AdminUserName which can be used by client to access to the Exchange via MAPIHTTP.
        /// </summary>
        protected string AdminUserPassword
        {
            get { return this.adminUserPassword; }
        }

        /// <summary>
        /// Gets a string containing the distinguished name (DN) of the user configured by "AdminUserName".
        /// </summary>
        protected string AdminUserDN
        {
            get { return this.adminUserDN; }
        }

        #region Test Case Initialization

        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            this.adapter = Site.GetAdapter<IMS_OXCMAPIHTTPAdapter>();
            this.sutControlAdapter = Site.GetAdapter<IMS_OXCMAPIHTTPSUTControlAdapter>();
            this.ropBufferHelper = new RopBufferHelper(this.Site);
            this.adminUserName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            this.adminUserPassword = Common.GetConfigurationPropertyValue("AdminUserPassword", this.Site);
            this.adminUserDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site);
        }

        /// <summary>
        /// Clean up the test suite.
        /// </summary>
        protected override void TestCleanup()
        {
            base.TestCleanup();
        }

        #endregion Test Case Initialization.

        #region Common methods for test cases.
        /// <summary>
        /// Check the preconditions of this test suite.
        /// </summary>
        protected void CheckMapiHttpIsSupported()
        {
            bool isSupported = bool.Parse(Common.GetConfigurationPropertyValue("MS-OXCMAPIHTTP_Supported", this.Site));

            this.Site.Assume.IsTrue(
                 isSupported,
                 "The MS-OXCMAPIHTTP is not supported when the MS-OXCMAPIHTTP_Supported property is false in Should/May PTFconfig file.");

            this.Site.Assume.AreEqual<string>(
                "mapi_http",
                Common.GetConfigurationPropertyValue("TransportSeq", this.Site),
                "The MS-OXCMAPIHTTP is not supported when the transport protocol sequence is set to either ncacn_ip_tcp or ncacn_http. The transport protocol sequence is determined by Common PTFConfig property named TransportSeq.");
        }
        #endregion Common methods for test cases.
    }
}