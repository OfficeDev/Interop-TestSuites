//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OFFICIALFILE
{
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This test class is used to test GetRecordRouting and GetRecordRoutingCollection operations.
    /// </summary>
    [TestClass]
    public class S02_GetRoutingInfo : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">An instance of an object that derives from the Microsoft.VisualStudio.TestTools.UnitTesting.TestContext class.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear up the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is used to test GetRouting on a repository that is configured for routing content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S02_TC01_GetRouting()
        {
            if (Common.Common.IsRequirementEnabled(362, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetRecordingRouting operation.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.DisableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            string routing = this.Adapter.GetRecordingRouting(this.DocumentContentTypeName);
            Site.CaptureRequirementIfIsNotNull(
                        routing,
                        "MS-OFFICIALFILE",
                        362,
                        @"[In GetRecordRouting] This method [GetRecordRouting] is not deprecated and can be called. (Office SharePoint Server 2007 follows this behavior.)");
        }

        /// <summary>
        /// This test case is used to test GetRoutingCollection on a repository that is configured for routing content.
        /// </summary>
        [TestCategory("MSOFFCIALFILE"), TestMethod()]
        public void MSOFFICIALFILE_S02_TC02_GetRoutingCollection()
        {
            if (Common.Common.IsRequirementEnabled(364, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the GetRoutingCollection operation.");
            }

            // Initial parameters to use a repository that is configured for routing content.
            InitialPara paras = new InitialPara();
            paras.Url = this.EnableRoutingFeatureRecordsCenterServiceUrl;
            paras.UserName = this.SubmitUserName;
            paras.Domain = this.DomainName;
            paras.Password = this.Password;

            this.Adapter.IntializeService(paras);

            string routingCollection = this.Adapter.GetRecordRoutingCollection();
            Site.CaptureRequirementIfIsNotNull(
                        routingCollection,
                        "MS-OFFICIALFILE",
                        364,
                        @"[In GetRecordRoutingCollection] This method GetRecordRoutingCollection] is not deprecated and can be called. (Office SharePoint Server 2007 follows this behavior.)");
        }
    }
}
