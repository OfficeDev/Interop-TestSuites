//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_SITESS
{
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The partial test class contains test case definitions related to GetSite Templates operation.
    /// </summary>
    [TestClass]
    public class S07_HTTPStatusCode : TestClassBase
    {
        /// <summary>
        /// An instance of protocol adapter class.
        /// </summary>
        private IMS_SITESSAdapter sitessAdapter;

        /// <summary>
        /// An instance of SUT control adapter class.
        /// </summary>
        private IMS_SITESSSUTControlAdapter sutAdapter;

        #region Test Suite Initialization & Cleanup

        /// <summary>
        /// Test Suite Initialization.
        /// </summary>
        /// <param name="testContext">The test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Test Suite Cleanup.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Cases

        #region Scenario 7 HTTP Status Code

        /// <summary>
        /// This test case is designed to verify the protocol server fault by using HTTP Status Codes. A HTTP status code "Unauthorized" is expected to throw.
        /// </summary>
        [TestCategory("MSSITESS"), TestMethod()]
        public void MSSITESS_S07_TC01_HttpStatusCodesFault()
        {
            #region Variables
            uint lcid = uint.Parse(Common.GetConfigurationPropertyValue(Constants.ValidLCID, this.Site));
            Template[] templates;
            bool isExpectedHttpCode = false;

            #endregion Variables

            // Initialize the web service with an unauthenticated account.
            this.sitessAdapter.InitializeWebService(UserAuthenticationOption.Unauthenticated);

            // Try to invoke the GetSite operation, because the user is Unauthenticated, so a HTTP status code "Unauthorized" is expected to be thrown.
            try
            {
                this.sitessAdapter.GetSiteTemplates(lcid, out templates);
            }
            catch (WebException exp)
            {
                if (exp.Response == null)
                {
                    throw;
                }

                string strExp = ((HttpWebResponse)exp.Response).StatusCode.ToString();

                Site.Assert.AreEqual<string>(
                    "Unauthorized",
                    strExp,
                    "Unauthorized error is expected to be returned.");

                isExpectedHttpCode = true;
            }

            #region Capture requirements

            // If expected HTTP Status Code is thrown, R365 can be captured.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-SITESS_R365, the expected HTTP Status Code {0} thrown", isExpectedHttpCode ? "is" : "is not");

            // Verify MS-SITESS requirement: MS-SITESS_R365
            Site.CaptureRequirementIfIsTrue(
                isExpectedHttpCode,
                365,
                @"[In Transport] Protocol server faults can be returned using HTTP Status Codes as specified in [RFC2616] section 10.");
            #endregion Capture requirements
        }

        #endregion Scenario 7 HTTP Status Code

        #endregion Test Cases

        #region Test Case Initialization & Cleanup

        /// <summary>
        /// Test case initialize method.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.sitessAdapter = Site.GetAdapter<IMS_SITESSAdapter>();
            Common.CheckCommonProperties(this.Site, true);
            this.sutAdapter = Site.GetAdapter<IMS_SITESSSUTControlAdapter>();
        }

        /// <summary>
        /// Test case cleanup method.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.sitessAdapter.Reset();
            this.sutAdapter.Reset();
        }

        #endregion
    }
}