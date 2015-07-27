//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the EditorsTable sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S07_EditorsTable : S07_EditorsTable
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S07_EditorsTable.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S07_EditorsTable.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify if server was unable to find the URL for the file specified in the Url attribute, the protocol server will report a failure by returning an error code value set to 'FileNotExistsOrCannotBeCreated'.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S07_TC01_EditorsTable_FileNotExistsOrCannotBeCreated()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Create a join editor session with the URL which could not be found.
            string url = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the service
            this.InitializeContext(url, this.UserName01, this.Password01, this.Domain);

            EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultTimeOut);

            // Call protocol adapter operation CellStorageRequest to join the editing session.
            CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(url, new SubRequestType[] { join });
            EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the ErrorCode attribute returned equals "FileNotExistsOrCannotBeCreated", MS-FSSHTTP_R1971 and MS-FSSHTTP_R358 can be covered.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotExistsOrCannotBeCreated,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1971,
                         @"[In EditorsTable Subrequest] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotExistsOrCannotBeCreated,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         358,
                         @"[In GenericErrorCodeTypes] FileNotExistsOrCannotBeCreated indicates an error when either the targeted URL for the file specified as part of the Request element does not exist or file creation failed on the protocol server.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotExistsOrCannotBeCreated,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                    @"[In GenericErrorCodeTypes] FileNotExistsOrCannotBeCreated indicates an error when either the targeted URL for the file specified as part of the Request element does not exist or file creation failed on the protocol server.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Url attribute of the corresponding Request element is an empty string.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S07_TC02_EditorsTable_JoinEditSession_EmptyUrl()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a EditorsTable subRequest with all valid parameters.
            EditorsTableSubRequestType subRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, 3600);

            // Send the serverTime subRequest to the protocol server with Url attribute set to en empty string.
            CellStorageResponse response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { subRequest });

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
            {
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <3> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
                else
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <3> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Url attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S07_TC03_EditorsTable_JoinEditSession_UrlNotSpecified()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a EditorsTable subRequest with all valid parameters.
            EditorsTableSubRequestType subRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, 3600);

            // Send a ServerTime subRequest to the protocol server without specifying Url attribute.
            CellStorageResponse response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { subRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3007
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3007,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <3> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                Site.Assert.IsNull(
                    response.ResponseCollection,
                    @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <3> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
            }
        }

        /// <summary>
        /// Initialize the shared context based on the specified request file URL, user name, password and domain for the MS-FSSHTTP test purpose.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        protected override void InitializeContext(string requestFileUrl, string userName, string password, string domain)
        {
            SharedContextUtils.InitializeSharedContextForFSSHTTP(userName, password, domain, this.Site);
        }

        /// <summary>
        /// Merge the common configuration and should/may configuration file.
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        protected override void MergeConfigurationFile(TestTools.ITestSite site)
        {
            ConfigurationFileHelper.MergeConfigurationFile(site);
        }
    }
}