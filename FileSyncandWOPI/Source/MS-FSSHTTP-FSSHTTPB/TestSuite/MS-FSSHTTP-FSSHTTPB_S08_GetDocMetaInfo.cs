namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the GetDocMetaInfo sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S08_GetDocMetaInfo : S08_GetDocMetaInfo
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S08_GetDocMetaInfo.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S08_GetDocMetaInfo.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify that GetDocMetaInfo sub-request when the URL is empty.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S08_TC01_GetDocMetaInfo_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Send the GetDocMetaInfo subRequest to the protocol server with Url attribute set to en empty string.
            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = SharedTestSuiteHelper.CreateGetDocMetaInfoSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { getDocMetaInfoSubRequest });

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
            {
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
                else
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify that GetDocMetaInfo sub request when the URL does not exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S08_TC02_GetDocMetaInfo_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Send the GetDocMetaInfo subRequest to the protocol server with Url attribute is not set.
            GetDocMetaInfoSubRequestType getDocMetaInfoSubRequest = SharedTestSuiteHelper.CreateGetDocMetaInfoSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { getDocMetaInfoSubRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3007
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3007,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                Site.Assert.IsNull(
                    response.ResponseCollection,
                    @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
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