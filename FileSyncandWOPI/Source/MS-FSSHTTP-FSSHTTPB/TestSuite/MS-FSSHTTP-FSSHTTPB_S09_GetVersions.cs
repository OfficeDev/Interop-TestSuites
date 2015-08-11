namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the GetVersions sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S09_GetVersions : S09_GetVersions
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S09_GetVersions.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S09_GetVersions.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify that GetVersions sub-request when the URL is empty string.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S09_TC01_GetVersions_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetVersions"sub-request with correct input parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse response = Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { getVersionsSubRequest });

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
        /// A method used to verify that GetVersions sub-request when the URL is not specified.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S09_TC02_GetVersions_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetVersions"sub-request with correct input parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse response = Adapter.CellStorageRequest(null, new SubRequestType[] { getVersionsSubRequest });

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