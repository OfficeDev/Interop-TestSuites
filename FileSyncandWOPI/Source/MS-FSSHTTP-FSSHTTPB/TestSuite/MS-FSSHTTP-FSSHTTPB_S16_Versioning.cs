namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with CellSubRequest operation for versioning request. 
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S16_Versioning : S16_Versioning 
    {
        #region Test Suite Initialization

        /// <summary>
        /// A method used to initialize this class.
        /// </summary>
        /// <param name="testContext">A parameter represents the context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            SharedTestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// A method used to clean up the test environment.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            SharedTestSuiteBase.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify that Versioning sub-request failed with empty url.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_TestCase_S16_TC01_Versioning_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { versioningSubRequest });
            SubResponseType subResponse = cellStoreageResponse.ResponseCollection.Response[0].SubResponse[0];
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11151
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    errorCode,
                    "MS-FSSHTTP",
                    11151,
                    @"[In VersioningSubResponseType] In the case of failure, the ErrorCode attribute that is part of a SubResponse element specifies the error code result for this subrequest.");
            }
            else
            {
                Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    errorCode,
                    "Error should occur if call versioning request with empty url.");
            }
        }

        /// <summary>
        /// A method used to verify that error FileNotExistsOrCannotBeCreated should be returned if the protocol server was unable to find the URL for the file.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_TestCase_S16_TC02_Versioning_FileNotExistsOrCannotBeCreated()
        {
            string fileUrlNotExit = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(fileUrlNotExit, new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11246
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotExistsOrCannotBeCreated,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    11246,
                    @"[In Versioning Subrequest] [The protocol returns results based on the following conditions:]If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotExistsOrCannotBeCreated,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(versioningSubResponse.ErrorCode, this.Site),
                    "Error FileNotExistsOrCannotBeCreated should be returned if the protocol server was unable to find the URL for the file.");
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