namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with GetVersions operation.
    /// </summary>
    [TestClass]
    public abstract class S16_Versioning : SharedTestSuiteBase
    {
        #region Test Suite Initialization and clean up

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
        /// A method used to clean up this class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            SharedTestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test Case Initialization

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S16_VersioningInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases for "Versioning" sub-request.

        /// <summary>
        /// A method used to verify that Versioning sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC01_Versioning_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getVersionsSubRequest });
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { versioningSubRequest });
            VersioningSubResponseType versioningSubResponse = SharedTestSuiteHelper.ExtractSubResponse<VersioningSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(versioningSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(versioningSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

        }

        #endregion
    }
}