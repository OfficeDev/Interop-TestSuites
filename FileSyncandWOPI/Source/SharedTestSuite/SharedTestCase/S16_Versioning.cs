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

        DateTime time = DateTime.Now.AddDays(-1);

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S16_VersioningInitialization()
        {
            this.DefaultFileUrl = this.PrepareFile();
            time = DateTime.UtcNow;
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

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11146
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.UserTable,
                    "MS-FSSHTTP",
                    11146,
                    @"[In VersioningSubResponseDataType] The UserTable element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11148
                Site.CaptureRequirementIfIsNotNull(
                    versioningSubResponse.SubResponseData.Versions,
                    "MS-FSSHTTP",
                    11148,
                    @"[In VersioningSubResponseDataType] The Versions element MUST be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11150
                // This requirement can be captured directly after capturing MS-FSSHTTP_R11146 and MS-FSSHTTP_R11148
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11150,
                    @"[In VersioningSubResponseType] In the case of success, it contains information requested as part of a versioning subrequest.");

                string expectLastModifiedTime =
                    ((long)(time - new DateTime(1601, 1, 1, 0, 0, 0)).TotalSeconds * 10000000).ToString();

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11179
                Site.CaptureRequirementIfAreEqual<string>(
                    expectLastModifiedTime,
                    versioningSubResponse.SubResponseData.Versions.Version[0].LastModifiedTime,
                    "MS-FSSHTTP",
                    11179,
                    @"[In FileVersionDataType] LastModifiedTime specifies the number of 100-nanosecond intervals that have elapsed since 00:00:00 on January 1, 1601, which MUST be Coordinated Universal Time (UTC).");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11178
                // Calculate expected last modified time by multiple the time span with ten million, this requirement can be captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    expectLastModifiedTime,
                    versioningSubResponse.SubResponseData.Versions.Version[0].LastModifiedTime,
                    "MS-FSSHTTP",
                    11178,
                    @"[In FileVersionDataType] A single tick represents 100 nanoseconds, or one ten-millionth of a second.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11180
                Site.CaptureRequirementIfAreEqual<string>(
                    versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                    versioningSubResponse.SubResponseData.Versions.Version[0].UserId,
                    "MS-FSSHTTP",
                    11180,
                    @"[In FileVersionDataType] UserId: An integer that specifies the user that last modified the version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11181
                Site.CaptureRequirementIfAreEqual<string>(
                    versioningSubResponse.SubResponseData.UserTable.User[0].UserId,
                    versioningSubResponse.SubResponseData.Versions.Version[0].UserId,
                    "MS-FSSHTTP",
                    11181,
                    @"[In FileVersionDataType] The number MUST match the UserId attribute of a UserDataType (section 2.3.1.42) described in the VersioningUserTableType in the current VersioningSubResponseDataType.");
            }
            else
            {
                Site.Assert.IsNotNull(
                    versioningSubResponse.SubResponseData.UserTable,
                    @"The UserTable element must be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");

                Site.Assert.IsNotNull(
                    versioningSubResponse.SubResponseData.Versions,
                    @"The UserTable element must be included in the response if the SubResponseType of the parent VersioningSubResponseType is of type ""GetVersionList.""");
            }
        }

        /// <summary>
        /// A method used to verify that Versioning sub-request failed with empty url.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S16_TC02_Versioning_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            VersioningSubRequestType versioningSubRequest = SharedTestSuiteHelper.CreateVersioningSubRequest(SequenceNumberGenerator.GetCurrentToken(), VersioningRequestTypes.GetVersionList, null, this.Site);
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { versioningSubRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11151
                Site.CaptureRequirementIfAreNotEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.Success,
                    cellStoreageResponse.ResponseVersion.ErrorCode,
                    "MS-FSSHTTP",
                    11151,
                    @"[In VersioningSubResponseType] In the case of failure, the ErrorCode attribute that is part of a SubResponse element specifies the error code result for this subrequest.");
            }
            else
            {
                Site.Assert.AreNotEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.Success,
                    cellStoreageResponse.ResponseVersion.ErrorCode,
                    "Error should occur if call versioning request with empty url.");
            }
        }

        #endregion
    }
}