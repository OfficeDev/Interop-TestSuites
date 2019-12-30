namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with Properties operation.
    /// </summary>
    [TestClass]
    public abstract class S20_Properties : SharedTestSuiteBase
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
        public void S19_LockStatusInitialization()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 111111, this.Site), "This test case only runs when FileOperation subrequest is supported.");
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "Properties" sub-request.

        /// <summary>
        /// A method used to verify that Properties sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S20_TC01_Properties_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            PropertiesSubRequestType properties = SharedTestSuiteHelper.CreatePropertiesSubRequest(SequenceNumberGenerator.GetCurrentToken(), PropertiesRequestTypes.PropertyEnumerate, null, this.Site);

            CellStorageResponse response = Adapter.CellStorageRequest(
     this.DefaultFileUrl,
     new SubRequestType[] { properties },
     "1", 2, 2, null, null, null, null, null, null, true);
            PropertiesSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<PropertiesSubResponseType>(response, 0, 0, this.Site);

            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    string.Format("Account {0} with client ID {1} and schema lock ID {2} should join the coauthoring session successfully.", this.UserName01, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID));
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            PropertiesSubRequestType propertiess = SharedTestSuiteHelper.CreatePropertiesSubRequest(SequenceNumberGenerator.GetCurrentToken(), PropertiesRequestTypes.PropertyEnumerate, null, this.Site);
            CellStorageResponse response1 = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { propertiess });
            PropertiesSubResponseType propertiesResponse = SharedTestSuiteHelper.ExtractSubResponse<PropertiesSubResponseType>(response1, 0, 0, this.Site);
            SubResponseType subresponse = response.ResponseCollection.Response[0].SubResponse[0];

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2301
                Site.CaptureRequirementIfAreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "MS-FSSHTTP",
                    2301,
                    @"[PropertiesSubResponseType]In the case of success, it contains information requested as part of a Properties subrequest. ");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "[PropertiesSubResponseType]In the case of success, it contains information requested as part of a Properties subrequest. ");
            }
        }
        #endregion
    }
}