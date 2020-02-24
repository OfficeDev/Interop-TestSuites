namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with LockStatus operation.
    /// </summary>
    [TestClass]
    public abstract class S19_LockStatus : SharedTestSuiteBase
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
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 246601, this.Site), "This test case only runs when FileOperation subrequest is supported.");
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "LockStatus" sub-request.

        /// <summary>
        /// A method used to verify that LockStatus sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S19_TC01_LockStatus_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    string.Format("Account {0} with client ID {1} and schema lock ID {2} should join the coauthoring session successfully.", this.UserName01, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID));
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            LockStatusSubRequestType lockStatus = SharedTestSuiteHelper.CreateLockStatusSubRequest();

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { lockStatus });
            LockStatusSubResponseType lockStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<LockStatusSubResponseType>(response, 0, 0, this.Site);
            SubResponseType subresponse = response.ResponseCollection.Response[0].SubResponse[0];

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture the requirement MS-FSSHTTP_R246601
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         246601,
                         @"[In Appendix B: Product Behavior] Implementation does support LockStatus operation. (SharePoint Server 2016 and above follow this behavior.)");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2272
                Site.CaptureRequirementIfAreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "MS-FSSHTTP",
                    2272,
                    @"[LockStatusSubResponseType]In the case of success, it contains information requested as part of a LockStatus subrequest. ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R401011
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    int.Parse(lockStatusResponse.SubResponseData.LockType),
                    "MS-FSSHTTP",
                    401011,
                    @"[In LockTypes] 1: The integer value ""1"", indicating a shared lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "In the case of success, it contains information requested as part of a LockStatus subrequest. ");
            }
        }
        #endregion
    }
}