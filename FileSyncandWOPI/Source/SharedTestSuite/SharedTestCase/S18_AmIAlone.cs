namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with AmIAlone operation.
    /// </summary>
    [TestClass]
    public abstract class S18_AmIAlone : SharedTestSuiteBase
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
        public void S18_AmIAloneInitialization()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 246401, this.Site), "This test case only runs when AmIAlone subrequest is supported.");
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "AmIAlone" sub-request.

        /// <summary>
        /// A method used to verify that AmIAlone sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S18_TC01_AmIAlone_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            // Join a Coauthoring session with time out value 3600.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 3600);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                            ErrorCodeType.Success,
                            SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                            "Test case cannot continue unless the user {0} using client id {1} and schema lock id {2} to join the coauthoring session succeed.",
                            this.UserName01,
                            SharedTestSuiteHelper.DefaultClientID,
                            SharedTestSuiteHelper.ReservedSchemaLockID);
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            string transitionId = joinResponse.SubResponseData.TransitionID;

            AmIAloneSubRequestType amIAlone = SharedTestSuiteHelper.CreateAmIAloneSubRequest();
            amIAlone.SubRequestData.TransitionID = transitionId;

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { amIAlone });
            AmIAloneSubResponseType amIAloneResponse = SharedTestSuiteHelper.ExtractSubResponse<AmIAloneSubResponseType>(response, 0, 0, this.Site);
            SubResponseType subresponse = response.ResponseCollection.Response[0].SubResponse[0];

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture the requirement MS-FSSHTTP_R246401
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         246401,
                         @"[In Appendix B: Product Behavior] Implementation does support AmIAlone operation. <60> (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R224911
                Site.CaptureRequirementIfAreEqual<string>(
                    "True",
                    amIAloneResponse.SubResponseData.AmIAlone,
                    "MS-FSSHTTP",
                    224911,
                    @"[In AmIAloneSubResponseDataType]AmIAlone: True means the user is alone in the coauthoring session.");

                //Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2181011
                //If MS-FSSHTTP224911 is verified, this requirement can be verified directly
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    2181011,
                    @"[In SubResponseDataOptionalAttributes]AmIAlone: True means the user is alone in the coauthoring session.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2251
                Site.CaptureRequirementIfAreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "MS-FSSHTTP",
                    2251,
                    @"[In AmIAloneSubResponseType]In the case of success, it contains information requested as part of an AmIAlone subrequest. ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2374
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subresponse.ErrorCode, this.Site),
                    "MS-FSSHTTP",
                    2374,
                    @"[AmIAlone Subrequest][The protocol returns results based on the following conditions]Otherwise, the protocol server sets the error code value to ""Success"" to indicate success in processing the AmIAlone subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    GenericErrorCodeTypes.Success.ToString(),
                    subresponse.ErrorCode,
                    "In the case of success, it contains information requested as part of an AmIAlone subrequest. ");
            }
        }
        #endregion
    }
}