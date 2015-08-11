namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with ServerTime operation.
    /// </summary>
    [TestClass]
    public abstract class S06_ServerTime : SharedTestSuiteBase
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
        /// A method used to clean up the test environment.
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
        public void S06_ServerTimeInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases for "ServerTime" sub-request.

        /// <summary>
        /// A method used to verify that ServerTime sub-request can be executed successfully, when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S06_TC01_ServerTime_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "ServerTime"sub-request with correct input parameters.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { serverTimeSubRequest });
            ServerTimeSubResponseType serverTimeSubResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(serverTimeSubResponse, "The object 'serverTimeSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(serverTimeSubResponse.ErrorCode, "The object 'serverTimeSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code in the sub-response equals "Success", then capture MS-FSSHTTP_R1342
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1342,
                         @"[In ServerTime Subrequest] The protocol returns results based on the following conditions: Otherwise [except: the processing of the ServerTime subrequest by the server fails to get the server time or encountered an unknown exception], the protocol server sets the error code value to ""Success"" to indicate success in processing the ServerTime subrequest.");

                bool isVerifyR737 = System.Convert.ToInt64(serverTimeSubResponse.SubResponseData.ServerTime) > 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R737, expect the serverTime larger than 0, the actual value is " + serverTimeSubResponse.SubResponseData.ServerTime);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R737
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR737,
                         "MS-FSSHTTP",
                         737,
                         @"[In ServerTimeSubResponseDataType] ServerTime: A positive integer that specifies the server time, which is expressed as a tick count.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(serverTimeSubResponse.ErrorCode, this.Site),
                    @"[In ServerTime Subrequest] The protocol returns results based on the following conditions: Otherwise [except: the processing of the ServerTime subrequest by the server fails to get the server time or encountered an unknown exception], the protocol server sets the error code value to ""Success"" to indicate success in processing the ServerTime subrequest.");

                bool isVerifyR737 = System.Convert.ToInt64(serverTimeSubResponse.SubResponseData.ServerTime) > 0;
                this.Site.Log.Add(
                  LogEntryKind.Debug,
                  "For MS-FSSHTTP_R737, expect the serverTime larger than 0, the actual value is " + serverTimeSubResponse.SubResponseData.ServerTime);

                Site.Assert.IsTrue(
                    isVerifyR737,
                    @"[In ServerTimeSubResponseDataType] ServerTime: A positive integer that specifies the server time, which is expressed as a tick count.");
            }
        }

        #endregion 
    }
}