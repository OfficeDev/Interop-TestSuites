namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with WhoAmI operation.
    /// </summary>
    [TestClass]
    public abstract class S05_WhoAmI : SharedTestSuiteBase
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
        public void S05_WhoAmIInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
        }

        #endregion

        #region Test Cases for "WhoAmI" sub-request.

        /// <summary>
        /// A method used to verify that WhoAmI sub-request can be executed successfully when all the input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S05_TC01_WhoAmI_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "WhoAmI"sub-request with correct input parameters.
            WhoAmISubRequestType whoAmISubRequest = SharedTestSuiteHelper.CreateWhoAmISubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { whoAmISubRequest });
            WhoAmISubResponseType whoAmISubResponse = SharedTestSuiteHelper.ExtractSubResponse<WhoAmISubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(whoAmISubResponse, "The object 'whoAmISubResponse' should not be null.");
            this.Site.Assert.IsNotNull(whoAmISubResponse.ErrorCode, "The object 'whoAmISubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code in the sub-response equals "Success", then capture MS-FSSHTTP_R1327, MS-FSSHTTP_R1434.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(whoAmISubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1327,
                         @"[In WhoAmI Subrequest] The protocol server returns results based on the following conditions: Otherwise[except: the processing of the WhoAmI subrequest by the protocol server failed to get the client-specific user information or encountered an unknown exception], the protocol server sets the error code value to ""Success"" to indicate success in processing the WhoAmI subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1434
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(whoAmISubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1434,
                         @"[In WhoAmISubResponseType] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the WhoAmI subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R7681
                Site.CaptureRequirementIfIsNotNull(
                         whoAmISubResponse.SubResponseData,
                         "MS-FSSHTTP",
                         7681,
                         @"[In WhoAmISubResponseType] As part of processing the WhoAmI subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true:
                         The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                bool isVerifyR904 = whoAmISubResponse.SubResponseData.UserName.IndexOf(this.UserName01, System.StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R904, expect the UserName attribute contains the value {0}, the actual UserName attribute value is {1}",
                    this.UserName01,
                    whoAmISubResponse.SubResponseData.UserName);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R904
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR904,
                         "MS-FSSHTTP",
                         904,
                         @"[In WhoAmISubResponseDataOptionalAttributes] UserName: [A UserNameType] that specifies the user name for the client.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(whoAmISubResponse.ErrorCode, this.Site),
                    @"[In WhoAmI Subrequest] The protocol server returns results based on the following conditions: Otherwise[except: the processing of the WhoAmI subrequest by the protocol server failed to get the client-specific user information or encountered an unknown exception], the protocol server sets the error code value to ""Success"" to indicate success in processing the WhoAmI subrequest.");

                Site.Assert.IsNotNull(
                    whoAmISubResponse.SubResponseData,
                    @"[In WhoAmISubResponseType] As part of processing the WhoAmI subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true:
                        The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                bool isVerifyR904 = whoAmISubResponse.SubResponseData.UserName.IndexOf(this.UserName01, System.StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Assert.IsTrue(
                    isVerifyR904,
                    @"[In WhoAmISubResponseDataOptionalAttributes] UserName: [A UserNameType] that specifies the user name for the client.");
            }
        }

        #endregion 
    }
}