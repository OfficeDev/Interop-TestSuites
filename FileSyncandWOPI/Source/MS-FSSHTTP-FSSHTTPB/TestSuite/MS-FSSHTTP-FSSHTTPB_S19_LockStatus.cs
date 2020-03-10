namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the LockStatus sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S19_LockStatus : S19_LockStatus
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
        /// A method used to verify that the Locktype is an exclusive lock in LockStatus sub-request.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_TestCase_S19_TC01_LockStatus_ExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the Coauthoring Feature
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Waiting change takes effect
            System.Threading.Thread.Sleep(30 * 1000);

            // Join a Coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, true, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    string.Format("Account {0} with client ID {1} and schema lock ID {2} should join the coauthoring session successfully.", this.UserName01, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultExclusiveLockID));
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultExclusiveLockID);

            LockStatusSubRequestType lockStatus = SharedTestSuiteHelper.CreateLockStatusSubRequest();

            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { lockStatus });
            LockStatusSubResponseType lockStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<LockStatusSubResponseType>(response, 0, 0, this.Site);
            SubResponseType subresponse = response.ResponseCollection.Response[0].SubResponse[0];

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R404011
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    int.Parse(lockStatusResponse.SubResponseData.LockType),
                    "MS-FSSHTTP",
                    404011,
                    @"[In LockTypes] 2: The integer value ""2"", indicating an exclusive lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<int>(
                    2,
                    int.Parse(lockStatusResponse.SubResponseData.LockType),
                    @"2: The integer value ""2"", indicating an exclusive lock on the file.");
            }
        }

        /// <summary>
        /// A method used to verify that LockStatus sub-request failed.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_TestCase_S19_TC02_LockStatus_Error()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            LockStatusSubRequestType lockStatus = SharedTestSuiteHelper.CreateLockStatusSubRequest();

            CellStorageResponse response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { lockStatus });

            if (Common.IsRequirementEnabled(2273, this.Site))
            {
                LockStatusSubResponseType lockStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<LockStatusSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    //Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2273
                    Site.CaptureRequirementIfAreNotEqual<string>(
                        "Success",
                        lockStatusResponse.ErrorCode,
                        "MS-FSSHTTP",
                        2273,
                        @"[LockStatusSubResponseType]In the case of failure, the ErrorCode attribute that is part of a SubResponse element specifies the error code result for this subrequest. ");
                }
                else
                {
                    Site.Assert.AreNotEqual<string>(
                        "Success",
                        lockStatusResponse.ErrorCode,
                        "In the case of failure, the ErrorCode attribute that is part of a SubResponse element specifies the error code result for this subrequest. ");
                }
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