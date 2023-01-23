namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using System.Net;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with Coauthoring operation.
    /// </summary>
    [TestClass]
    public abstract class S02_Coauth : SharedTestSuiteBase
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

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S02_CoauthInitialization()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93101, this.Site))
            {
                this.Site.Assume.Inconclusive("Implementation does not support the coauthoring subrequest.");
            }

            // Initialize the default file URL, for this scenario, the target file URL should be unique for each test case
            this.DefaultFileUrl = this.PrepareFile();
        }

        #region Test Cases

        #region Join Coauthoring session

        /// <summary>
        /// A method used to verify that only the clients which have the same schema lock identifier can lock the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC01_JoinCoauthSession_SameSchemaLockID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join Coauthoring session using the first user
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    string.Format("Account {0} with client ID {1} and schema lock ID {2} should join the coauthoring session successfully.", this.UserName01, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID));
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join the Coauthoring session using the second user with same SchemaLockId
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string secondClientId = System.Guid.NewGuid().ToString();
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType secondResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1490
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(secondResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1490,
                         @"[In CoauthSubRequestDataType][SchemaLockID] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(secondResponse.ErrorCode, this.Site),
                    "After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }

            // Join the Coauthoring session using the third user with a different SchemaLockId
            this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);
            string thirdClientId = System.Guid.NewGuid().ToString();
            string newSchemaId = System.Guid.NewGuid().ToString();
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(thirdClientId, newSchemaId);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType thirdResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R581
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(thirdResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         581,
                         @"[In CoauthSubRequestDataType][SchemaLockID] The protocol server ensures that at any instant of time, only clients having the same schema lock identifier can lock the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R579
                // If the client fails in joining the Coauthoring session with a different SchemaLockId, then MS-FSSHTTP_R579 could be captured.
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(thirdResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         579,
                         @"[In CoauthSubRequestDataType][SchemaLockID] The schema lock identifier is used by the protocol server to block other clients with different schema identifiers.");

                // If this operation succeeds, it can approve that the implementation can support the coauthoring feature.
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         93101,
                         @"[In Appendix B: Product Behavior] Implementation does support the coauthoring subrequest. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.AreNotEqual<ErrorCodeType>(
                   ErrorCodeType.Success,
                   SharedTestSuiteHelper.ConvertToErrorCodeType(thirdResponse.ErrorCode, this.Site),
                   "The protocol server ensures that at any instant of time, only clients having the same schema lock identifier can lock the file.");
            }
        }

        /// <summary>
        /// A method used to verify that a coauthoring shared lock is allowed to fall back to an exclusive lock when AllowFallbackToExclusive attribute value in the JoinCoauthoringSession subRequest is set to true.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC02_JoinCoauthoringSession_CoauthoringDisabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the Coauthoring Feature
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Waiting change takes effect
            System.Threading.Thread.Sleep(30 * 1000);

            // Join a Coauthoring session with AllowFallbackToExclusive attribute set to true
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, true, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType firstJoinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(firstJoinResponse.ErrorCode, this.Site), "The client should join the coauthoring session successfully.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            this.CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(firstJoinResponse);
            this.CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(firstJoinResponse);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled(11274, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11274
                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.RequestNotSupported,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(firstJoinResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             11274,
                             @"[In Appendix B: Product Behavior] [The protocol server MUST follow the following common processing rules for all types of subrequests] The implementation does not return an error code value set to ""RequestNotSupported"" for a cell storage service subrequest if the following conditions are all true: 

                             The protocol client sent a coauthoring subrequest;
                             The protocol server supports shared locking with tracking of the coauthoring transition;
                             The coauthoring administrator setting for the server is turned off. (Microsoft SharePoint Foundation 2010 / Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1019
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         firstJoinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         1019,
                         @"[In Join Coauthoring Session][If the coauthoring feature is disabled on the protocol server, it does one of the following:] If the AllowFallbackToExclusive attribute is set to true, the protocol server gets an exclusive lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R442
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         firstJoinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         442,
                         @"[In SubRequestDataOptionalAttributes] When shared locking on the file is not supported: An AllowFallbackToExclusive attribute value set to true indicates that a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is allowed to fall back to an exclusive lock subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R441
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         441,
                         @"[In SubRequestDataOptionalAttributes] AllowFallbackToExclusive: A Boolean value that specifies to a protocol server whether a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is allowed to fall back to an exclusive lock subrequest when shared locking on the file is not supported.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R404
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         firstJoinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         404,
                         @"[In LockTypes] ExclusiveLock: The string value ""ExclusiveLock"", indicating an exclusive lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R406
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         firstJoinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         406,
                         @"[In LockTypes][ExclusiveLock or 2] In a cell storage service response message, an exclusive lock indicates that an exclusive lock is granted to the current client for that specific file.");

                this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "For requirement MS-FSSHTTP_R482 and MS-FSSHTTP_R617, the ExclusiveLockReturnReason should be included in the response, but actually the attribute value is {0}",
                        firstJoinResponse.SubResponseData.ExclusiveLockReturnReasonSpecified ? "exist" : "not exist");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R482
                Site.CaptureRequirementIfIsTrue(
                         firstJoinResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         482,
                         @"[In SubResponseDataOptionalAttributes][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations when the LockType attribute in the subresponse is set to ""ExclusiveLock"": A coauthoring subrequest of type ""Join coauthoring session"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R617
                Site.CaptureRequirementIfIsTrue(
                         firstJoinResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         617,
                         @"[In CoauthSubResponseDataType][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a coauthoring subresponse that is generated in response to the JoinCoauthoring type of coauthoring subrequest when the LockType attribute in the subresponse is set to ""ExclusiveLock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R349
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CoauthoringDisabled,
                         firstJoinResponse.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         349,
                         @"[In ExclusiveLockReturnReasonTypes] CoauthoringDisabled: The string value ""CoauthoringDisabled"", indicating that an exclusive lock is granted on a file because coauthoring is disabled.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R568 and MS-FSSHTTP_R1019
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CoauthoringDisabled,
                         firstJoinResponse.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         568,
                         @"[In CoauthSubRequestDataType][AllowFallbackToExclusive] When shared locking on the file is not supported:
                         An AllowFallbackToExclusive attribute value set to true indicates that a coauthoring subrequest of type ""Join coauthoring session"" is allowed to fall back to an exclusive lock subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                         "ExclusiveLock",
                         firstJoinResponse.SubResponseData.LockType,
                         "When the coauthoring feature is disabled on the protocol, if the AllowFallbackToExclusive attribute is set to true, the protocol server should get an exclusive lock on the file.");

                Site.Assert.IsTrue(
                    firstJoinResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                    "When the lock type is set to ExclusiveLock, the ExclusiveLockReturnReason attribute should be specified.");

                Site.Assert.AreEqual<ExclusiveLockReturnReasonTypes>(
                    ExclusiveLockReturnReasonTypes.CoauthoringDisabled,
                    firstJoinResponse.SubResponseData.ExclusiveLockReturnReason,
                    "When the coauthoring feature is disabled on the protocol, if the AllowFallbackToExclusive attribute is set to true, the ExclusiveLockReturnReasonTypes should be set to CoauthoringDisabled.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for Join Coauthoring Session subRequest if there is a current exclusive lock on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC03_JoinCoauthoringSession_FileAlreadyLockedOnServer_ExclusiveLock()
        {
            // Get the exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Join the Coauthoring session using the second user
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1553
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1553,
                         @"[In Join Coauthoring Session] If there is a current exclusive lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If there is a current exclusive lock on the file by current user and the other user requests an exclusive lock on the same file, the protocol server should return an error code value set to ""FileAlreadyLockedOnServer""");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for Join Coauthoring Session subRequest if there is a current shared lock on the file with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC04_JoinCoauthoringSession_FileAlreadyLockedOnServer_DifferentSchemaLockID()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join the Coauthoring session with a different schema lock identifier
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, System.Guid.NewGuid().ToString());
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1554
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1554,
                         @"[In Join Coauthoring Session] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the client can join the coauthoring session successfully with correct parameters.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC05_JoinCoauthoringSession_Success()
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

            this.CaptureSucceedCoauthSubRequest(joinResponse);
            this.CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);
            this.CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                bool isVerifyR614 = !string.IsNullOrEmpty(joinResponse.SubResponseData.TransitionID);
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    @"For MS-FSSHTTP_R614, the transition identifier should be returned for the coauthoring subrequest of type ""Join coauthoring session"", but actually the attribute value is {0}",
                    joinResponse.SubResponseData.TransitionID);

                Site.CaptureRequirementIfIsTrue(
                         isVerifyR614,
                         "MS-FSSHTTP",
                         614,
                         @"[In CoauthSubResponseDataType] The transition identifier MUST be returned by a coauthoring subrequest of type ""Join coauthoring session"".");

                // If the coauthoring status returned by the join coauthoring session operation is Alone, then indicates the server actually check and gets the coauthoring status of the file. 
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         joinResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1016,
                         @"[In Join Coauthoring Session] The protocol server also checks and gets the coauthoring status of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R316
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         joinResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         316,
                         @"[In CoauthStatusType] Alone [means]: A string value of ""Alone"", indicating a coauthoring status of alone.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R996
                Site.CaptureRequirementIfAreEqual<string>(
                          "SchemaLock",
                         joinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         996,
                         @"[In Coauth Subrequest] If the coauthoring subrequest is of type ""Join coauthoring session"", the protocol server MUST return the lock type granted to the protocol client as part of the response message to the protocol client—if  the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R401
                Site.CaptureRequirementIfAreEqual<string>(
                         "SchemaLock",
                         joinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         401,
                         @"[In LockTypes] SchemaLock: The string value ""SchemaLock"", indicating a shared lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R403
                Site.CaptureRequirementIfAreEqual<string>(
                         "SchemaLock",
                         joinResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         403,
                         @"[In LockTypes][SchemaLock or 1] In a cell storage service response message, a shared lock indicates that the current client is granted a shared lock on the file, which allows for coauthoring the file along with other clients.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Alone,
                    joinResponse.SubResponseData.CoauthStatus,
                    "When only one client joins the coauth session, the server should return the CoauthStatusType as Alone.");

                Site.Assert.AreEqual<string>(
                    "SchemaLock",
                    joinResponse.SubResponseData.LockType,
                    "When one client joins the coauth session, the server should return the LockType as SchemaLock");
            }

            if (!StatusManager.RollbackStatus())
            {
                this.Site.Assert.Fail("Release the shared lock {0} fails", SharedTestSuiteHelper.ReservedSchemaLockID);
            }

            // Exit the coauthoring session not means immediately release the schema lock
            // Here it will be retried several times to wait the lock is actually released.
            bool isLockExist = false;
            int waitTime = 5;
            do
            {
                if (waitTime == 0)
                {
                    this.Site.Assert.Fail("After wait 10 seconds for releasing the schema lock, but the lock still exists.");
                }

                // Sleep 2 seconds once
                SharedTestSuiteHelper.Sleep(2);
                isLockExist = this.CheckSchemaLockExist(this.DefaultFileUrl, SharedTestSuiteHelper.ReservedSchemaLockID);
                waitTime--;
            }
            while (isLockExist == true);

            // Join a Coauthoring session using different schema lock id again.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, System.Guid.NewGuid().ToString(), null, null, 3600);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                            ErrorCodeType.Success,
                            SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                            "Test case cannot continue unless the user {0} using client id {1} and schema lock id {2} to join the coauthoring session succeed.",
                            this.UserName01,
                            subRequest.SubRequestData.ClientID,
                            subRequest.SubRequestData.SchemaLockID);
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R582
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         582,
                         @"[In CoauthSubRequestDataType][SchemaLockID] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    "After all the protocol clients have released their lock for that file, the server should allow client with a different schema lock identifier to get a shared lock for that file.");
            }
        }

        /// <summary>
        /// A method used to verify the response information for Join Coauthoring subRequest when the file already has a shared lock on the protocol server with the given schema lock identifier, and the client has already joined the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC06_JoinCoauthoringSession_SameClientID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 60);
            subRequest.SubRequestData.Timeout = "120";
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The operation CoauthingSubRequest should succeed.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join the Coauthoring session again, refresh timeout value to 3600.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 3600);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1018
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1018,
                         @"[In Join Coauthoring Session][If the file already has a shared lock on the server with the given schema lock identifier, and the client has already joined the coauthoring session, the protocol server does both of the following:] 
                         Returns an error code value set to ""Success"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    "If the file already has a shared lock on the server with the given schema lock identifier, and the client has already joined the coauthoring session, the protocol server should return Success.");
            }
        }

        /// <summary>
        /// A method used to verify that the CoauthStatus is set to "Coauthoring" by the protocol server if the current client is the second, third, or later coauthor joining the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC07_JoinCoauthoringSession_TwoOrThirdClients()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Use SUT method to set max number of coauthors to 4.
            bool isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(4);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");

            // Join a Coauthoring session using the first client.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                ErrorCodeType.Success, 
                SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), 
                "The first client should join the coauthoring session successfully.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            this.CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);
            this.CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);

            // Join the Coauthoring session using the second client.
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string secondClientId = System.Guid.NewGuid().ToString();
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The second client should join the coauthoring session successfully.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            this.CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);
            this.CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1025
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         joinResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1025,
                         @"[In Join Coauthoring Session] If the current client is the second coauthor joining the coauthoring session, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                   CoauthStatusType.Coauthoring,
                   joinResponse.SubResponseData.CoauthStatus,
                   "If the current client is the second coauthor joining the coauthoring session, the protocol server returns the CoauthStatusType as Coauthoring.");
            }

            // Join the Coauthoring session using the third client.
            this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);
            string thirdClientId = System.Guid.NewGuid().ToString();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            while (retryCount > 0)
            {
                subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(thirdClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
                cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site) == ErrorCodeType.Success
                    && joinResponse.SubResponseData.CoauthStatus == CoauthStatusType.Coauthoring)
                {
                    break;
                }

                retryCount--;
                if (retryCount == 0)
                {
                    Site.Assert.Fail("Join the coauthoring session should be succeed if the current client is the third coauthor.");
                }

                System.Threading.Thread.Sleep(waitTime);
            }

            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, thirdClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName03, this.Password03, this.Domain);

            this.CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);
            this.CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(joinResponse);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement MS-FSSHTTP_R1877
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         joinResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1877,
                         @"[In Join Coauthoring Session] If the current client is the third coauthor joining the coauthoring session, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1908
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         joinResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1908,
                         @"[In Get Coauthoring Session] If the current client is the third client trying to edit the document, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                  CoauthStatusType.Coauthoring,
                  joinResponse.SubResponseData.CoauthStatus,
                  "If the current client is the third coauthor joining the coauthoring session, the protocol server should return the CoauthStatusType as Coauthoring.");
            }

            isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");
        }

        /// <summary>
        /// A method used to verify the protocol server does not return CoauthStatus attribute when client tries to join the coauthoring session and the subRequest falls back to an exclusive lock subRequest.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC08_JoinCoauthoringSession_AllowFallbackToExclusiveIsTrue()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the coauthoring feature
            bool isSwitchSuccessful = this.SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchSuccessful, "The coauthoring feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Create a JoinCoauthoringSession subRequest with AllowFallbackToExclusive set to true.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, true);

            // Send this subRequest to the protocol server, expect the protocol server not to return CoauthStatus attribute.
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType subReponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subReponse.ErrorCode, this.Site), "The coauthoring subRequest should be falls back to exclusive lock successfully.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3064, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "For the requirement MS-FSSHTTP_R3064, the coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3064
                    Site.CaptureRequirementIfIsFalse(
                             subReponse.SubResponseData.CoauthStatusSpecified,
                             "MS-FSSHTTP",
                             3064,
                             @"[In Appendix B: Product Behavior] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the implementation does not return CoauthStatus attribute when client tries to join the coauthoring session and the subrequest falls back to an exclusive lock subrequest. <23> Section 2.2.8.2:  SharePoint Server 2010 will not return CoauthStatus attribute when client tries to join the coauthoring session and the subrequest falls back to an exclusive lock subrequest.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3070, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "For the requirement MS-FSSHTTP_R3070, the coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3070
                    Site.CaptureRequirementIfIsFalse(
                             subReponse.SubResponseData.CoauthStatusSpecified,
                             "MS-FSSHTTP",
                             3070,
                             @"[In Appendix B: Product Behavior] The implementation does not return CoauthStatus attribute and falls back the subrequest to an exclusive lock subrequest. <27> Section 2.3.1.7:  SharePoint Server 2010 will not return CoauthStatus attribute when client tries to join the coauthoring session and the subrequest falls back to an exclusive lock subrequest.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 609, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        @"The implementation does return CoauthStatus attribute when sending a Join coauth subRequest successfully and the subRequest isn't fallen back to an exclusive lock, actual the attribute {0}",
                        subReponse.SubResponseData.LockTypeSpecified ? "exists" : "does not exist");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R609
                    Site.CaptureRequirementIfIsTrue(
                             subReponse.SubResponseData.CoauthStatusSpecified,
                             "MS-FSSHTTP",
                             609,
                             @"[In Appendix B: Product Behavior] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the implementation does specify the CoauthStatus in a coauthoring subresponse that is generated in response to Join coauthoring session. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 474, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "For the requirement MS-FSSHTTP_R474, the coauth status should be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R474
                    Site.CaptureRequirementIfIsTrue(
                             subReponse.SubResponseData.CoauthStatusSpecified,
                             "MS-FSSHTTP",
                             474,
                             @"[In Appendix B: Product Behavior] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the implementation does specify CoauthStatus attribute in a subresponse that is generated in response to a coauthoring subrequest of type ""Join coauthoring session"". (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3064, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    Site.Assert.IsFalse(
                        subReponse.SubResponseData.CoauthStatusSpecified,
                        "The coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3070, this.Site))
                {
                    Site.Log.Add(
                       LogEntryKind.Debug,
                       "The coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    Site.Assert.IsFalse(
                      subReponse.SubResponseData.CoauthStatusSpecified,
                      "The coauth status should not be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2010/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 609, this.Site))
                {
                    Site.Log.Add(
                      LogEntryKind.Debug,
                      "The coauth status should be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and lock server responds Success. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior)");

                    Site.Assert.IsTrue(
                      subReponse.SubResponseData.CoauthStatusSpecified,
                      "The coauth status should be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and lock server responds Success. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 474, this.Site))
                {
                    Site.Log.Add(
                      LogEntryKind.Debug,
                      "The coauth status should be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");

                    Site.Assert.IsTrue(
                      subReponse.SubResponseData.CoauthStatusSpecified,
                      "The coauth status should be returned when operation is \"Join Coauthoring Session\" and falls back to exclusive and server responds Success (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the specified attributes CoauthRequestType are not provided in the sub request.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC09_JoinCoauthoringSession_NoCoauthRequestType()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            // Create a coauthoring subRequet without specifying CoauthRequestType attribute.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            subRequest.SubRequestData.CoauthRequestTypeSpecified = false;

            // Send the subRequet to the protocol server, expect the protocol server returns error code "InvalidArgument" or "HighLevelExceptionThrown"
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3065
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3065, this.Site))
                {
                    CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
                    ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidArgument,
                             errorCode,
                             "MS-FSSHTTP",
                             3065,
                             @"[In Appendix B: Product Behavior]  If the specified attributes[CoauthRequestType] are not provided, the implementation does return error code. &lt;25&gt; Section 2.3.1.5:  SharePoint Server 2010 returns an ""InvalidArgument"" error code as part of the SubResponseData element associated with the coauthoring subresponse(Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior).");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3066
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3066, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.HighLevelExceptionThrown,
                             response.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             3066,
                             @"[In Appendix B: Product Behavior]  If the specified attributes[CoauthRequestType] are not provided, the implementation does return error code. &lt;25&gt; Section 2.3.1.5:  SharePoint Server 2013 and SharePoint Server 2016, return ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the coauthoring subresponse(Microsoft Sharepoint 2013/Microsoft Office 2016/Microsoft Sharepoint Server 2016).");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3025, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.HighLevelExceptionThrown,
                             response.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             3025,
                             @"[In Appendix B: Product Behavior] Implementation does return the value ""HighLevelExceptionThrown"" of GenericErrorCodeTypes when any undefined error that occurs during the processing of the cell storage service request. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }
            }
            else
            {
                if (response != null)
                {
                    if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3065, this.Site))
                    {
                        CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
                        ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

                        Site.Assert.AreEqual<ErrorCodeType>(
                            ErrorCodeType.InvalidArgument,
                            errorCode,
                            @"If the specified attributes[CoauthRequestType] are not provided, the implementation does return ""InvalidArgument"" error code as part of the SubResponseData element associated with the coauthoring subresponse. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                    }

                    if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3025, this.Site))
                    {
                        Site.Assert.AreEqual<GenericErrorCodeTypes>(
                            GenericErrorCodeTypes.HighLevelExceptionThrown,
                            response.ResponseVersion.ErrorCode,
                            @"[In Appendix B: Product Behavior] Implementation does return the value ""HighLevelExceptionThrown"" of GenericErrorCodeTypes when any undefined error that occurs during the processing of the cell storage service request. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                    }
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to "NumberOfCoauthorsReachedMax" when there are multiple clients in the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC10_GetCoauthoringSession_NumberOfCoauthorsReachedMax()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Use SUT method to set the max number of coauthors to 2.
            bool isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");

            // Join a Coauthoring session using the first user
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The first client should join the coauthoring session successfully.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join the Coauthoring session using the second user
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(System.Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The second user should join the coauthoring session successfully");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
            {
                // Join the Coauthoring session using the third user
                this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);

                int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
                int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

                while (retryCount > 0)
                {
                    subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(System.Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID);
                    cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                    response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

                    if (SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site) == ErrorCodeType.NumberOfCoauthorsReachedMax)
                    {
                        break;
                    }

                    retryCount--;
                    if (retryCount == 0)
                    {
                        Site.Assert.Fail("NumberOfCoauthorsReachedMax error should be returned if the maximum number of coauthorable clients allowed to join a coauthoring session to edit a coauthorable file has been reached.");
                    }

                    System.Threading.Thread.Sleep(waitTime);
                }

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R933
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.NumberOfCoauthorsReachedMax,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             933,
                             @"[In Common Message Processing Rules and Events] The coauthoring transition allows for the number of users editing the file to increase from 1 to n or to decrease from n to 1, where n is the maximum number of users who are allowed to edit a single file at an instant in time.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1027
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.NumberOfCoauthorsReachedMax,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1027,
                             @"[In Join Coauthoring Session] The protocol server returns an error code value set to ""NumberOfCoauthorsReachedMax"" when all of the following conditions are true: 
                         1.The maximum number of coauthorable clients allowed to join a coauthoring session to edit a coauthorable file has been reached;
                         2.The current client is not allowed to edit the file because the limit has been reached.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R392
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.NumberOfCoauthorsReachedMax,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             392,
                             @"[In LockAndCoauthRelatedErrorCodeTypes] NumberOfCoauthorsReachedMax indicates an error when the number of users that coauthor a file has reached the threshold limit.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R393
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.NumberOfCoauthorsReachedMax,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             393,
                             @"[In LockAndCoauthRelatedErrorCodeTypes][NumberOfCoauthorsReachedMax] The threshold limit specifies the maximum number of users allowed to coauthor a file at any instant in time.");
                }
                else
                {
                    this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.NumberOfCoauthorsReachedMax,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        "After set the max number of coauthUsers to 2, the protocol should not allow the third user to join the coauth session.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code when the request token is not specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC11_MissingRequestToken()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join Coauthoring session without the request token.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest }, null);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3011, this.Site))
            {
                SharedTestSuiteHelper.CheckCellStorageResponse(cellResponse, this.Site, 0);
                Response response = cellResponse.ResponseCollection.Response[0];
                this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "If the RequestToken attribute of the corresponding Request element is an empty string, the implementation does not return the ErrorCode attribute in ResponseVersion element, it actually {0}",
                    response.ErrorCodeSpecified ? "returns" : "does not return");

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsFalse(
                             response.ErrorCodeSpecified,
                             "MS-FSSHTTP",
                             3011,
                             @"[In Appendix B: Product Behavior] The implementation does not return the ErrorCode attribute in Response element. (SharePoint Foundation 2013 , SharePoint Server 2013 and above follow this behavior.)");
                }
                else
                {
                    Site.Assert.IsFalse(
                            response.ErrorCodeSpecified,
                            @"The implementation does not return the ErrorCode attribute in Response element. (SharePoint Foundation 2013 , SharePoint Server 2013 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify that a coauthoring shared lock isn't allowed to fall back to an exclusive lock when AllowFallbackToExclusive attribute value in the JoinCoauthoringSession subRequest is set to false.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC12_JoinCoauthoringSession_AllowFallbackToExclusiveIsFalse()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Disable the Coauthoring Feature
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            CoauthSubResponseType subResponse = null;
            System.Threading.Thread.Sleep(60 * 1000);
            while (retryCount > 0)
            {
                // Join a Coauthoring session with AllowFallbackToExclusive attribute set to false
                CoauthSubRequestType coauthRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, false, null);
                CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { coauthRequest });
                subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site) != ErrorCodeType.Success)
                {
                    break;
                }
                 
                retryCount--;
                if (retryCount == 0)
                {
                    Site.Assert.Fail("Error should be returned when shared locking on file is not supported and AllowFallbackToExclusive attribute value set to false.");
                }

                System.Threading.Thread.Sleep(waitTime);
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R569 and MS-FSSHTTP_R443
                // If the protocol server does not fall back the lock to an exclusive lock on the file, the client will fail joining the coauthoring session,
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         569,
                         @"[In CoauthSubRequestDataType][AllowFallbackToExclusive] When shared locking on the file is not supported:
                         An AllowFallbackToExclusive attribute value set to false indicates that a coauthoring subrequest is not allowed to fall back to an exclusive lock subrequest.");

                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         443,
                         @"[In SubRequestDataOptionalAttributes][When shared locking on the file is not supported:] An AllowFallbackToExclusive attribute value set to false indicates that a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is not allowed to fall back to an exclusive lock subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3104 and MS-FSSHTTP_R382
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3104, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3104,
                             @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<42> Section 3.1.4.3.1:  SharePoint Server 2010 will return an error code value ""FileNotLockedOnServerAsCoauthDisabled"", if the AllowFallbackToExclusive attribute is set to false.)");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             382,
                             @"[In LockAndCoauthRelatedErrorCodeTypes] FileNotLockedOnServerAsCoauthDisabled indicates an error when no shared lock exists on a file because coauthoring of the file is disabled on the server.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R310401 and MS-FSSHTTP_R383
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 310401, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             310401,
                             @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<42> Section 3.1.4.3.1:  SharePoint Server 2019 and SharePoint Server Subscription Edition will return an error code value ""LockNotConvertedAsCoauthDisabled"", if the AllowFallbackToExclusive attribute is set to false.)");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             383,
                             @"[In LockAndCoauthRelatedErrorCodeTypes] Indicates an error when a protocol server fails to process a lock conversion request sent as part of a cell storage service request because coauthoring of the file is disabled on the server.");
                }

                if (Common.IsRequirementEnabled(11274, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11274
                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.RequestNotSupported,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             11274,
                             @"[In Appendix B: Product Behavior] [The protocol server MUST follow the following common processing rules for all types of subrequests] The implementation does not return an error code value set to ""RequestNotSupported"" for a cell storage service subrequest if the following conditions are all true: 

                             The protocol client sent a coauthoring subrequest;
                             The protocol server supports shared locking with tracking of the coauthoring transition;
                             The coauthoring administrator setting for the server is turned off. (Microsoft SharePoint Foundation 2010 / Microsoft SharePoint Server 2010 and above follow this behavior.)");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3105
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3105, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3105,
                             @"[In Appendix B: Product Behavior]When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServer"". (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3104, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 310401, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""LockNotConvertedAsCoauthDisabled"". (SharePoint Server 2019 and SharePoint Server Subscription Edition follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3105, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServer,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServer"". (<27> Section 3.1.4.3.1:  SharePoint Foundation 2013 and SharePoint Server 2013 will return an error code value ""FileNotLockedOnServer"" if the AllowFallbackToExclusive attribute is set to false.)");
                }
            }
        }
        #endregion

        #region Exit Coauthoring session
        /// <summary>
        /// A method used to verify the client exits a coauthoring session successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC13_ExitCoauthoringSession_Success()
        {
            // Join the coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Exit the Coauthoring session using the same clientID and SchemaLockId.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                "User {0} exit the coauthoring session with client ID {1} and schema lock ID {2} on the file {3} should succeed.",
                this.UserName01,
                SharedTestSuiteHelper.DefaultClientID,
                SharedTestSuiteHelper.ReservedSchemaLockID,
                this.DefaultFileUrl);

            this.CaptureSucceedCoauthSubRequest(exitResponse);
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to "Success" for ExitCoauthoringSession subRequest when the coauthoring session has already been deleted.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC14_ExitCoauthoringSession_DeletedCoauthoringSession()
        {
            // Join the coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Exit the coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                        "The user {0} to exist the coauthoring session using the client id {1} and schema lock ID {2} should succeed.",
                        this.UserName01,
                        SharedTestSuiteHelper.DefaultClientID,
                        SharedTestSuiteHelper.ReservedSchemaLockID);

            // Cancel the record manager
            this.StatusManager.CancelSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Exit the Coauthoring Session again on the file whose coauthoring session has been deleted, the protocol server returns a Success error code.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1038
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1038,
                         @"[In Exit Coauthoring Session][When sending Exit Coauthoring Session subrequest] If the coauthoring session has already been deleted, the protocol server returns an error code value set to ""Success"", as specified in section 2.2.5.6.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                    @"If the coauthoring session has already been deleted, the protocol server returns an error code value set to ""Success""");
            }
        }

        /// <summary>
        /// A method used to verify the response information to the ExitCoauthoringSession subRequest when the client is not in the current coauth session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC15_ExitCoauthoringSession_NotPresentInCurrentSession()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Exit the Coauthoring session using a clientID which is not present in the session.
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string secondClientId = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1556, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1556
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1556,
                             @"[In Appendix B: Product Behavior] If the current client is not already present in the coauthoring session when doing the ""exiting coauthoring session"" operation, the implementation does return an error code of ""Success"", if there are other clients present in the coauthoring session. (&lt;41&gt; Section 3.1.4.3.2:  Microsoft SharePoint Foundation 2010/SharePoint Server 2010 /Microsoft Office 2019/Microsoft SharePoint Server 2019 InvalidCoauthSession.) ");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3107
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3107,
                             @"[In Appendix B: Product Behavior] If the current client is not already present in the coauthoring session when doing the ""exiting coauthoring session"" operation, the implementation does return an error code of ""Success"", if there are other clients present in the coauthoring session. (&lt;41&gt; Section 3.1.4.3.2: Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016  will return an error code of ""InvalidCoauthSession"", if there are other clients present in the coauthoring session.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1556, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                        @"If the current client is not already present in the coauthoring session when doing the ""exiting coauthoring session"" operation, the implementation does return an error code of ""Success"", if there are other clients present in the coauthoring session. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                        @"If the current client is not already present in the coauthoring session when doing the ""exiting coauthoring session"" operation, the implementation does return an error code of ""InvalidCoauthSession"", if there are other clients present in the coauthoring session. (&lt;36&gt; Section 3.1.4.3.2:  SharePoint Server 2013 will return an error code of ""InvalidCoauthSession"" if there are other clients present in the coauthoring session.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the response information to the ExitCoauthoringSession subRequest when there is a current exclusive lock on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC16_ExitCoauthoringSession_FileAlreadyLockedOnServer_ExclusiveLock()
        {
            // Get an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Exit the Coauthoring session when the file has an exclusive lock.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1605
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1605,
                         @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" when sending Exit Coauthoring Session subrequest. (<42> Section 3.1.4.3.2:  SharePoint Server 2013 and SharePoint Server 2010 return an error code of ""Success"" if there is an exclusive lock on the file.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                    @"Implementation does return an error code of ""Success"" The Office 2010 release server returns an error code of ""Success"" if there is an exclusive lock on the file.");
            }
        }

        /// <summary>
        /// A method used to verify the response information to the ExitCoauthoringSession subRequest when there is a shared lock on the file with a different or same schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC17_ExitCoauthoringSession_FileAlreadyLockedOnServer_SharedLock()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Exit the Coauthoring session using a different shared lock identifier.
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, System.Guid.NewGuid().ToString());
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 156001, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R156001
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             156001,
                             @"[In Appendix B: Product Behavior] Implementation does return an error code of ""FileAlreadyLockedOnServer"" if there is a shared lock on the file with a different schema lock identifier When sending Exit Coauthoring Session subrequest. (<42> Section 3.1.4.3.2: SharePoint Server 2013 and SharePoint Server 2010 return an error code of ""FileAlreadyLockedOnServer"" if there is a shared lock with a different shared lock identifier and a coauthoring session containing one client.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 156001, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                    @"Implementation does return an error code of ""FileAlreadyLockedOnServer"" if there is a shared lock on the file with a different schema lock identifier. (SharePoint Server 2013 and SharePoint Server 2010)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the Success response to the ExitCoauthoringSession subRequest when there is a shared lock with a different shared lock identifier and a valid coauthoring session containing more than one clients.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC18_ExitCoauthoringSession_Success_SharedLock()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Using another user to join the coauth session again
            this.PrepareCoauthoringSession(this.DefaultFileUrl, Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Exit the coauth session using different lock identifier with the previous two steps.
            this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForExitCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, Guid.NewGuid().ToString());
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType exitResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2047,
                         @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" when sending Exit Coauthoring Session subrequest. (<37> Section 3.1.4.3.2: SharePoint Server 2013 and SharePoint Server 2010 return an error code of ""Success"" if there is a shared lock with a different shared lock identifier and a valid coauthoring session containing more than one clients.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exitResponse.ErrorCode, this.Site),
                    @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" when sending Exit Coauthoring Session subrequest. (&lt;37&gt; Section 3.1.4.3.2: SharePoint Server 2013 and SharePoint Server 2010 return an error code of ""Success"" if there is a shared lock with a different shared lock identifier and a valid coauthoring session containing more than one clients.)");
            }
        }

        #endregion

        #region Mark Transition complete
        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for MarkTransitionToComplete subRequest when there is no coauthoring session on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC19_MarkTransitionComplete_InvalidCoauthSession_NoSharedLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Mark Transition to Complete
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForMarkTransitionComplete(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1904, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1904
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1904,
                             @"[In Appendix B: Product Behavior][When sending Mark Transition to Complete subrequest] The implementation does return an error code value set to ""InvalidCoauthSession"" to indicate failure if there is no coauthoring session for the file. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1903, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1903
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1903,
                             @"[In Appendix B: Product Behavior][When sending Mark Transition to Complete subrequest] The implementation does return an error code value set to ""InvalidCoauthSession"" to indicate failure if there is no shared lock. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3111, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3111
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.LockRequestFail,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3111,
                             @"[In Appendix B: Product Behavior][When sending Mark Transition to Complete subrequest] The implementation does return an error code value set to ""LockRequestFail"" to indicate failure if there is no shared lock. (<44> Section 3.1.4.3.6:  SharePoint Server 2010 returns an error code value set to ""LockRequestFail"" if there is no shared lock.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3113, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3113
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.LockRequestFail,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3113,
                             @"[In Appendix B: Product Behavior][When sending Mark Transition to Complete subrequest] The implementation does return an error code value set to ""LockRequestFail"" to indicate failure if there is no coauthoring session for the file. (<45> Section 3.1.4.3.6:  SharePoint Server 2010 returns an error code value set to ""LockRequestFail"" if there is no coauthoring session for the file)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1904, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"[When sending Mark Transition to Complete subrequest] The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any of the following conditions is true: There is no coauthoring session for the file. (SharePoint Foundation 2010/SharePoint Server 2010 follow this behaviors.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1903, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"[When sending Mark Transition to Complete subrequest] The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any of the following conditions is true: There is no shared lock. (SharePoint Foundation 2010/SharePoint Server 2010 follow this behaviors.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3111, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.LockRequestFail,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"[When sending Mark Transition to Complete subrequest] The protocol server returns an error code value set to ""LockRequestFail"" to indicate failure if any of the following conditions is true: There is no coauthoring session for the file. (SharePoint Foundation 2013/SharePoint Server 2013 and above follow this behaviors.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3113, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.LockRequestFail,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"[When sending Mark Transition to Complete subrequest] The protocol server returns an error code value set to ""LockRequestFail"" to indicate failure if any of the following conditions is true: There is no shared lock. (SharePoint Foundation 2013/SharePoint Server 2013 and above follow this behaviors.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for MarkTransitionToComplete subRequest when there is no coauthoring session on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC20_MarkTransitionComplete_Success()
        {
            // Join the Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Mark transition using the same client id and schema lock  id.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForMarkTransitionComplete(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                "Test case cannot continue unless the mark transition complete succeeds.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R1568, the CoauthStatus should not be specified, actually this attribute {0}",
                    response.SubResponseData.CoauthStatusSpecified ? "exists" : "does not exist.");

                Site.CaptureRequirementIfIsFalse(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1568,
                         @"[In Mark Transition to Complete] The CoauthStatus attribute is not set by the server in the subresponse returned for this subrequest[Mark transition to complete].");
            }
            else
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The CoauthStatus should not be specified, actually this attribute {0}",
                    response.SubResponseData.CoauthStatusSpecified ? "exists" : "does not exist.");

                Site.Assert.IsFalse(
                    response.SubResponseData.CoauthStatusSpecified,
                    @"[In Mark Transition to Complete] The CoauthStatus attribute is not set by the server in the subresponse returned for this subrequest[Mark transition to complete].");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for MarkTransitionToComplete subRequest when the current client is not present in the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC21_MarkTransitionComplete_InvalidCoauthSession_NotPresentClient()
        {
            // Join the Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Mark transition to complete using a client which isn't present in the coauthoring session
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string clientId = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForMarkTransitionComplete(clientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
            {
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1571
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1571,
                             @"[In Mark Transition to Complete][When sending Mark Transition to Complete subrequest] The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any of the following conditions is true:
                         The current client is not present in the coauthoring session.");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"The protocol server returns an error code value set to ""InvalidCoauthSession"" if the current client is not present in the coauthoring session.");
                }
            }
        }
        #endregion

        #region Check Lock Availability
        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for CheckLockAvailability subRequest if there is a current exclusive lock on the file with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC22_CheckLockAvailability_FileAlreadyLockedOnServer_ExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock.
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Check lock availability
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1492
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1492,
                         @"[In Check Lock Availability] If there is a current exclusive lock on the file  with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If there is a current exclusive lock on the file  with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for CheckLockAvailability subRequest if there is a current shared lock on the file with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC23_CheckLockAvailability_FileAlreadyLockedOnServer_SharedLock()
        {
            // Join the Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Check lock availability with a different schema lock identifier
            string schemaLock = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, schemaLock);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1493
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1493,
                         @"[In Check Lock Availability] If there is a current shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If there is a current shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify that the protocol server can check the lock availability successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC24_CheckLockAvailability_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check Lock Availability when the file does not have any lock or is not checked out.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType checkResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1084
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         10841,
                         @"[In Check Lock Availability] In all other cases[If the coauthorable file isn't checked out and there is no exclusive lock and shared lock on the file], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");

                this.CaptureSucceedCoauthSubRequest(checkResponse);
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                    @"If the coauthorable file isn't checked out and there is no exclusive lock and shared lock on the file, the protocol server returns an error code value set to ""Success"".");
            }

            // Join the Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Check the Lock again using the same schema lock ID.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1084
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         10842,
                         @"[In Check Lock Availability] In all other cases[If there is a shared lock on the file with the same lock identifier], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");

                this.CaptureSucceedCoauthSubRequest(response);
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                    @"If there is a shared lock on  the file with the same lock identifier, the protocol server returns an error code value set to ""Success"".");
            }
        }

        #endregion

        #region Get Coauthoring Session Status

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for GetCoauthoringSession subRequest if the coauthoring session does not exist.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC25_GetCoauthStatus_InvalidCoauthSession_NoCoauthoringSession()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get Coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForGetCoauthSessionStatus(System.Guid.NewGuid().ToString(), System.Guid.NewGuid().ToString());
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1106, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1106
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1106,
                             @"[In Appendix B: Product Behavior] If the coauthoring session does not exist when doing the ""Get Coauthoring session"" operation, the implementation does return an error code value set to ""InvalidCoauthSession"". (<46> Section 3.1.4.3.7:  In SharePoint Server 2010, the protocol server returns an error code value set to ""InvalidCoauthSession"".)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3115, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3115
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3115,
                             @"[In Appendix B: Product Behavior] If the coauthoring session does not exist when doing the ""Get Coauthoring session"" operation, the implementation does return an error code value set to ""Success"". (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1106, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"[In Get Coauthoring Session] If the coauthoring session does not exist, the protocol server returns an error code value set to ""InvalidCoauthSession"".");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3115, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"If the coauthoring session does not exist when doing the ""Get Coauthoring session"" operation, the implementation does return an error code value set to ""Success"". (<41> Section 3.1.4.3.7: In SharePoint Server 2013, the protocol server returns an error code value set to ""Success"".)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for GetCoauthoringSession subRequest when the current client is not present in the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC26_GetCoauthStatus_InvalidCoauthSession_NotPresentInCurrentSession()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get the Coauthoring session using a client which isn't present in the coauthoring session
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string clientId = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForGetCoauthSessionStatus(clientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1107, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1107 and MS-FSSHTTP_R2073
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1107,
                             @"[In Appendix B: Product Behavior] if the current client is not present in the coauthoring session when doing the ""Get Coauthoring session"" operation, the implementation does return an error code value set to ""InvalidCoauthSession"". (<46> Section 3.1.4.3.7:  In SharePoint Server 2010, the protocol server returns an error code value set to ""InvalidCoauthSession"".)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3117, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3117
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3117,
                             @"[In Appendix B: Product Behavior] if the current client is not present in the coauthoring session when doing the ""Get Coauthoring session"" operation, the implementation does return an error code value set to ""Success"". (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1107, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"The current client is not present in the coauthoring session, the protocol server returns an error code value set to ""InvalidCoauthSession"".(Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3117, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        @"The current client is not present in the coauthoring session, the protocol server returns an error code value set to ""Success"".(Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned by the protocol server when only one client is editing the shared file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC27_GetCoauthStatus_OnlyOneClientEditing()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get the coauthoring status of the client
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForGetCoauthSessionStatus(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType getStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(getStatusResponse.ErrorCode, this.Site), "The client should get the coauth status successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture coauthoring status common related requirements.
                this.CaptureCoauthStatusRelatedRequirementsWhenGetCoauthoringStatus(getStatusResponse);
                this.CaptureSucceedCoauthSubRequest(getStatusResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1108
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1108,
                         @"[In Get Coauthoring Session] If the current client is the only client editing the coauthorable file, the protocol server MUST set the CoauthStatus attribute value to ""Alone"", indicating that no one else is editing the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2067
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         2067,
                         @"[In CoauthStatusType][Alone] The alone status specifies that there is only one user in the coauthoring session who is editing the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1024
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1024,
                         @"[In Join Coauthoring Session] If the current client is the only client editing the file, the protocol server MUST return a CoauthStatus set to ""Alone"", which indicates that no one else is editing the file.");

                // If the current client is the only client editing the file, the protocol server returns a CoauthStatus attribute set to "Alone", which indicates that no one else is editing the file.
                // If the returned CoauthStatus field is set to "Alone", MS-FSSHTTP_R1023 could be captured.
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1023,
                         @"[In Join Coauthoring Session] To get the coauthoring status, the protocol server checks the number of clients editing the file at that instant in time.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Alone,
                    getStatusResponse.SubResponseData.CoauthStatus,
                    @"If the current client is the only client editing the coauthorable file, the protocol server MUST set the CoauthStatus attribute value to ""Alone"".");
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned by the protocol server when two clients are editing the shared file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC28_GetCoauthStatus_TwoClientsEditing()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // join the Coauthoring session using the second client
            string secondClientId = System.Guid.NewGuid().ToString();
            this.PrepareCoauthoringSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Get the coauthoring status of the second client
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForGetCoauthSessionStatus(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType getStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(getStatusResponse.ErrorCode, this.Site), "The second client should get the coauth status successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture coauthoring status common related requirements.
                this.CaptureCoauthStatusRelatedRequirementsWhenGetCoauthoringStatus(getStatusResponse);
                this.CaptureSucceedCoauthSubRequest(getStatusResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1109
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1109,
                         @"[In Get Coauthoring Session] If the current client is the second client trying to edit the document, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R317
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         317,
                         @"[In CoauthStatusType] Coauthoring [means]: A string value of ""Coauthoring"", indicating a coauthoring status of coauthoring.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2068
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         2068,
                         @"[In CoauthStatusType][Coauthoring] The coauthoring status specifies that the targeted URL for the file has more than one user in the coauthoring session and that the file is being edited by more than one user.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Coauthoring,
                    getStatusResponse.SubResponseData.CoauthStatus,
                    @"If the current client is not the only client trying to edit the document, the protocol server should return a CoauthStatus set to ""Coauthoring""");
            }
        }
        #endregion

        #region Refresh Coauthoring Session
        /// <summary>
        /// A method used to verify the protocol server refreshes the client's timeout and checks the coauthoring status of the file successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC29_RefreshCoauthoringSession_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session with Timeout attribute set to 60
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 60);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3067
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3067,
                         @"[In CoauthSubRequestDataType][Timeout] When the Timeout attribute is set to a value ranging from 60 to 3600, the server also returns success [but sets Timeout to an implementation-specific default value].");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    "[Refresh coauthoring session subrequest] When the Timeout attribute is set to a value ranging from 60 to 3600, the server also returns success.");
            }

            // Refresh the Coauthoring session with a timeout set to 3600
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, 3600);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "The client should refresh the coauthoring session successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.CaptureSucceedCoauthSubRequest(subResponse);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expect the LockType exist when the server response error code Success when sending the refresh coauthoring session subrequest, actual result is {0}",
                    subResponse.SubResponseData.LockTypeSpecified ? "exist" : "not exist");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R997
                Site.CaptureRequirementIfIsTrue(
                         subResponse.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         997,
                         @"[In Coauth Subrequest][If the coauthoring subrequest is of type] ""Refresh coauthoring session"", the protocol server MUST return the lock type granted to the client as part of the response message to the client—if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");
            }
            else
            {
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expect the LockType exist when the server response error code Success when sending the refresh coauthoring session subrequest, actual result is {0}",
                    subResponse.SubResponseData.LockTypeSpecified ? "exist" : "not exist");

                Site.Assert.AreEqual<string>(
                    "SchemaLock",
                    subResponse.SubResponseData.LockType,
                    @"When sending the refresh coauthoring session subrequest, the server return the LockType attribute when the error code is ""Success"". ");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for RefreshCoauthoringSession subRequest if there is a current exclusive lock on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC30_RefreshCoauthoringSession_FileAlreadyLockedOnServer_ExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Refresh the Coauthoring session
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1574
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1574,
                         @"[In Refresh Coauthoring Session][When sending Refresh Coauthoring Session subrequest] If there is a current exclusive lock on the file on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"[When sending Refresh Coauthoring Session subrequest] If there is a current exclusive lock on the file on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "FileAlreadyLockedOnServer" for RefreshCoauthoringSession subRequest if there is a shared lock on the file with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC31_RefreshCoauthoringSession_FileAlreadyLockedOnServer_SharedLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The client should join the coauthoring session successfully.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Refresh the Coauthoring session with a different schema lock identifier
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(System.Guid.NewGuid().ToString(), System.Guid.NewGuid().ToString());
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1565
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1565,
                         @"[In Refresh Coauthoring Session][When sending Refresh Coauthoring Session subrequest] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"[When sending Refresh Coauthoring Session subrequest] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned for RefreshCoauthoringSession subRequest when there is only one client editing the shared file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC32_RefreshCoauthSession_OnlyOneClient()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session.
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Refresh the Coauthoring session
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType refreshResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(refreshResponse.ErrorCode, this.Site), "The client should refresh the coauthoring session successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.CaptureCoauthStatusRelatedRequirementsWhenRefreshCoauthoringSession(refreshResponse);
                this.CaptureLockTypeRelatedRequirementsWhenRefreshCoauthoringSession(refreshResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1055
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         refreshResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1055,
                         @"[In Refresh Coauthoring Session] If the current client is the only client editing the file, the protocol server MUST return a CoauthStatus set to ""Alone"", which indicates that no one else is editing the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1054
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         refreshResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1054,
                         @"[In Refresh Coauthoring Session] To get the coauthoring status, the protocol server checks the number of clients editing the file at that instant in time.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Alone,
                    refreshResponse.SubResponseData.CoauthStatus,
                    @"If the current client is the only client editing the file, the protocol server MUST return a CoauthStatus set to ""Alone"".");
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned for RefreshCoauthoringSession subRequest when there are two clients editing the shared file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC33_RefreshCoauthSession_TwoOrThirdClients()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Use SUT method to set max number of coauthors to 4.
            bool isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(4);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers to 4 should succeed.");

            // Join a Coauthoring session.
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The operation CoauthingSubRequest should succeed.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // join the Coauthoring session using the second client
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string secondClientId = System.Guid.NewGuid().ToString();
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The operation CoauthingSubRequest should succeed.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Refresh the Coauthoring session using the second client
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType refreshResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(refreshResponse.ErrorCode, this.Site), "The second client should refresh the coauthoring session successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.CaptureCoauthStatusRelatedRequirementsWhenRefreshCoauthoringSession(refreshResponse);
                this.CaptureLockTypeRelatedRequirementsWhenRefreshCoauthoringSession(refreshResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1054
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         refreshResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1054,
                         @"[In Refresh Coauthoring Session] To get the coauthoring status, the protocol server checks the number of clients editing the file at that instant in time.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1056
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         refreshResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1056,
                         @"[In Refresh Coauthoring Session] If the current client is the second coauthor joining the coauthoring session, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Coauthoring,
                    refreshResponse.SubResponseData.CoauthStatus,
                    @"If the coauthoring session contains more than one client, the protocol server should return a CoauthStatus set to ""Coauthoring"", if the client sends refresh coauthoring session.");
            }

            // Join the Coauthoring session using the third client
            this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);
            string thirdClientId = System.Guid.NewGuid().ToString();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            while (retryCount > 0)
            {
                request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(thirdClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
                response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
                joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site) != ErrorCodeType.Success)
                {
                    System.Threading.Thread.Sleep(waitTime);
                    retryCount--;
                }
                else
                {
                    this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, thirdClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName03, this.Password03, this.Domain);
                    break;
                }

                if (retryCount == 0)
                {
                    this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The third client should join the coauthoring session successfully.");
                }
            }

            // Refresh the Coauthoring session using the third client
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(thirdClientId, SharedTestSuiteHelper.ReservedSchemaLockID);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType anotherRefreshResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(anotherRefreshResponse.ErrorCode, this.Site), "The third client should refresh the coauthoring session successfully.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.CaptureCoauthStatusRelatedRequirementsWhenRefreshCoauthoringSession(anotherRefreshResponse);
                this.CaptureLockTypeRelatedRequirementsWhenRefreshCoauthoringSession(anotherRefreshResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1887
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         refreshResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1887,
                         @"[In Refresh Coauthoring Session] If the current client is the third coauthor joining the coauthoring session, the protocol server MUST return a CoauthStatus set to ""Coauthoring"", which indicates that the current client is coauthoring when editing the document.");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Coauthoring,
                    refreshResponse.SubResponseData.CoauthStatus,
                    @"If the coauthoring session contains more than one client, the protocol server should return a CoauthStatus set to ""Coauthoring"", if the client sends refresh coauthoring session.");
            }

            isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");
        }

        /// <summary>
        /// A method used to verify the protocol server's behavior when refreshing the coauthoring lock using a clientID which isn't present in the session if the coauthoring feature is disabled.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC34_RefreshCoauthoringSession_CoauthDisableAndTimeExpires()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session with timeout set to 60s.
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 60);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site), "The operation CoauthingSubRequest should succeed.");
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Disable the coauthoring feature.
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Sleep 70s to wait the timeout expires.
            SharedTestSuiteHelper.Sleep(70);

            // Refresh the coauthoring session
            request = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType refreshSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3108, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3108
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3108,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<45> Section 3.1.4.3.3:  SharePoint Server 2010 will return an error code value ""FileNotLockedOnServerAsCoauthDisabled"".)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 310801, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R310801
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             310801,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<45> Section 3.1.4.3.3:  SharePoint Server 2019 and SharePoint Server Subscription Edition will return an error code value ""Success"".)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3109, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3109
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3109,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""FileNotLockedOnServer"". (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3108, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                        @"When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (Microsoft Office 2010 suites/Microsoft Office 2013/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 310801, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                        @"When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""Success"". (SharePoint Server 2019 and SharePoint Server Subscription Edition follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3109, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServer,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(refreshSubResponse.ErrorCode, this.Site),
                        @"When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock in the file coauthoring tracker, if the coauthoring feature is disabled, the implementation does return an error code value set to ""FileNotLockedOnServer"". (<38> Section 3.1.4.3.3: SharePoint Foundation 2013 and SharePoint Server 2013 will return an error code value ""FileNotLockedOnServer"".)");
                }
            }
        }
        #endregion

        #region Convert To Exclusive Lock

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for ConvertToExclusiveLock subRequest if there is a current exclusive lock on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC35_ConverToExclusiveLock_InvalidCoauthSession_ExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert the shared lock to an exclusive lock using another user with a different exclusive lock identifier
            string clientId = System.Guid.NewGuid().ToString();
            string schemaLockId = System.Guid.NewGuid().ToString();
            string exclusiveLockId = System.Guid.NewGuid().ToString();

            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(clientId, schemaLockId, exclusiveLockId, false);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1892
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1892,
                         @"[In Convert to Exclusive Lock] If there is a current exclusive lock on the file, the request cannot be processed successfully so the protocol server returns ""InvalidCoauthSession"" as defined in section 2.2.5.8.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                    @"If there is a current exclusive lock on the file, the request cannot be processed successfully so the protocol server returns ""InvalidCoauthSession"".");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidCoauthSession" for ConvertToExclusiveLock subRequest if there is a shared lock on the file from another client with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC36_ConverToExclusiveLock_InvalidCoauthSession_DifferentSchemaLockID()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Convert the shared lock to an exclusive lock using another client with a different schema lock identifier
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string clientId = System.Guid.NewGuid().ToString();
            string schemaLockId = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(clientId, schemaLockId, SharedTestSuiteHelper.DefaultExclusiveLockID, false);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1893
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1893,
                         @"[In Convert to Exclusive Lock] If there is a shared lock on the file from another client with a different schema lock identifier, the request cannot be processed successfully so the protocol server returns ""InvalidCoauthSession"" as defined in section 2.2.5.8.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                    @"If there is a shared lock on the file from another client with a different schema lock identifier, the request cannot be processed successfully so the protocol server returns ""InvalidCoauthSession"".");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server converts the shared lock to an exclusive lock and deletes the coauthoring session when it receives a ConvertToExclusiveLock subRequest.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC37_ConvertToExclusiveLock_Success()
        {
            // Join a Coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Convert the shared lock to an exclusive lock
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                "The client should covert the shared lock to en exclusive lock successfully.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Capture the succeed coauthoring request related requirements.
            this.CaptureSucceedCoauthSubRequest(convertResponse);

            // Check the exclusive lock exist with different exclusive lock id.
            ExclusiveLockSubRequestType exclusiveLockSubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            exclusiveLockSubRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLockSubRequest });
            ExclusiveLockSubResponseType checkLockResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the file is locked, then indicates previous conversion operation succeeds.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1895
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(checkLockResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1895,
                         @"[In Convert to Exclusive Lock] The shared lock is converted to an exclusive lock if one client is currently editing the document. ");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(checkLockResponse.ErrorCode, this.Site),
                    "The shared lock should be converted to an exclusive lock if one client is currently editing the document. ");
            }
        }

        /// <summary>
        /// A method used to verify the response information when the ReleaseLockOnConversionToExclusiveFailure attribute is set to true and the conversion to exclusive lock failed.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC38_ConvertToExclusiveLock_MultipleClients_ReleaseLockTrue()
        {
            // Use SUT method to set the max number of coauthors to 3.
            bool isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(3);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");

            // Join a Coauthoring session use the first user.
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join a Coauthoring session use the second user.
            string secondClientId = System.Guid.NewGuid().ToString();
            this.PrepareCoauthoringSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Convert the shared Lock to exclusive Lock using the second user with ReleaseLockOnConversionToExclusiveFailure attribute set to true,
            // the protocol server should return a ExitCoauthSessionAsConvertToExclusiveFailed error code.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, SharedTestSuiteHelper.DefaultExclusiveLockID, true);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertToExclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1072 and MS-FSSHTTP_R395
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertToExclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1072,
                         @"[In Convert to Exclusive Lock] The protocol server returns an error code value set to ""ExitCoauthSessionAsConvertToExclusiveFailed"" when the following conditions are both true:
                         1.The ReleaseLockOnConversionToExclusiveFailure attribute is set to true in the coauthoring subrequest;
                         2.Multiple clients are in the coauthoring session.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R395.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertToExclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         395,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] ExitCoauthSessionAsConvertToExclusiveFailed indicates an error when a coauthoring subrequest or schema lock subrequest of type ""Convert to exclusive lock"" is sent by the client with the ReleaseLockOnConversionToExclusiveFailure attribute set to true, and there is more than one client editing the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R444
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertToExclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         444,
                         @"[In SubRequestDataOptionalAttributes] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker when all of the following conditions are true:
Either the type of coauthoring subrequest is ""Convert to an exclusive lock"" or the type of the schema lock subrequest is ""Convert to an Exclusive Lock"".
The conversion to an exclusive lock failed.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    bool isInCoauthoringSesssion = this.IsPresentInCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "The user {0} with client ID {0} and schema Lock ID {1} should not exist in the coauthoring session, but actual it {2}",
                        this.UserName02,
                        secondClientId,
                        SharedTestSuiteHelper.ReservedSchemaLockID,
                        isInCoauthoringSesssion ? "exists" : "does not exist");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1073, MS-FSSHTTP_R573 and MS-FSSHTTP_R445.
                    Site.CaptureRequirementIfIsFalse(
                             isInCoauthoringSesssion,
                             "MS-FSSHTTP",
                             1073,
                             @"[In Convert to Exclusive Lock] When the ReleaseLockOnConversionToExclusiveFailure attribute is set to true and the conversion to an exclusive lock failed, the protocol server removes the client from the coauthoring session for the file.");

                    Site.CaptureRequirementIfIsFalse(
                             isInCoauthoringSesssion,
                             "MS-FSSHTTP",
                             573,
                             @"[In CoauthSubRequestDataType][ReleaseLockOnConversionToExclusiveFailure] When all the above conditions[1. The type of coauthoring subrequest is ""Convert to an exclusive lock"" 2. The conversion to an exclusive lock failed.] are true:
                         A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of true indicates that the protocol server is allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.");

                    Site.CaptureRequirementIfIsFalse(
                             isInCoauthoringSesssion,
                             "MS-FSSHTTP",
                             445,
                             @"[In SubRequestDataOptionalAttributes] When all the above conditions[1. The type of co-authoring sub request is ""Convert to an exclusive lock"" or the type of the schema lock sub request is ""Convert to an Exclusive Lock"" 2. The conversion to an exclusive lock failed] are true:
A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of true indicates that the protocol server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker.");
                }
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(convertToExclusiveResponse.ErrorCode, this.Site),
                    @"The protocol server returns an error code value set to ""ExitCoauthSessionAsConvertToExclusiveFailed"" when the following conditions are both true:
                        1.The ReleaseLockOnConversionToExclusiveFailure attribute is set to true in the coauthoring subrequest;
                        2.Multiple clients are in the coauthoring session.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    bool isInCoauthoringSesssion = this.IsPresentInCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "The user {0} with client ID {0} and schema Lock ID {1} should not exist in the coauthoring session, but actual it {2}",
                        this.UserName02,
                        secondClientId,
                        SharedTestSuiteHelper.ReservedSchemaLockID,
                        isInCoauthoringSesssion ? "exists" : "does not exist");

                    Site.Assert.IsFalse(
                             isInCoauthoringSesssion,
                             @"[In Convert to Exclusive Lock] When the ReleaseLockOnConversionToExclusiveFailure attribute is set to true and the conversion to an exclusive lock failed, the protocol server removes the client from the coauthoring session for the file.");
                }
            }

            isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");
        }

        /// <summary>
        /// A method used to verify the coauthoring transition allows for the number of users editing the file to increase from 1 to n or to decrease from n to 1, where n is the maximum number of users who are allowed to edit a single file at an instant time.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC39_ConvertToExclusiveLock_MultipleClients_ReleaseLockFalse()
        {
            // Use SUT method to set the max number of coauthors to 3.
            bool isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(3);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");

            // Join a Coauthoring session use the first user.
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Join a Coauthoring session use the second user.
            string secondClientId = System.Guid.NewGuid().ToString();
            this.PrepareCoauthoringSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Convert the shared lock to an exclusive lock using the second user with ReleaseLockOnConversionToExclusiveFailure attribute set to false
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, SharedTestSuiteHelper.DefaultExclusiveLockID, false);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1071 and MS-FSSHTTP_R38901
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.MultipleClientsInCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1071,
                         @"[In Convert to Exclusive Lock] If there is more than one client currently editing the file and the ReleaseLockOnConversionToExclusiveFailure attribute is set to false, then the protocol server returns an error code value set to ""MultipleClientsInCoauthSession"" to indicate the failure to convert to exclusive lock.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.MultipleClientsInCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         38901,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] MultipleClientsInCoauthSession indicates an error when all of the following conditions are true:
                         1.A coauthoring subrequest of type ""Convert to exclusive lock"" is requested on a file;
                         2.There is more than one client in the current coauthoring session for that file;
                         3.The ReleaseLockOnConversionToExclusiveFailure attribute specified as part of the subrequest is set to false.");

                bool isInCoauthoringSesssion = this.IsPresentInCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The user {0} with client ID {0} and schema Lock ID {1} should exist in the coauthoring session, but actual it {2}",
                    this.UserName02,
                    secondClientId,
                    SharedTestSuiteHelper.ReservedSchemaLockID,
                    isInCoauthoringSesssion ? "exists" : "does not exist");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R574 and MS-FSSHTTP_R446
                Site.CaptureRequirementIfIsTrue(
                         isInCoauthoringSesssion,
                         "MS-FSSHTTP",
                         446,
                         @"[In SubRequestDataOptionalAttributes][When all the above conditions 1. The type of co-authoring sub request is ""Convert to an exclusive lock"" or the type of the schema lock sub request is ""Convert to an Exclusive Lock"" 2. The conversion to an exclusive lock failed] are true:
A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of false indicates that the protocol server is not allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker.");

                Site.CaptureRequirementIfIsTrue(
                         isInCoauthoringSesssion,
                         "MS-FSSHTTP",
                         574,
                         @"[In CoauthSubRequestDataType][ReleaseLockOnConversionToExclusiveFailure] When all the above conditions[1. The type of coauthoring subrequest is ""Convert to an exclusive lock"" 2. The conversion to an exclusive lock failed.] are true: A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of false indicates that the protocol server is not allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.MultipleClientsInCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If there is more than one client currently editing the file and the ReleaseLockOnConversionToExclusiveFailure attribute is set to false, then the protocol server returns an error code value set to ""MultipleClientsInCoauthSession"" to indicate the failure to convert to exclusive lock.");

                bool isInCoauthoringSesssion = this.IsPresentInCoauthSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The user {0} with client ID {0} and schema Lock ID {1} should exist in the coauthoring session, but actual it {2}",
                    this.UserName02,
                    secondClientId,
                    SharedTestSuiteHelper.ReservedSchemaLockID,
                    isInCoauthoringSesssion ? "exists" : "does not exist");

                Site.Assert.IsTrue(
                        isInCoauthoringSesssion,
                        @"[In SubRequestDataOptionalAttributes][When all the above conditions 1. The type of co-authoring sub request is ""Convert to an exclusive lock"" or the type of the schema lock sub request is ""Convert to an Exclusive Lock"" 2. The conversion to an exclusive lock failed] are true:
A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of false indicates that the protocol server is not allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker.");
            }

            isSetMaxNumberSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isSetMaxNumberSuccess, "The operation SetMaxNumOfCoauthUsers should succeed.");
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to "InvalidCoauthSession" for ConvertToExclusiveLock subRequest when there is no shared lock or coauthoring session on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC40_ConvertToExclusiveLock_NoSharedLockOrSession()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Convert a non-exist shared lock to an exclusive lock
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1064, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1064
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1064,
                             @"[In Convert to Exclusive Lock] The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate a failure if any of the following conditions is true: 
                         There is no shared lock.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1065
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1065,
                             @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] 
                         There is no coauthoring session for the file.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1064, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                    @"The protocol server returns an error code value set to ""InvalidCoauthSession"", if there is no shared lock or there is no coauthoring session for the file.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to "InvalidCoauthSession" for ConvertToExclusiveLock subRequest when the current client is not present in the coauthoring session on the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC41_ConvertToExclusiveLock_NotPresentInSession()
        {
            // Join a Coauthoring session use the first user.
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Convert the shared lock to an exclusive lock using another client
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            string secondClientId = System.Guid.NewGuid().ToString();
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForConvertToExclusiveLock(secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, SharedTestSuiteHelper.DefaultExclusiveLockID, false);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType convertResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1066
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1066,
                         @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] 
                         The current client is not present in the coauthoring session.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2073
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2073,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][InvalidCoauthSession indicates an error when one of the following conditions is true when a coauthoring subrequest or schema lock subrequest is sent:] The current client does not exist in the coauthoring session for the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(convertResponse.ErrorCode, this.Site),
                    @"The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if the current client is not present in the coauthoring session.");
            }
        }
        #endregion

        #region Common

        /// <summary>
        /// A method used to verify the protocol server returns an error code "InvalidArgument" for Coauth subRequest when the schema lock ID is not provided in the subRequest.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC42_Coauth_NoSchemaLockID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session without schema lock identifier
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.JoinCoauthoring, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R5601
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         5601,
                         @"[In CoauthSubRequestDataType]If other attributes are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the coauthoring subresponse.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R68
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         68,
                         @"[In Request][Errors that occur during the parsing of the Request element] The protocol server MUST send the error code[InvalidArgument] as an error code attribute in the Response element.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R365
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         365,
                         @"[In GenericErrorCodeTypes] InvalidArgument indicates an error when any of the cell storage service subrequests for the targeted URL for the file contains input parameters that are not valid.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2246
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2246,
                         @"[In CoauthSubResponseType] The protocol server must not set the value of the ErrorCode attribute to ""Success"" if the protocol server fails in processing the coauthoring subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R365
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         365,
                         @"[In GenericErrorCodeTypes] InvalidArgument indicates an error when any of the cell storage service subrequests for the targeted URL for the file contains input parameters that are not valid.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidArgument,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"If the specified attributes[SchemaLockID] are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the coauthoring subresponse.");
            }

            // Refresh Coauthoring session without schema lock identifier
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.RefreshCoauthoring, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.InvalidArgument, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The protocol server should return errorCode 'InvalidArgument'.");

            // Mark Transition to Complete without schema lock identifier
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.MarkTransitionComplete, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.InvalidArgument, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The protocol server should return errorCode 'InvalidArgument'.");

            // convert to Exclusive without schema lock identifier
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.ConvertToExclusive, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.InvalidArgument, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The protocol server should return errorCode 'InvalidArgument'.");

            // Check Lock Availability without schema lock identifier
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.CheckLockAvailability, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.InvalidArgument, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The protocol server should return errorCode 'InvalidArgument'.");

            // Exit Coauthoring session without schema lock identifier
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequest(SequenceNumberGenerator.GetCurrentToken(), CoauthRequestTypes.ExitCoauthoring, SharedTestSuiteHelper.DefaultClientID, null, false, null, SharedTestSuiteHelper.DefaultExclusiveLockID, SharedTestSuiteHelper.DefaultTimeOut);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.InvalidArgument, SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site), "The protocol server should return errorCode 'InvalidArgument'.");
        }

        /// <summary>
        /// A method used to verify the response has unique TransitionID when use different files.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC43_Coauth_UniqueTransitionID()
        {
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    "Join coauthoring session with client ID {0} and schema lock ID {4} by the user {1}@{2} on the file {3} should succeed.",
                    SharedTestSuiteHelper.DefaultClientID,
                    this.UserName01,
                    this.Domain,
                    this.DefaultFileUrl,
                    SharedTestSuiteHelper.ReservedSchemaLockID);
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);

            // Store the first transition ID.
            string transitionId1 = joinResponse.SubResponseData.TransitionID;

            // Prepare second file.
            string anotherFileUrl = this.PrepareFile();

            // Join a Coauthoring session on the second file
            this.InitializeContext(anotherFileUrl, this.UserName01, this.Password01, this.Domain);
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(anotherFileUrl, new SubRequestType[] { subRequest });
            joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                   ErrorCodeType.Success,
                   SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                   "Join coauthoring session with client ID {0} and schema lock ID {4} by the user {1}@{2} on the file {3} should succeed.",
                   SharedTestSuiteHelper.DefaultClientID,
                   this.UserName01,
                   this.Domain,
                   anotherFileUrl,
                   SharedTestSuiteHelper.ReservedSchemaLockID);

            this.StatusManager.RecordCoauthSession(anotherFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);

            // Store the second transition ID.
            string transitionId2 = joinResponse.SubResponseData.TransitionID;

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreNotEqual<string>(
                         transitionId1,
                         transitionId2,
                         "MS-FSSHTTP",
                         61101,
                         @"[In CoauthSubResponseDataType] TransitionID: A guid specifies that if 2 requests are operated on 2 different files in the protocol server, the TransitionID values returned in the 2 corresponding responses are different.");
            }
            else
            {
                Site.Assert.AreNotEqual<string>(
                    transitionId1,
                    transitionId2,
                    @"[In CoauthSubResponseDataType] TransitionID: A guid specifies that if 2 requests are operated on 2 different files in the protocol server, the TransitionID values returned in the 2 corresponding responses are different.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code when the request token is empty string.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC44_RequestTokenIsEmptyString()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join Coauthoring session without the request token.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest }, string.Empty);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3010, this.Site))
            {
                SharedTestSuiteHelper.CheckCellStorageResponse(cellResponse, this.Site, 0);
                Response response = cellResponse.ResponseCollection.Response[0];
                this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "If the RequestToken attribute of the corresponding Request element is an empty string, the implementation does return the ErrorCode attribute in ResponseVersion element, it actually {0}",
                    response.ErrorCodeSpecified ? "returns" : "does not return");

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfIsTrue(
                             response.ErrorCodeSpecified,
                             "MS-FSSHTTP",
                             3010,
                             @"[In Appendix B: Product Behavior] If the RequestToken attribute of the corresponding Request element is an empty string, the implementation does return the ErrorCode attribute in Response element. (<12> Section 2.2.3.5:  In SharePoint Foundation 2010 and SharePoint Server 2010, the ErrorCode attribute is present if the RequestToken attribute of the corresponding Request element is an empty string.)");
                }
                else
                {
                    Site.Assert.IsTrue(
                            response.ErrorCodeSpecified,
                            @"If the RequestToken attribute of the corresponding Request element is an empty string, the implementation does return the ErrorCode attribute in Response element. (<12> Section 2.2.3.5:  In SharePoint Foundation 2010 and SharePoint Server 2010, the ErrorCode attribute is present if the RequestToken attribute of the corresponding Request element is an empty string.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned None when 1. file check out by current user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC45_JoinCoauthoringSession_None_FileCheckOutByCurrentUser()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check out the file
            if (!this.SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CheckLockAvailability();

            // Get the coauthoring status of the client
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType getStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187801
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187801,
                         @"[In Join Coauthoring Session] A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187802
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187802,
                         @"[In Join Coauthoring Session] This can happen when 1. the file is checked out by the current user [or 2. an exclusive lock is held by the current user or 3. the coauthoring feature is disabled and the AllowFallbackToExclusive attribute is set to true on the request.] ");

            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.None,
                    getStatusResponse.SubResponseData.CoauthStatus,
                    @"A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned None when 2. an exclusive lock is held by the current user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC46_JoinCoauthoringSession_None_FileExclusiveLockHeldByCurrentUser()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Get the exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Get the coauthoring status of the client
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            
            CoauthSubResponseType getStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187801
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187801,
                         @"[In Join Coauthoring Session] A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187802
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187803,
                         @"[In Join Coauthoring Session] This can happen when [1. the file is checked out by the current user or] 2. an exclusive lock is held by the current user [or 3. the coauthoring feature is disabled and the AllowFallbackToExclusive attribute is set to true on the request]. ");

            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.None,
                    getStatusResponse.SubResponseData.CoauthStatus,
                    @"A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");
            }
        }

        /// <summary>
        /// A method used to verify the coauthoring status returned None when 2. an exclusive lock is held by the current user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S02_TC47_JoinCoauthoringSession_None_FileCoauthoringFeatureDisableAndAllowFallbackToExclusiveSetTrue()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);          

            // Disable the Coauthoring Feature
            bool isSwitchedSuccessfully = SutPowerShellAdapter.SwitchCoauthoringFeature(true);
            this.Site.Assert.IsTrue(isSwitchedSuccessfully, "The Coauthoring Feature should be disabled.");
            this.StatusManager.RecordDisableCoauth();

            // Waiting change takes effect
            System.Threading.Thread.Sleep(20 * 1000);

            // Create a JoinCoauthoringSession subRequest with AllowFallbackToExclusive set to true.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            subRequest.SubRequestData.AllowFallbackToExclusive = true;
            
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            
            CoauthSubResponseType getStatusResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187801
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187801,
                         @"[In Join Coauthoring Session] A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_187802
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.None,
                         getStatusResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         187804,
                         @"[In Join Coauthoring Session] This can happen when [1. the file is checked out by the current user or 2. an exclusive lock is held by the current user or] 3. the coauthoring feature is disabled and the AllowFallbackToExclusive attribute is set to true on the request.  ");
            }
            else
            {
                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.None,
                    getStatusResponse.SubResponseData.CoauthStatus,
                    @"A CoauthStatus of ""None"" can be returned in situations where coauthoring is not achieved because an exclusive lock is returned instead of a shared lock.");
            }
        }
        #endregion

        #endregion

        #region Private helper function

        /// <summary>
        /// A method used to capture LockType related requirements when joining coauthoring session.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureLockTypeRelatedRequirementsWhenJoinCoauthoringSession(CoauthSubResponseType response)
        {
            this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "When the request is join coauthoring session, for the requirement MS-FSSHTTP_R469, MS-FSSHTTP_R1021 and MS-FSSHTTP_R605, the LockTypeSpecified value MUST be true, but actual value is {0}",
                    response.SubResponseData.LockTypeSpecified);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1021
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         1021,
                         @"[In Join Coauthoring Session] The result of the lock type gotten by the server MUST be sent as the LockType attribute in the CoauthSubResponseDataType.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R469
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         469,
                         @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A coauthoring subrequest of type ""Join coauthoring session"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R605
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         605,
                         @"[In CoauthSubResponseDataType][LockType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Join coauthoring session.");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.SubResponseData.LockTypeSpecified,
                    @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A coauthoring subrequest of type ""Join coauthoring session"".");
            }
        }

        /// <summary>
        /// A method used to capture LockType related requirements when refreshing coauthoring session.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureLockTypeRelatedRequirementsWhenRefreshCoauthoringSession(CoauthSubResponseType response)
        {
            this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "When the request is refresh coauthoring session, for the requirement MS-FSSHTTP_R1430 and MS-FSSHTTP_R1421 the LockTypeSpecified value MUST be true, but actual value is {0}",
                    response.SubResponseData.LockTypeSpecified);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1430
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         1430,
                         @"[In CoauthSubResponseDataType][LockType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Refresh coauthoring session.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1421
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         1421,
                         @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A coauthoring subrequest of type ""Refresh coauthoring session"".");
            }
            else
            {
                Site.Assert.IsTrue(
                     response.SubResponseData.LockTypeSpecified,
                     @"If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Refresh coauthoring session.");
            }
        }

        /// <summary>
        /// A method used to capture CoauthStatus related requirements when joining coauthoring session.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureCoauthStatusRelatedRequirementsWhenJoinCoauthoringSession(CoauthSubResponseType response)
        {
            Site.Log.Add(
                LogEntryKind.Debug,
                "The implementation does return CoauthStatus attribute when sending a Join coauth subRequest successfully and the subRequest isn't fallen back to an exclusive lock, actual the attribute {0}",
                response.SubResponseData.LockTypeSpecified ? "exist" : "not exist");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         3330,
                         @"[In Appendix B: Product Behavior] The implementation does return CoauthStatus attribute when sending a Join coauth subRequest successfully and the subRequest isn't fallen back to an exclusive lock subrequest. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");

                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         3331,
                         @"[In Appendix B: Product Behavior] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the implementation does return CoauthStatus attribute when client tries to join the coauthoring session and the subrequest does not fall back to an exclusive lock subrequest. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.SubResponseData.LockTypeSpecified,
                    @"If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the implementation does return CoauthStatus attribute when client tries to join the coauthoring session and the subrequest does not fall back to an exclusive lock subrequest. ");
            }
        }

        /// <summary>
        /// A method used to capture CoauthStatus related requirements when refreshing coauthoring session.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureCoauthStatusRelatedRequirementsWhenRefreshCoauthoringSession(CoauthSubResponseType response)
        {
            // Add the log information.
            Site.Log.Add(
                LogEntryKind.Debug,
                "For the requirement MS-FSSHTTPB_R1424 and MS-FSSHTTP_R1431, the coauth status should be returned when operation is \"Refresh coauthoring session\" and server responds Success");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1424
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1424,
                         @"[In SubResponseDataOptionalAttributes][CoauthStatus] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the CoauthStatus attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A coauthoring subrequest of type ""Refresh coauthoring session"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1431
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1431,
                         @"[In CoauthSubResponseDataType][CoauthStatusType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", CoauthStatus MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Refresh coauthoring session.");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.SubResponseData.CoauthStatusSpecified,
                    @"[Refresh] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", CoauthStatus MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Refresh coauthoring session.");
            }
        }

        /// <summary>
        /// A method used to capture CoauthStatus related requirements when getting coauthoring status.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureCoauthStatusRelatedRequirementsWhenGetCoauthoringStatus(CoauthSubResponseType response)
        {
            // Add the log information.
            Site.Log.Add(
                LogEntryKind.Debug,
                "For the requirement MS-FSSHTTPB_R1432 and MS-FSSHTTP_R934, the coauth status should be returned when operation is \"Get coauthoring status\" and server responds Success");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1432
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1432,
                         @"[In CoauthSubResponseDataType][CoauthStatusType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", CoauthStatus MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Get coauthoring status.");

                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1425,
                         @"[In SubResponseDataOptionalAttributes][CoauthStatus] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the CoauthStatus attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A coauthoring subrequest of type ""Get coauthoring status"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R934
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         934,
                         @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] If the protocol server supports the coauthoring subrequest, it MUST return a coauthoring status as specified in section 2.3.1.7.");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.SubResponseData.CoauthStatusSpecified,
                    @"If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", CoauthStatus MUST be specified in a coauthoring subresponse that is generated in response to all of the following types of coauthoring subrequests: Get coauthoring status.");
            }
        }

        /// <summary>
        /// A method used to capture CoauthSubRequest related requirements when the Coauth operation executes successfully.
        /// </summary>
        /// <param name="response">A return value represents the CoauthSubResponse information.</param>
        private void CaptureSucceedCoauthSubRequest(CoauthSubResponseType response)
        {
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1013
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1013,
                         @"[In Coauth Subrequest][The protocol returns results based on the following conditions:] An ErrorCode value of ""Success"" indicates success in processing the coauthoring request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1473
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1473,
                         @"[In CoauthSubResponseType] The protocol server sets the value of the ErrorCode attribute to ""Success""  if the protocol server succeeds in processing the coauthoring subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R931
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         931,
                         @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] If the protocol server supports shared locking, the protocol server MUST support at least one of the following subrequests:
                         The coauthoring subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R355
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         355,
                         @"[In GenericErrorCodeTypes] Success indicates that the cell storage service subrequest succeeded for the given URL for the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                    @"An ErrorCode value of ""Success"" indicates success in processing the coauthoring request.");
            }
        }

        #endregion
    }
}