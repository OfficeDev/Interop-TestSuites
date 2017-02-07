namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with ExclusiveLock operation.
    /// </summary>
    [TestClass]
    public abstract class S04_ExclusiveLock : SharedTestSuiteBase
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
        public void S04_ExclusiveLockInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should be unique for each test case
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test cases

        #region Get Exclusive Lock
        /// <summary>
        /// A method used to verify the related requirements when the Get Lock of ExclusiveLock subRequest succeeds. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC01_GetExclusiveLock_Success()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Set the timeout value to 60 seconds and then processing the exclusive lock request.
            subRequest.SubRequestData.Timeout = "60";

            // Get the exclusive lock with time out value 60 seconds and expect the server returns error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1236
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1236,
                         @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1416
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1416,
                         @"[In SubResponseElementGenericType] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the cell storage service subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");
            }

            // When get exclusive lock succeeds, then record the current lock status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            // Set the time out value to the twice time out value of DefaultTimeOut.
            subRequest.SubRequestData.Timeout = System.Convert.ToString(SharedTestSuiteHelper.DefaultTimeOut * 2);

            // Get the exclusive lock again with same exclusive lock ID as the previous step and the twice time out value to refresh the time out of the exclusive lock, expect the error code "Success".
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If error codes returns success, the capture R1242.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1242,
                         @"[In Get Lock][If the file is locked with the same exclusive lock identifier that is sent in the exclusive lock subrequest of type ""Get lock"", the protocol server] returns an error code value set to ""Success"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock][If the file is locked with the same exclusive lock identifier that is sent in the exclusive lock subrequest of type ""Get lock"", the protocol server] returns an error code value set to ""Success"".");
            }

            // In this case, the exclusive lock will not be expired, because the second getting exclusive lock to refresh the time out value.
            SharedTestSuiteHelper.Sleep(65);

            ExclusiveLockSubRequestType checkAvailabilitySubrequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            checkAvailabilitySubrequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Check lock availability with different exclusive lock ID, expect the server returns error code "FileAlreadyLockedOnServer".
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { checkAvailabilitySubrequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the Check Lock Availability of ExclusiveLock sub request returns the FileAlreadyLockedOnServer status.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the check lock availability with different exclusive lock ID returns the FileAlreadyLockedOnServer, this indicates that the original exclusive lock is not expired. In this situation, the requirement R1241 can be captured.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1241,
                         @"[In Get Lock] If the file is locked with the same exclusive lock identifier that is sent in the exclusive lock subrequest of type ""Get lock"", the protocol server refreshes the existing exclusive lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] If the file is locked with the same exclusive lock identifier that is sent in the exclusive lock subrequest of type ""Get lock"", the protocol server refreshes the existing exclusive lock.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Get Lock of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC02_GetExclusiveLock_InvalidArgument()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Call Get Lock of exclusive sub request without ExclusiveLockID, expect the server responds error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements when get an exclusive lock on a file which is already locked with an exclusive lock with a different exclusive lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC03_GetExclusiveLock_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with all valid parameters, expect the server responds the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Get Lock of ExclusiveLock sub request succeeds.");

            // When get exclusive lock succeeds, then record the current lock status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Get the exclusive lock with different exclusive lock ID compared with the previous steps, expect the server responds error code "FileAlreadyLockedOnServer".
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP_R1239
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1239,
                         @"[In Get Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if one of the following conditions is true:
                         The file is already locked with an exclusive lock with a different exclusive lock identifier.");

                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP_R37801
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         37801,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileAlreadyLockedOnServer indicates an error when there is an already existing exclusive lock on the targeted URL for the file.");

                bool isR379Verified = exclusiveResponse.ErrorMessage.IndexOf(this.UserName01, 0, StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"When the ""FileAlreadyLockedOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the users who are currently holding the lock on the file in the ErrorMessage attribute, actual error message is {0}",
                    exclusiveResponse.ErrorMessage);

                // If the error message equals the value of UserName01, then capture MS-FSSHTTP_R379
                Site.CaptureRequirementIfIsTrue(
                         isR379Verified,
                         "MS-FSSHTTP",
                         379,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyLockedOnServer] When the ""FileAlreadyLockedOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the users who are currently holding the lock on the file in the ErrorMessage attribute.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if one of the following conditions is true:
                        The file is already locked with an exclusive lock with a different exclusive lock identifier.");

                bool isR379Verified = exclusiveResponse.ErrorMessage.IndexOf(this.UserName01, 0, StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"When the ""FileAlreadyLockedOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the users who are currently holding the lock on the file in the ErrorMessage attribute, actual error message is {0}",
                    exclusiveResponse.ErrorMessage);

                Site.Assert.IsTrue(
                    isR379Verified,
                    exclusiveResponse.ErrorMessage,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyLockedOnServer] When the ""FileAlreadyLockedOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the users who are currently holding the lock on the file in the ErrorMessage attribute.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get an exclusive lock on a file which is already locked with a shared lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC04_GetExclusiveLock_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the context.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a shared lock with all valid parameters, expect the server responds the error code "Success".
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with all valid parameters, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP_R1240
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1240,
                         @"[In Get Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if one of the following conditions is true: The file is already locked with a shared lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if one of the following conditions is true: The file is already locked with a shared lock.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get an exclusive lock on a file which is already check out by a different client.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC05_GetExclusiveLock_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock using the different user name comparing with the previous step for the already checked out file, expect the server responds the error code "FileAlreadyCheckedOutOnServer".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyCheckedOutOnServer", then capture MS-FSSHTTP_R1246
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1246,
                         @"[In Get Lock] If the protocol server encounters an issue in locking the file because a checkout has already been done by another client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                // If the error code equals "FileAlreadyCheckedOutOnServer", then capture MS-FSSHTTP_R384
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         384,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileAlreadyCheckedOutOnServer indicates an error when the file is checked out by another client, which is preventing the file from being locked by the current client.");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    exclusiveResponse.ErrorMessage);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R385
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] If the protocol server encounters an issue in locking the file because a checkout has already been done by another client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get an exclusive lock on a file in the case that a checkout needs to be done before the file is locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC06_GetExclusiveLock_DocumentCheckoutRequired()
        {
            // Set the library to a status in which the checkout operation is needed before the file is locked.
            if (!this.SutPowerShellAdapter.ChangeDocLibraryStatus(true))
            {
                this.Site.Assert.Fail("Cannot change a document library status to save a file to the document library that requires check out files.");
            }

            // Record the document library check out required status.
            this.StatusManager.RecordDocumentLibraryCheckOutRequired();

            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with all valid parameters, expect the server responds the error code "DocumentCheckoutRequired"
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "DocumentCheckoutRequired", then capture MS-FSSHTTP_R1243 and R1244
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1243,
                         @"[In Get Lock] If the protocol server encounters an error because the document is not being checked out on the server and the file is saved to a document library that requires checking out files, the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1244
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1244,
                         @"[In Get Lock] The ""DocumentCheckoutRequired"" error code value indicates to the protocol client that a checkout needs to be done before the file can be locked.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R363
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         363,
                         @"[In GenericErrorCodeTypes] DocumentCheckoutRequired indicates an error when the targeted URL for the file is not yet checked out by the current client before sending a lock request on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R364
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         364,
                         @"[In GenericErrorCodeTypes][DocumentCheckoutRequired] If the document is not checked out by the current client, the protocol server MUST return an error code value set to ""DocumentCheckoutRequired"" in the cell storage service response message.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DocumentCheckoutRequired,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] If the protocol server encounters an error because the document is not being checked out on the server and the file is saved to a document library that requires checking out files, the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");
            }
        }

        #endregion

        #region Release Lock

        /// <summary>
        /// A method used to verify the related requirements when the Release Lock of ExclusiveLock sub request succeeds. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC07_ReleaseExclusiveLock_Success()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock 
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Release the exclusive lock with the exclusive lock ID which has been locked by the previous step, expect the server responds the error code "Success".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ReleaseLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1236
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1236,
                         @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");

                bool isExclusiveLockExist = this.CheckExclusiveLockExist(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

                // The Check Lock Availability returns succeeds, then indicate the server really release the lock
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R1252, expect the exclusive lock session with ExclusiveLockID {0} on the file: {1} has been released, actually the exclusive lock is {2}.",
                    SharedTestSuiteHelper.DefaultExclusiveLockID,
                    this.DefaultFileUrl,
                    isExclusiveLockExist ? "not released" : "released");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1252
                Site.CaptureRequirementIfIsFalse(
                         isExclusiveLockExist,
                         "MS-FSSHTTP",
                         1252,
                         @"[In Release Lock] The protocol server releases the exclusive lock session on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");

                bool isExclusiveLockExist = this.CheckExclusiveLockExist(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);
                Site.Assert.IsFalse(
                    isExclusiveLockExist,
                    @"[In Release Lock] The protocol server releases the exclusive lock session on the file.");
            }

            this.StatusManager.CancelExclusiveLock(this.DefaultFileUrl);
        }

        /// <summary>
        /// A method used to verify the related requirements when the Release Lock of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC08_ReleaseExclusiveLock_InvalidArgument()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ReleaseLock);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Call release exclusive lock without ExclusiveLockID, expect the server responds the error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements when release an exclusive lock on a file which is not locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC09_ReleaseExclusiveLock_FileNotLockedOnServer()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ReleaseLock);

            // Release a file lock which is not locked, expect the server responds the error code "FileNotLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP_R1253
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1253,
                         @"[In Release Lock] If the protocol server encounters an error because no lock currently exists on the file, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");

                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP_R381
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         381,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileNotLockedOnServer indicates an error when no exclusive lock exists on a file and a release of the lock is requested as part of a cell storage service request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Release Lock] If the protocol server encounters an error because no lock currently exists on the file, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert an exclusive lock on a file which is not locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC43_ConvertExclusiveLock_FileNotLockedOnServer()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);

            // Convert a file lock which is not locked, expect the server responds the error code "FileNotLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP_R3811
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3811,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileNotLockedOnServer indicates an error when no exclusive lock on a file and a conversion of the lock is requested as part of a cell storage service request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert Lock] If the protocol server encounters an error because no lock currently exists on the file, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when release an exclusive lock on a file which is already locked with an exclusive lock with a different exclusive lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC11_ReleaseExclusiveLock_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get an exclusive lock ID with the default exclusive lock id
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ReleaseLock);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Release the exclusive lock with the different exclusive lock ID comparing with the previous step, expect the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1592
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1592,
                         @"[In Release Lock][Release Lock in ExclusiveLock Sub Request] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:
                         The file is already locked with an exclusive lock that has a different exclusive lock identifier.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Release Lock][Release Lock in ExclusiveLock Sub Request] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:
                       The file is already locked with an exclusive lock that has a different exclusive lock identifier.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when release an exclusive lock on a file which is already locked with a shared lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC12_ReleaseExclusiveLock_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a shared lock
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ReleaseLock);

            // Release the exclusive lock with all valid parameters, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1593
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1593,
                         @"[In Release Lock][Release Lock in ExclusiveLock Sub Request] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:
                         The file is already locked with a shared lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Release Lock][Release Lock in ExclusiveLock Sub Request] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:
                        The file is already locked with a shared lock.");
            }
        }

        #endregion

        #region Refresh Lock
        /// <summary>
        /// A method used to verify the related requirements when the Refresh Lock of ExclusiveLock sub request succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC10_RefreshExclusiveLock_Success()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock with timeout value 60 seconds
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID, 60);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);

            // Refresh the exclusive lock value twice than the initial timeout value, but less than 3600 seconds.
            subRequest.SubRequestData.Timeout = "120";

            // Refresh the exclusive lock with same exclusive lock ID as the previous step and the twice time out value, expect the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "Test case cannot continue unless the Refresh Lock of ExclusiveLock sub request succeeds.");
        }

        /// <summary>
        /// A method used to verify the related requirements when the Refresh Lock of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC13_RefreshExclusiveLock_InvalidArgument()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Call Get Lock of exclusive sub request without ExclusiveLockID, expect the server responds the error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh an exclusive lock on a file which is not locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC14_RefreshExclusiveLock_NoExclusiveLockExists()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);

            // Refresh the exclusive lock with all valid parameters when there is no lock on the file, expect the server responds the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Refresh Lock of ExclusiveLock sub request succeeds.");

            // Record the current file status
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            // Check lock availability with different exclusive lock ID as the previous step, expect the server responds the error code "Success".
            ExclusiveLockSubRequestType checkAvailabilitySubrequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            checkAvailabilitySubrequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { checkAvailabilitySubrequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the Check Lock Availability of ExclusiveLock sub request returns the FileAlreadyLockedOnServer status.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "Success", this means the refresh lock to get an exclusive lock, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1259
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1259,
                         @"[In Refresh Lock] If the refresh of the exclusive lock fails because no exclusive lock exists on the file, the protocol server gets a new exclusive lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If the refresh of the exclusive lock fails because no exclusive lock exists on the file, the protocol server gets a new exclusive lock on the file.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh an exclusive lock on a file which is already locked with an exclusive lock with a different exclusive lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC15_RefreshExclusiveLock_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock with the default exclusive lock.
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Refresh the exclusive lock with different exclusive lock ID as the previous step, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyLockedOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1261
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1261,
                         @"[In Refresh Lock][The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:] The protocol server is unable to refresh the lock on the file because an exclusive lock with a different exclusive lock identifier exists on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock][The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true:] The protocol server is unable to refresh the lock on the file because an exclusive lock with a different exclusive lock identifier exists on the file.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh an exclusive lock on a file which is already locked with a shared lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC16_RefreshExclusiveLock_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a shared lock
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);

            // Refresh the exclusive lock on the same file as the previous step, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyLockedOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1260
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1260,
                         @"[In Refresh Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true: The protocol server is unable to refresh the lock on the file because a shared lock already exists on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] The protocol server returns an error code value set to ""FileAlreadyLockedOnServer"" if any one of the following conditions is true: The protocol server is unable to refresh the lock on the file because a shared lock already exists on the file.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh an exclusive lock on a file in the case that a checkout needs to be done before the file is locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC17_RefreshExclusiveLock_DocumentCheckoutRequired()
        {
            // Set the library status a checkout operation needs to be done before the file is locked.
            if (!this.SutPowerShellAdapter.ChangeDocLibraryStatus(true))
            {
                this.Site.Assert.Fail("Cannot change a document library status to save a file to the document library that requires check out files.");
            }

            // Record the document library check out required status.
            this.StatusManager.RecordDocumentLibraryCheckOutRequired();

            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.RefreshLock);

            // Refresh the exclusive lock with all valid parameters, expect the server responses the error code "DocumentCheckoutRequired".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "DocumentCheckoutRequired", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1262, MS-FSSHTTP_R1263
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1262,
                         @"[In Refresh Lock] If the protocol server encounters an error because the document is not being checked out on the server and the file is saved to a document library that requires check out files, then the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1263
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1263,
                         @"[In Refresh Lock] ""DocumentCheckoutRequired"" error code value indicates to the protocol client that a checkout needs to be done before the file is to be locked and the lock is refreshed.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DocumentCheckoutRequired,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If the protocol server encounters an error because the document is not being checked out on the server and the file is saved to a document library that requires check out files, then the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh an exclusive lock on a file which has been checked out.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC18_RefreshExclusiveLock_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Refresh the exclusive lock with different user name comparing with the previous step for the already checked out file, expect the server responds the error code "FileAlreadyCheckedOutOnServer".
            // Now the service channel is initialized using the userName01 account by default.
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyCheckedOutOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1265
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1265,
                         @"[In Refresh Lock] If the protocol server encounters an error in locking the file because a checkout has already been done by another client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                // If the error code equals "FileAlreadyCheckedOutOnServer", then capture MS-FSSHTTP_R384
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         384,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileAlreadyCheckedOutOnServer indicates an error when the file is checked out by another client, which is preventing the file from being locked by the current client.");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    exclusiveResponse.ErrorMessage);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R385
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If the protocol server encounters an error in locking the file because a checkout has already been done by another client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
        }

        #endregion

        #region Check Lock Availability

        /// <summary>
        /// A method used to verify the related requirements when the Check Lock Availability of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC19_CheckExclusiveLockAvailability_InvalidArgument()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Check exclusive lock availability without ExclusiveLockID, expect the server responds the error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements when check an exclusive lock availability on a file which is already locked with an exclusive lock with a different exclusive lock identifier. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC20_CheckExclusiveLockAvailability_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock with the default exclusive lock id.
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Refresh the exclusive lock with the different exclusive lock ID comparing the previous step, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyLockedOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1310
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1310,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: If there is a current exclusive lock on the file with a different exclusive lock identifier than the one specified by the current client on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: If there is a current exclusive lock on the file with a different exclusive lock identifier than the one specified by the current client on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when check an exclusive lock availability on a file which is already locked with a shared lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC21_CheckExclusiveLockAvailability_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a shared lock
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);

            // Check the exclusive lock availability with all valid parameters, expect the server returns the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyLockedOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1954
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1954,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: or if there is a shared lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: or if there is a shared lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when check an exclusive lock availability on a file which is checkout by different user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC22_CheckExclusiveLockAvailability_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name. 
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check the exclusive lock availability using the different user name comparing with the previous step for the already checked out file, expect the server responses the error code "FileAlreadyCheckedOutOnServer".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code is "FileAlreadyCheckedOutOnServer", capture MS-FSSHTTP requirement: MS-FSSHTTP_R1494
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1494,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: If the file is checked out on the server, but it is checked out by a client with a different user name than that of the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    exclusiveResponse.ErrorMessage);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R385
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: If the file is checked out on the server, but it is checked out by a client with a different user name than that of the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                bool isVerifyR385 = exclusiveResponse.ErrorMessage != null && exclusiveResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
        }

        #endregion

        #region Convert to Schema Lock with Coauthoring Transition Tracked

        /// <summary>
        /// A method used to verify the related requirements when the Convert to Schema Lock with Coauthoring Transition Tracked of ExclusiveLock sub request succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC23_ConvertToSchemaJoinCoauth_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock with default exclusive lock id
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert current exclusive lock to a coauthoring shared lock, expect the server response the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1236
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1236,
                         @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In ExclusiveLock Subrequest] An ErrorCode value of ""Success"" indicates success in processing the exclusive lock request.");
            }

            // Now record the file status with coauth shared lock.
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Make sure the coauth status is returned by server.
            if (!exclusiveResponse.SubResponseData.CoauthStatusSpecified)
            {
                this.Site.Assert.Fail("Cannot get the coauth status when calling ConvertToSchemaJoinCoauth sub request.");
            }

            this.Site.Log.Add(
                LogEntryKind.Debug,
                "The CoauthStatus attribute should be specified if the request to convert the exclusive lock to a shared lock is processed successfully. Actually the attribute is {0}.",
                exclusiveResponse.SubResponseData.CoauthStatusSpecified ? "specified" : "not specified");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the coauth status is specified, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1272
                Site.CaptureRequirementIfIsTrue(
                         exclusiveResponse.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1272,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] After the request to convert the exclusive lock to a shared lock is processed successfully, the protocol server gets the coauthoring status and returns the status to the client.");

                // If the coauth status is specified, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1426
                Site.CaptureRequirementIfIsTrue(
                         exclusiveResponse.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1426,
                         @"[In SubResponseDataOptionalAttributes][CoauthStatus] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the CoauthStatus attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: An exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"".");

                // If the coauth status is specified, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1476
                Site.CaptureRequirementIfIsTrue(
                         exclusiveResponse.SubResponseData.CoauthStatusSpecified,
                         "MS-FSSHTTP",
                         1476,
                         @"[In ExclusiveLockSubResponseDataType] CoauthStatus MUST be specified in an exclusive lock subresponse that is generated in response to an exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"" if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                // If the coauth status is specified and equals Alone, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1283
                Site.CaptureRequirementIfAreEqual<CoauthStatusType>(
                         CoauthStatusType.Alone,
                         exclusiveResponse.SubResponseData.CoauthStatus,
                         "MS-FSSHTTP",
                         1283,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] If the current client is the only client editing the file, the protocol server MUST return a CoauthStatus attribute set to ""Alone"", which indicates that no one else is editing the file.");

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R652, the TransitionID cannot be null.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R652
                Site.CaptureRequirementIfIsNotNull(
                         exclusiveResponse.SubResponseData.TransitionID,
                         "MS-FSSHTTP",
                         652,
                         @"[In ExclusiveLockSubResponseDataType] TransitionID MUST be returned as part of the response for an exclusive lock subrequest of type ""Convert to Schema lock with coauthoring transition tracked"" if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                bool isInCoauthSession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The client id {0} should in the coauth session with schema lock id {1} after convert to schema join coauth.",
                    SharedTestSuiteHelper.DefaultClientID,
                    SharedTestSuiteHelper.ReservedSchemaLockID);

                // If there is coauth session exist, then capture requirement MS-FSSHTTP_R1594 and MS-FSSHTTP_R1595
                Site.CaptureRequirementIfIsTrue(
                         isInCoauthSession,
                         "MS-FSSHTTP",
                         1594,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] When the protocol server receives this subrequest, it does all of the following:
                         Converts the exclusive lock on the file to a shared lock.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1595
                Site.CaptureRequirementIfIsTrue(
                         isInCoauthSession,
                         "MS-FSSHTTP",
                         1595,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] When the protocol server receives this subrequest, it does all of the following:
                         Starts a coauthoring session for the file if one is not already present and adds the client to that session.");
            }
            else
            {
                Site.Assert.IsTrue(
                    exclusiveResponse.SubResponseData.CoauthStatusSpecified,
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] After the request to convert the exclusive lock to a shared lock is processed successfully, the protocol server gets the coauthoring status and returns the status to the client.");

                Site.Assert.AreEqual<CoauthStatusType>(
                    CoauthStatusType.Alone,
                    exclusiveResponse.SubResponseData.CoauthStatus,
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] If the current client is the only client editing the file, the protocol server MUST return a CoauthStatus attribute set to ""Alone"", which indicates that no one else is editing the file.");

                Site.Assert.IsNotNull(
                    exclusiveResponse.SubResponseData.TransitionID,
                    @"[In ExclusiveLockSubResponseDataType] TransitionID MUST be returned as part of the response for an exclusive lock subrequest of type ""Convert to Schema lock with coauthoring transition tracked"" if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                bool isInCoauthSession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);
                Site.Assert.IsTrue(
                    isInCoauthSession,
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] When the protocol server receives this subrequest, it does all of the following:
                        Converts the exclusive lock on the file to a shared lock.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Convert to Schema Lock with Coauthoring Transition Tracked of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC24_ConvertToSchemaJoinCoauth_InvalidArgument()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Call Get Lock of exclusive sub request without ExclusiveLockID
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Convert exclusive lock to a coauth shared lock without ExclusiveLockID, expect the server response the error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements whether another client can share the lock after finish converting an exclusive lock to the coauth lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC25_ConvertToSchemaJoinCoauth_SharedLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert current exclusive lock to a coauthoring shared lock with a specified client ID and schema lock ID, expect the server response the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                                ErrorCodeType.Success,
                                SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                                "Test case cannot continue unless the ConvertToSchemaJoinCoauth of ExclusiveLock sub request succeeds.");

            // Now record the file status with coauthoring shared lock.
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            SchemaLockSubRequestType getSharedLockSubRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);
            getSharedLockSubRequest.SubRequestData.ClientID = System.Guid.NewGuid().ToString();

            // Get shared lock with the same schema lock ID as the previous step, but with a new different client ID, expect the server returns the error code "Success".
            SchemaLockSubResponseType getSharedLockSubresponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getSharedLockSubRequest }), 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server response error code Success, then capture requirements R1486 and R1485.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1486,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server uses the SchemaLockId attribute sent in an exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"" to ensure that after the exclusive lock on the file is converted to a shared lock, the protocol server MUST allow other clients with the same schema lock identifier to share the lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1485
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1485,
                         @"[In ExclusiveLockSubRequestDataType] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server uses the SchemaLockId attribute sent in an exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"" to ensure that after the exclusive lock on the file is converted to a shared lock, the protocol server MUST allow other clients with the same schema lock identifier to share the lock on the file.");
            }

            // Now record the file status with coauthoring shared lock.
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, getSharedLockSubRequest.SubRequestData.ClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get shared lock with the different schema lock ID and a new different client ID as the previous step, expect the server returns the error code "Success".
            getSharedLockSubRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);
            getSharedLockSubRequest.SubRequestData.ClientID = System.Guid.NewGuid().ToString();
            getSharedLockSubRequest.SubRequestData.SchemaLockID = System.Guid.NewGuid().ToString();
            getSharedLockSubresponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getSharedLockSubRequest }), 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If fail when using different schema lock ID, then capture requirement R1371 and R1373.
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1371,
                         @"[In ExclusiveLockSubRequestDataType] The schema lock identifier is used by the protocol server to block other clients that have different schema identifiers.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1373
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1373,
                         @"[In ExclusiveLockSubRequestDataType] The protocol server ensures that at any instant in time, only clients having the same schema lock identifier can lock the document.");
            }
            else
            {
                Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                    @"[In ExclusiveLockSubRequestDataType] The schema lock identifier is used by the protocol server to block other clients that have different schema identifiers.");
            }

            // Release two shared lock on the default file.
            if (!this.StatusManager.RollbackStatus())
            {
                this.Site.Assert.Fail("Failed to release the two shared lock on the file {0}", this.DefaultFileUrl);
            }

            // Get shared lock with the different schema lock ID and a new different client ID as the previous steps when all the shared locks are released on the file, expect the server returns the error code "Success".
            getSharedLockSubresponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { getSharedLockSubRequest }), 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals Success, then capture requirement R1374
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1374,
                         @"[In ExclusiveLockSubRequestDataType] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getSharedLockSubresponse.ErrorCode, this.Site),
                    @"[In ExclusiveLockSubRequestDataType] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
            }

            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, getSharedLockSubRequest.SubRequestData.ClientID, getSharedLockSubRequest.SubRequestData.SchemaLockID);
        }

        /// <summary>
        /// A method used to test the related requirements convert the exclusive lock to a coauthoring shared when the coauthoring feature is disabled.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC26_ConvertToSchemaJoinCoauth_LockNotConvertedAsCoauthDisabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Set the library turn off the coauthoring feature.
            if (!this.SutPowerShellAdapter.SwitchCoauthoringFeature(true))
            {
                this.Site.Assert.Fail("Cannot disable the coauthoring feature");
            }

            // Record the disable the coauthoring feature status.
            this.StatusManager.RecordDisableCoauth();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            ExclusiveLockSubResponseType exclusiveResponse = null;

            while (retryCount > 0)
            {
                ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

                // Convert the current exclusive lock to a coauthoring shared lock when the coauthoring feature is disabled, expect the server returns the error code "LockNotConvertedAsCoauthDisabled".
                CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site) == ErrorCodeType.LockNotConvertedAsCoauthDisabled)
                {
                    break;
                }

                retryCount--;
                if (retryCount == 0)
                {
                    Site.Assert.Fail("Error LockNotConvertedAsCoauthDisabled should be returned if coauthoring feature is disabled.");
                }

                System.Threading.Thread.Sleep(waitTime);
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals FileNotLockedOnServerAsCoauthDisabled, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1278
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1278,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the coauthoring feature is disabled by the protocol server, the protocol server returns an error code value set to ""LockNotConvertedAsCoauthDisabled"".");

                // If the error code equals LockNotConvertedAsCoauthDisabled, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R383
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         383,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] LockNotConvertedAsCoauthDisabled indicates an error when a protocol server fails to process a lock conversion request sent as part of a cell storage service request because coauthoring of the file is disabled on the server.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the coauthoring feature is disabled by the protocol server, the protocol server returns an error code value set to ""LockNotConvertedAsCoauthDisabled"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to coauthoring lock on a file which is already locked with an exclusive lock with a different exclusive lock identifier.  
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC27_ConvertToSchemaJoinCoauth_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();

            // Convert an exclusive lock to a coauthoring shared lock with the different exclusive lock id as the previous step, expect the server returns the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals FileAlreadyLockedOnServer, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1596
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1596,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because either there is an exclusive lock with a different exclusive lock identifier already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because either there is an exclusive lock with a different exclusive lock identifier already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to coauthoring lock on a file which is already locked with a shared lock. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC28_ConvertToSchemaJoinCoauth_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a shared lock
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert an exclusive lock to a coauthoring shared lock, expect the server returns the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1597
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1597,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] If the protocol server is unable to convert the lock because] there is a shared lock already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] If the protocol server is unable to convert the lock because] there is a shared lock already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to coauthoring lock on a file which is not locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC29_ConvertToSchemaJoinCoauth_FileNotLockedOnServer()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Convert an exclusive lock to a coauthoring shared lock on a file which is not locked, expect the server returns the error code "FileNotLockedOnServer".
            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1280
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1280,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock on the file because no lock exists on the server, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");

                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R3813
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3813,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileNotLockedOnServer indicates an error when no exclusive lock on a file and a conversion of the lock is requested as part of a cell storage service request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock on the file because no lock exists on the server, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert an exclusive lock to coauthoring lock on a file in the case that a checkout needs to be done before the file is locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC30_ConvertToSchemaJoinCoauth_DocumentCheckoutRequired()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Set the library status a checkout operation needs to be done before the file is locked.
            if (!this.SutPowerShellAdapter.ChangeDocLibraryStatus(true))
            {
                this.Site.Assert.Fail("Cannot change a document library status to save a file to the document library that requires check out files.");
            }

            // Record the document library check out required status.
            this.StatusManager.RecordDocumentLibraryCheckOutRequired();

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert the exist exclusive lock to a coauthoring shared lock, expect the server returns the error code "DocumentCheckoutRequired".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server response error code DocumentCheckoutRequired, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1598, MS-FSSHTTP_R1941 
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1598,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because the document is saved to a document library that requires checking out files and the document is not checked out on the server, the protocol server returns an error code value set to ""DocumentCheckoutRequired"". ");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1941
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1941,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The ""DocumentCheckoutRequired"" error code value indicates to the protocol client that a checkout needs to be done before the exclusive lock is converted to a shared lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DocumentCheckoutRequired,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because the document is saved to a document library that requires checking out files and the document is not checked out on the server, the protocol server returns an error code value set to ""DocumentCheckoutRequired"". ");
            }
        }

        /// <summary>
        /// A method used aims to verify the response has unique TransitionID when use different files.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC31_ConvertToSchemaJoinCoauth_UniqueTransitionID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock on the default file
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert an exclusive lock to a coauthoring shared lock on the first file, expect the server returns the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless ConvertToSchemaJoinCoauth of exclusive lock sub request succeeds.");

            // Record the file coauthoring lock status
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Store the first transition ID.
            string transitionId1 = exclusiveResponse.SubResponseData.TransitionID;

            // Prepare another file.
            string anotherFileUrl = this.PrepareFile(); 

            // Prepare an exclusive lock on the another existent file
            this.PrepareExclusiveLock(anotherFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);

            // Convert an exclusive lock to a coauthoring shared lock on the second file, expect the server returns the error code "Success".
            response = this.Adapter.CellStorageRequest(anotherFileUrl, new SubRequestType[] { subRequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless ConvertToSchemaJoinCoauth of exclusive lock sub request succeeds.");

            // Record the file coauthoring lock status
            this.StatusManager.RecordCoauthSession(anotherFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Store the second transition ID.
            string transitionId2 = exclusiveResponse.SubResponseData.TransitionID;

            // When call ConvertToSchemaJoinCoauthInExclusiveLockSubRequest twice use different files,
            // the server returns different transitionIDs.
            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1487
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreNotEqual<string>(
                         transitionId1,
                         transitionId2,
                         "MS-FSSHTTP",
                         1487,
                         @"[In ExclusiveLockSubResponseDataType] TransitionID: A guid specifies that if 2 requests are operated on 2 different files in the protocol server, the TransitionID values returned in the 2 corresponding responses are different.");
            }
            else
            {
                Site.Assert.AreNotEqual<string>(
                    transitionId1,
                    transitionId2,
                    @"[In ExclusiveLockSubResponseDataType] TransitionID: A guid specifies that if 2 requests are operated on 2 different files in the protocol server, the TransitionID values returned in the 2 corresponding responses are different.");
            }
        }

        #endregion

        #region Convert to Schema Lock

        /// <summary>
        /// A method used to verify the related requirements when the Convert to Schema Lock of ExclusiveLock sub request succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC32_ConvertToSchema_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock on the default file
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert current exist exclusive lock to a schema shared lock, expect the server returns the error code "Success".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the ConvertToSchema of exclusive lock succeeds.");

            // Now record the file status with exclusive lock.
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            bool isSchemaLockExist = this.CheckSchemaLockExist(this.DefaultFileUrl, subRequest.SubRequestData.SchemaLockID);
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "After convert to exclusive lock {1} to schema lock with the ID {0}, this schema lock should exist",
                subRequest.SubRequestData.SchemaLockID,
                SharedTestSuiteHelper.DefaultExclusiveLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1290
                Site.CaptureRequirementIfIsTrue(
                         isSchemaLockExist,
                         "MS-FSSHTTP",
                         1290,
                         @"[In Convert to Schema Lock] The protocol server process the request[exclusive lock sub request of type, ""Convert to schema lock""] by converting an exclusive lock on the file to a shared lock.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isSchemaLockExist,
                    @"[In Convert to Schema Lock] The protocol server process the request[exclusive lock sub request of type, ""Convert to schema lock""] by converting an exclusive lock on the file to a shared lock.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Convert to Schema Lock of ExclusiveLock sub request does not include ExclusiveLockID or ExclusiveLockRequestType attribute.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC33_ConvertToSchema_InvalidArgument()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);
            subRequest.SubRequestData.ExclusiveLockID = null;

            // Convert an exclusive lock to a schema shared lock without providing ExclusiveLockID, expect the server returns the error code "InvalidArgument".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.CheckInvalidArgumentRelatedRequirements(exclusiveResponse);
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to schema lock on a file which is already locked with an exclusive lock with a different exclusive lock identifier.  
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC34_ConvertToSchema_FileAlreadyLockedOnServer_Case1()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock with same user name on the checked out file.
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert an exclusive lock to a schema shared lock with different exclusive lock ID as the previous step, expect the server returns the error code "FileAlreadyLockedOnServer".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);
            subRequest.SubRequestData.ExclusiveLockID = System.Guid.NewGuid().ToString();
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FileAlreadyLockedOnServer", then capture MS-FSSHTTP_R1300
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1300,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because either there is an exclusive lock with a different exclusive lock identifier already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because either there is an exclusive lock with a different exclusive lock identifier already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to schema lock on a file which is already locked with a shared lock. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC35_ConvertToSchema_FileAlreadyLockedOnServer_Case2()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a shared lock
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);

            // Convert a different exclusive lock to a schema shared lock, expect the server returns the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1599
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1599,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because there is a shared lock already present on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because there is an exclusive lock with a different exclusive lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements convert an exclusive lock to a schema lock on a file which is not locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC36_ConvertToSchema_FileNotLockedOnServer()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);

            // Convert an exclusive lock to an schema lock on a file which is not locked, expect the server returns the error code "FileNotLockedOnServer".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1301
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1301,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock on the file because no lock exists on the server, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");

                // If the error code equals "FileNotLockedOnServer", then capture MS-FSSHTTP requirement: MS-FSSHTTP_R3813
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3813,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileNotLockedOnServer indicates an error when no exclusive lock on a file and a conversion of the lock is requested as part of a cell storage service request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileNotLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock on the file because no lock exists on the server, the protocol server returns an error code value set to ""FileNotLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert an exclusive lock to schema lock on a file in the case that a checkout needs to be done before the file is locked.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC37_ConvertToSchema_DocumentCheckoutRequired()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Set the library to a status in which a checkout operation needs to be done before the file is locked.
            if (!this.SutPowerShellAdapter.ChangeDocLibraryStatus(true))
            {
                this.Site.Assert.Fail("Cannot change a document library status to save a file to the document library that requires check out files.");
            }

            // Record the document library check out required status.
            this.StatusManager.RecordDocumentLibraryCheckOutRequired();

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);

            // Convert the current existed exclusive lock to a schema lock, expect the server returns the error code "DocumentCheckoutRequired".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server response error code DocumentCheckoutRequired, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1302, MS-FSSHTTP_R1303 
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1302,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because the document is saved to a document library that requires checking out files and the document is not checked out on the server, the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1303
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.DocumentCheckoutRequired,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1303,
                         @"[In Convert to Schema Lock] The ""DocumentCheckoutRequired"" error code value indicates to the protocol client that a checkout needs to be done before the exclusive lock is converted to a shared lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.DocumentCheckoutRequired,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the lock because the document is saved to a document library that requires checking out files and the document is not checked out on the server, the protocol server returns an error code value set to ""DocumentCheckoutRequired"".");
            }
        }

        /// <summary>
        /// A method used to test the related requirements when convert the exclusive lock to a schema shared when the coauthoring  feature is disabled.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC38_ConvertToSchema_LockNotConvertedAsCoauthDisabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Set the library turn off the coauthoring feature.
            if (!this.SutPowerShellAdapter.SwitchCoauthoringFeature(true))
            {
                this.Site.Assert.Fail("Cannot disable the coauthoring feature");
            }

            // Record the disable the coauthoring feature status.
            this.StatusManager.RecordDisableCoauth();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            ExclusiveLockSubResponseType exclusiveResponse = null;

            while (retryCount > 0)
            {
                ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);

                // Convert the current exclusive lock to a schema shared lock when the coauthoring  feature is disabled, expect the server returns the error code "LockNotConvertedAsCoauthDisabled".
                CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site) == ErrorCodeType.LockNotConvertedAsCoauthDisabled)
                {
                    break;
                }

                retryCount--;
                if (retryCount == 0)
                {
                    Site.Assert.Fail("Error LockNotConvertedAsCoauthDisabled should be returned if feature of coauthoring is not completely supported.");
                }

                System.Threading.Thread.Sleep(waitTime);
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals LockNotConvertedAsCoauthDisabled, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1298
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1298,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the feature of coauthoring is not completely supported by the protocol server, the protocol server returns an error code value set to ""LockNotConvertedAsCoauthDisabled"".");

                // If the error code equals LockNotConvertedAsCoauthDisabled, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R383
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         383,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] LockNotConvertedAsCoauthDisabled indicates an error when a protocol server fails to process a lock conversion request sent as part of a cell storage service request because coauthoring of the file is disabled on the server.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the feature of coauthoring is not completely supported by the protocol server, the protocol server returns an error code value set to ""LockNotConvertedAsCoauthDisabled"".");
            }
        }
        #endregion

        #region Common

        /// <summary>
        /// A method used to test that the ErrorCode attribute is present in the response when the value of the attribute RequestVersion in the subRequest is less than 2. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC39_ExclusiveLock_RequestVersionLessThanTwo()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get an exclusive lock on a file with a Version attribute less than 2 and expect the ErrorCode in the response is present.
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest }, SequenceNumberGenerator.GetCurrentToken().ToString(), 1, 1);

            // If ErrorCode is not null, MS-FSSHTTP_R15141 should be covered.
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "For MS-FSSHTTP_R15141, The ErrorCode attribute should be present when the Version attribute is less than 2. Actually the ErrorCode attribute is {0}.",
                response.ResponseVersion.ErrorCodeSpecified);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R15141
                Site.CaptureRequirementIfIsTrue(
                         response.ResponseVersion.ErrorCodeSpecified,
                         "MS-FSSHTTP",
                         15141,
                         @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if the following is true:
                         The Version attribute of the RequestVersion element of the request message has a value that is less than 2.");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.ResponseVersion.ErrorCodeSpecified,
                    @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if the following is true:
                        The Version attribute of the RequestVersion element of the request message has a value that is less than 2.");
            }
        }

        /// <summary>
        /// A method used to verify the ErrorCode attribute is present if the conditions are satisfied.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC40_ExclusiveLock_ErrorCodeNotPresentOnResponseElement()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site), "The test case cannot continue until the user {0} get the exclusive lock {1} on the file {2}", this.UserName01, SharedTestSuiteHelper.DefaultExclusiveLockID, this.DefaultFileUrl);
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2269
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "For MS-FSSHTTP_R2269, the ErrorCode attribute should not be present, actually it {0}.",
                response.ResponseCollection.Response[0].ErrorCodeSpecified ? "present" : "not present");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2269
                Site.CaptureRequirementIfIsFalse(
                         response.ResponseCollection.Response[0].ErrorCodeSpecified,
                         "MS-FSSHTTP",
                         2269,
                         @"[In Response] This attribute[ErrorCode] MUST NOT be present if all of the followings are true: 
                         1.The Url attribute of the corresponding Request element does exist; 
                         2.The Url attribute of the corresponding Request element is not an empty string; 
                         3.The RequestToken attribute of the corresponding Request element does exist; 
                         4.The RequestToken attribute of the corresponding Request element is not an empty string; 
                         5.No exceptions occurred during the processing of a subrequest that was not entirely handled by the subrequest processing logic.");

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R2268, the ErrorCode attribute in the ResponseVersion should not be present, actually it {0}.",
                    response.ResponseVersion.ErrorCodeSpecified ? "present" : "not present");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2268
                Site.CaptureRequirementIfIsFalse(
                         response.ResponseVersion.ErrorCodeSpecified,
                         "MS-FSSHTTP",
                         2268,
                         @"[In ResponseVersion] This attribute[ErrorCode] MUST NOT be present if all of the followings are true: 
                         1.The RequestVersion element is not missing from the Body element of the SOAP request message; 
                         2.The Version attribute of the RequestVersion element of the request message has not a value that is less than 2; 
                         3.The protocol server identified by the WebUrl attribute of the ResponseCollection element does exist; 
                         4.The protocol server identified by the WebUrl attribute of the ResponseCollection element is available; 
                         5.The user does have permission to issue a cell storage service request to the file identified by the Url attribute of the Request element; 
                         6.This protocol is enabled on the protocol server.");
            }
            else
            {
                Site.Assert.IsFalse(
                    response.ResponseCollection.Response[0].ErrorCodeSpecified,
                    @"[In Response] This attribute[ErrorCode] MUST NOT be present if all of the followings are true: 
                    1.The Url attribute of the corresponding Request element does exist; 
                    2.The Url attribute of the corresponding Request element is not an empty string; 
                    3.The RequestToken attribute of the corresponding Request element does exist; 
                    4.The RequestToken attribute of the corresponding Request element is not an empty string; 
                    5.No exceptions occurred during the processing of a subrequest that was not entirely handled by the subrequest processing logic.");

                Site.Assert.IsFalse(
                    response.ResponseVersion.ErrorCodeSpecified,
                    @"[In ResponseVersion] This attribute[ErrorCode] MUST NOT be present if all of the followings are true: 
                        1.The RequestVersion element is not missing from the Body element of the SOAP request message; 
                        2.The Version attribute of the RequestVersion element of the request message has not a value that is less than 2; 
                        3.The protocol server identified by the WebUrl attribute of the ResponseCollection element does exist; 
                        4.The protocol server identified by the WebUrl attribute of the ResponseCollection element is available; 
                        5.The user does have permission to issue a cell storage service request to the file identified by the Url attribute of the Request element; 
                        6.This protocol is enabled on the protocol server.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when version set as not supported value.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC41_ExclusiveLock_IncompatibleVersion()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get an exclusive lock with the version set to 1 and minorVersion set to 5, expect server returns the error code "IncompatibleVersion".
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest }, null, 1, 5, null, null);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R88
                Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                         GenericErrorCodeTypes.IncompatibleVersion,
                         response.ResponseVersion.ErrorCode,
                         "MS-FSSHTTP",
                         88,
                         @"[In RequestVersion] Errors that occur because a version is not supported cause an IncompatibleVersion error code value to be set.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R89
                Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                         GenericErrorCodeTypes.IncompatibleVersion,
                         response.ResponseVersion.ErrorCode,
                         "MS-FSSHTTP",
                         89,
                         @"[In RequestVersion] [Errors that occur because a version is not supported cause an IncompatibleVersion error code value to be] sent as part of the ResponseVersion element.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R356
                Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                         GenericErrorCodeTypes.IncompatibleVersion,
                         response.ResponseVersion.ErrorCode,
                         "MS-FSSHTTP",
                         356,
                         @"[In GenericErrorCodeTypes] IncompatibleVersion indicates an error when any an incompatible version number is specified as part of the RequestVersion element of the cell storage service.");
            }
            else
            {
                Site.Assert.AreEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.IncompatibleVersion,
                    response.ResponseVersion.ErrorCode,
                    @"[In RequestVersion] Errors that occur because a version is not supported cause an IncompatibleVersion error code value to be set.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns ErrorCode "InvalidArgument" or "HighLevelExceptionThrown" when the attribute ExclusiveLockRequestType is not provided in the ExclusiveLock request.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S04_TC42_ExclusiveLock_MissingExclusiveLockRequestType()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create an exclusive sub request without ExclusiveLockRequestType.
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            subRequest.SubRequestData.ExclusiveLockRequestTypeSpecified = false;

            // Call an exclusive sub request without ExclusiveLockRequestType, expect the server responds with error code "InvalidArgument" or "HighLevelExceptionThrown".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3071
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3071, this.Site))
                {
                    ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidArgument,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3071,
                             @"[In Appendix B: Product Behavior] If the specified attributes[ExclusiveLockRequestType attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the ResponseVersion element. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3072
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3072, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.HighLevelExceptionThrown,
                             response.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             3072,
                             @"[In Appendix B: Product Behavior] The implementation does return a ""HighLevelExceptionThrown"" error code as part of the ResponseVersion element. <22> Section 2.3.1.9: In SharePoint Server 2013 [and Microsoft SharePoint Foundation 2013], if the ExclusiveLockRequestType attribute is not provided, a ""HighLevelExceptionThrown"" error code MUST be returned as part of the ResponseVersion element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3071, this.Site))
                {
                    ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidArgument,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] If the specified attributes[ExclusiveLockRequestType attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the exclusive lock subresponse. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3072, this.Site))
                {
                    Site.Assert.AreEqual<GenericErrorCodeTypes>(
                        GenericErrorCodeTypes.HighLevelExceptionThrown,
                        response.ResponseVersion.ErrorCode,
                        @"[In Appendix B: Product Behavior] The implementation does return a ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the exclusive lock subresponse. <15> Section 2.3.1.9: In SharePoint Server 2013 [and Microsoft SharePoint Foundation 2013], if the ExclusiveLockRequestType attribute is not provided, a ""HighLevelExceptionThrown"" error code MUST be returned as part of the SubResponseData element associated with the exclusive lock subresponse.");
                }
            }
        }
        #endregion

        #endregion

        /// <summary>
        /// A method used to verify the invalid argument related requirements.
        /// </summary>
        /// <param name="response">The ExclusiveLockSubResponseType instance returns from the server.</param>
        private void CheckInvalidArgumentRelatedRequirements(ExclusiveLockSubResponseType response)
        {
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R6341
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         6341,
                         @"[In ExclusiveLockSubRequestDataType] If the specified attributes[ExclusiveLockID attribute] are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the exclusive lock subresponse.<22>");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                     ErrorCodeType.InvalidArgument,
                     SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                     @"[In ExclusiveLockSubRequestDataType] If the specified attributes[ExclusiveLockID attribute] are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the exclusive lock subresponse.<15>");
            }
        }
    }
}