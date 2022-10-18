namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with SchemaLock operation.
    /// </summary>
    [TestClass]
    public abstract class S03_SchemaLock : SharedTestSuiteBase
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
        public void S03_SchemaLockInitialization()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93102, this.Site))
            {
                this.Site.Assume.Inconclusive("Implementation does not support the schema lock subrequest.");
            }

            // Initialize the default file URL, for this scenario, the target file URL should be unique for each test case
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases

        #region Check lock availability

        /// <summary>
        /// A method used to verify the related requirements when check a schema lock availability on a file which is checkout by a different user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC01_CheckLockAvailability_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name.
            bool isCheckOutSuccess = SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            Site.Assert.AreEqual(true, isCheckOutSuccess, "Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Check schema lock availability with all valid parameters, expect the server returns the error code "FileAlreadyCheckedOutOnServer".
            // Now the web service is initialized using the user01, so the user is different with the user who check the file out.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1208, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1208
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyCheckedOutOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1208,
                             @"[In Check Lock Availability] If the coauthorable file is checked out on the server and it is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1208, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] If the coauthorable file is checked out on the server and it is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when check a schema lock availability on a file which is already locked with an exclusive lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC02_CheckLockAvailability_FileAlreadyLockedOnServer_CurrentExclusiveLock()
        {
            // Get an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Check schema lock availability with all valid parameters, expect the server returns the error code "FileAlreadyLockedOnServer".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, false);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1206
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1206,
                         @"[In Check Lock Availability] If there is a current exclusive lock[ on the file], the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] If there is a current exclusive lock[ on the file], the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when checking a schema lock availability on a file is already locked by a schema lock with a different schemaLockID.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC03_CheckLockAvailability_FileAlreadyLockedOnServer_DifferentSchemaLockId()
        {
            // Prepare a schema lock.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Check schema lock availability with a different schemaLockID comparing with the previous step, expect the server returns the error code "FileAlreadyLockedOnServer".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
            subRequest.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1207
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1207,
                         @"[In Check Lock Availability][If there is a] shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         37802,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] FileAlreadyLockedOnServer indicates an error when there is an already existing schema lock on the file with a different schema lock identifier.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability][If there is a] shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Check Lock Availability of schema lock subRequest succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC04_CheckLockAvailability_Success()
        {
            // Prepare a schema lock.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Check schema lock availability with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1210
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1210,
                         @"[In Check Lock Availability] In all other cases[conditions except: 1. there is a current exclusive lock or shared lock on the file with a different schema lock identifier. 2.  the coauthorable file is checked out on the server and is checked out by a client with a different user name than the current client or is checked out by the current client], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] In all other cases[conditions except: 1. there is a current exclusive lock or shared lock on the file with a different schema lock identifier. 2.  the coauthorable file is checked out on the server and is checked out by a client with a different user name than the current client or is checked out by the current client], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }
        }

        #endregion

        #region Convert to Exclusive Lock

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive lock with ReleaseLockOnConversionToExclusiveFailure attribute set to true and multiple clients are in the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC05_ConvertToExclusiveLock_ExitCoauthSessionAsConvertToExclusiveFailed()
        {
            // the user01 to prepare a schema lock on the default file
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // the user02 to prepare a shared schema lock with different client id on the default file.
            this.PrepareSchemaLock(this.DefaultFileUrl, System.Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Convert the schema lock to an exclusive lock with ReleaseLockOnConversionToExclusiveFailure set to true and the same clientID with the first client, expect the server returns the error code "ExitCoauthSessionAsConvertToExclusiveFailed".
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, true);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1197 and MS-FSSHTTP_R395
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1197,
                         @"[In Convert to Exclusive Lock] The protocol server returns an error code value set to ""ExitCoauthSessionAsConvertToExclusiveFailed"" when the following conditions are both true:
1.The ReleaseLockOnConversionToExclusiveFailure attribute is set to true in the coauthoring subrequest;
2.Multiple clients are in the coauthoring session.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ExitCoauthSessionAsConvertToExclusiveFailed,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         395,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] ExitCoauthSessionAsConvertToExclusiveFailed indicates an error when a coauthoring subrequest or schema lock subrequest of type ""Convert to exclusive lock"" is sent by the client with the ReleaseLockOnConversionToExclusiveFailure attribute set to true, and there is more than one client editing the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R444
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         444,
                         @"[In SubRequestDataOptionalAttributes] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker when all of the following conditions are true:
Either the type of coauthoring subrequest is ""Convert to an exclusive lock"" or the type of the schema lock subrequest is ""Convert to an Exclusive Lock"".
The conversion to an exclusive lock failed.");

                bool isInSession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The user {0} with client ID {0} and schema Lock ID {1} should not exist in the coauthoring session, but actual it {2}",
                    this.UserName01,
                    SharedTestSuiteHelper.DefaultClientID,
                    SharedTestSuiteHelper.ReservedSchemaLockID,
                    isInSession ? "exists" : "does not exist");
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1198
                    Site.CaptureRequirementIfIsFalse(
                         isInSession,
                         "MS-FSSHTTP",
                         1198,
                         @"[In Convert to Exclusive Lock] When the ReleaseLockOnConversionToExclusiveFailure attribute is set to true and the conversion to an exclusive lock failed, the protocol server removes the client from the coauthoring session on the file.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R678
                    Site.CaptureRequirementIfIsFalse(
                             isInSession,
                             "MS-FSSHTTP",
                             678,
                             @"[In SchemaLockSubRequestDataType] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker, provided that all of the following conditions are true:
The type of the schema lock subrequest is ""Convert to an Exclusive Lock"";
The conversion to an exclusive lock failed.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R679
                    Site.CaptureRequirementIfIsFalse(
                             isInSession,
                             "MS-FSSHTTP",
                             679,
                             @"[In SchemaLockSubRequestDataType] When all the preceding conditions[1. The type of the schema lock subrequest is ""Convert to an Exclusive Lock"". 2. The conversion to an exclusive lock failed.] are true, the following apply:
A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of true indicates that the protocol server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R445
                    Site.CaptureRequirementIfIsFalse(
                             isInSession,
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
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Convert to Exclusive Lock] The protocol server returns an error code value set to ""ExitCoauthSessionAsConvertToExclusiveFailed"" when the following conditions are both true:
1.The ReleaseLockOnConversionToExclusiveFailure attribute is set to true in the coauthoring subrequest;
2.Multiple clients are in the coauthoring session.");

                Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SubRequestDataOptionalAttributes] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker when all of the following conditions are true:
Either the type of coauthoring subrequest is ""Convert to an exclusive lock"" or the type of the schema lock subrequest is ""Convert to an Exclusive Lock"".
The conversion to an exclusive lock failed.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
                {
                    bool isInSession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);
                    Site.Assert.IsFalse(
                        isInSession,
                        @"[In Convert to Exclusive Lock] When the ReleaseLockOnConversionToExclusiveFailure attribute is set to true and the conversion to an exclusive lock failed, the protocol server removes the client from the coauthoring session on the file.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive lock on a file which is already locked with an exclusive lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC06_ConvertToExclusiveLock_InvalidCoauthSession_CurrentExclusiveLock()
        {
            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert a schema lock to an exclusive lock, expect the server returns the error code "InvalidCoauthSession".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, true);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1589
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1589,
                         @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] There is a current exclusive lock on the file from another client with a different schema lock identifier.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] There is a current exclusive lock on the file from another client with a different schema lock identifier.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive lock on a file which is already locked with a schema lock with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC07_ConvertToExclusiveLock_InvalidCoauthSession_DifferentSchemaLockId()
        {
            // the user01 to prepare a schema lock on the default file
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, false);
            subRequest.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            subRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();

            // Convert a schema lock to an exclusive lock with the different schema lock ID and client ID comparing with the previous step, expect the server returns the error code "InvalidCoauthSession".
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1590
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1590,
                         @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] There is a shared lock on the file from another client with a different schema lock identifier.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] There is a shared lock on the file from another client with a different schema lock identifier.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive lock on a file which no shared lock exists on.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC08_ConvertToExclusiveLock_InvalidCoauthSession_NoSharedLockOrNotPresentInSessesion()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Convert a schema lock to an exclusive lock when the schema lock does not exist, expect the server returns the error code "InvalidCoauthSession".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, false);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1189, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1189
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             errorCode,
                             "MS-FSSHTTP",
                             1189,
                             @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any one of the following conditions is true:] There is no shared lock.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1190
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             errorCode,
                             "MS-FSSHTTP",
                             1190,
                             @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any one of the following conditions is true:] There is no coauthoring session for the file.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2255
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             2255,
                             @"[In Convert to Exclusive Lock] The shared lock is not converted to an exclusive lock if no clients is currently editing the document.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1189, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    errorCode,
                    @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure if any one of the following conditions is true:] There is no shared lock.");

                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                        @"[In Convert to Exclusive Lock] The shared lock is not converted to an exclusive lock if no clients is currently editing the document.");
                }
            }

            // User01 join the coauthoring session
            this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // User02 convert the schema lock to exclusive using different client ID
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, false);
            subRequest.SubRequestData.ClientID = System.Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1191
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidCoauthSession,
                         errorCode,
                         "MS-FSSHTTP",
                         1191,
                         @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] The current client is not present in the coauthoring session.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidCoauthSession,
                    errorCode,
                    @"[In Convert to Exclusive Lock][The protocol server returns an error code value set to ""InvalidCoauthSession"" to indicate failure, if any one of the following conditions is true:] The current client is not present in the coauthoring session.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive with the ReleaseLockOnConversionToExclusiveFailure attribute is set to false and more than one client currently editing the file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC09_ConvertToExclusiveLock_MultipleClientsInCoauthSession()
        {
            // the user01 to prepare a schema lock on the default file
            // Initialize the service
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // the user02 to prepare a shared schema lock with different client id on the default file.
            this.PrepareSchemaLock(this.DefaultFileUrl, System.Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);

            // Convert a schema lock to an exclusive lock with all valid parameters, expect the server returns the error code "MultipleClientsInCoauthSession".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, false);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1196 and MS-FSSHTTP_R38902
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.MultipleClientsInCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1196,
                         @"[In Convert to Exclusive Lock] If there is more than one client currently editing the file and the ReleaseLockOnConversionToExclusiveFailure attribute is set to false, the protocol server returns an error code value set to ""MultipleClientsInCoauthSession"" to indicate the failure to convert to an exclusive lock.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.MultipleClientsInCoauthSession,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         38902,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] MultipleClientsInCoauthSession indicates an error when all of the following conditions are true:
                         1.A schema lock subrequest of type ""Convert to exclusive lock"" is requested on a file;
                         2.There is more than one client in the current coauthoring session for that file;
                         3.The ReleaseLockOnConversionToExclusiveFailure attribute specified as part of the subrequest is set to false.");

                bool isInsession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R678 and MS-FSSHTTP_R680, expect the user {0} with clientID {1} and SchemaLockId {2} exist in the coauthoring session, actually it {3}.",
                    this.UserName01,
                    SharedTestSuiteHelper.DefaultClientID,
                    SharedTestSuiteHelper.ReservedSchemaLockID,
                    isInsession ? "exists" : "not exists");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R678
                Site.CaptureRequirementIfIsTrue(
                         isInsession,
                         "MS-FSSHTTP",
                         678,
                         @"[In SchemaLockSubRequestDataType] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker, provided that all of the following conditions are true:
The type of the schema lock subrequest is ""Convert to an Exclusive Lock"";
The conversion to an exclusive lock failed.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R680
                Site.CaptureRequirementIfIsTrue(
                         isInsession,
                         "MS-FSSHTTP",
                         680,
                         @"[In SchemaLockSubRequestDataType] When all the preceding conditions[1. The type of the schema lock subrequest is ""Convert to an Exclusive Lock"". 2. The conversion to an exclusive lock failed.] are true, the following apply: A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of false indicates that the protocol server is not allowed to remove the ClientId entry associated with the current client in the File coauthoring tracker.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R446
                Site.CaptureRequirementIfIsTrue(
                         isInsession,
                         "MS-FSSHTTP",
                         446,
                         @"[In SubRequestDataOptionalAttributes][When all the above conditions 1. The type of co-authoring sub request is ""Convert to an exclusive lock"" or the type of the schema lock sub request is ""Convert to an Exclusive Lock"" 2. The conversion to an exclusive lock failed] are true:
A ReleaseLockOnConversionToExclusiveFailure attribute set to a value of false indicates that the protocol server is not allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.MultipleClientsInCoauthSession,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Convert to Exclusive Lock] If there is more than one client currently editing the file and the ReleaseLockOnConversionToExclusiveFailure attribute is set to false, the protocol server returns an error code value set to ""MultipleClientsInCoauthSession"" to indicate the failure to convert to an exclusive lock.");

                bool isInsession = this.IsPresentInCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);
                this.Site.Log.Add(
                  LogEntryKind.Debug,
                  "For MS-FSSHTTP_R678 and MS-FSSHTTP_R680, expect the user {0} with clientID {1} and SchemaLockId {2} exist in the coauthoring session, actually it {3}.",
                  this.UserName01,
                  SharedTestSuiteHelper.DefaultClientID,
                  SharedTestSuiteHelper.ReservedSchemaLockID,
                  isInsession ? "exists" : "not exists");
                Site.Assert.IsTrue(
                    isInsession,
                    @"[In SchemaLockSubRequestDataType] ReleaseLockOnConversionToExclusiveFailure: A Boolean value that specifies to the protocol server whether the server is allowed to remove the ClientID entry associated with the current client in the File coauthoring tracker, provided that all of the following conditions are true:
The type of the schema lock subrequest is ""Convert to an Exclusive Lock"";
The conversion to an exclusive lock failed.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert a schema lock to an exclusive succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC10_ConvertToExclusiveLock_Success()
        {
            // Prepare a schema lock
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Convert a schema lock to an exclusive lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ConvertToExclusive, null, false);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1188
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1188,
                         @"[In Convert to Exclusive Lock] The protocol server performs the following operations:
                         1.The conversion of the shared lock to an exclusive lock.;
                         2.The deletion of the coauthoring session.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2254
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2254,
                         @"[In Convert to Exclusive Lock] The shared lock is converted to an exclusive lock if one client is currently editing the document.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Convert to Exclusive Lock] The protocol server performs the following operations:
                    1.The conversion of the shared lock to an exclusive lock.;
                    2.The deletion of the coauthoring session.");
            }

            this.StatusManager.CancelSharedLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);
        }

        #endregion

        #region Get schema lock

        /// <summary>
        /// A method used to verify the related requirements when the Get Lock of SchemaLock on a file which is already check out by different user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC11_GetLock_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name.
            bool isCheckOutSuccess = SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            Site.Assert.AreEqual(true, isCheckOutSuccess, "Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters but with different user account, expect the server responses the error code "FileAlreadyCheckedOutOnServer".
            // Now the web service is initialized using the user01, so the user is different with the user who check the file out.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R675
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         675,
                         @"[In SchemaLockSubRequestDataType][AllowFallbackToExclusive] When shared locking on the file is not supported: An AllowFallbackToExclusive attribute value set to false indicates that a schema lock subrequest is not allowed to fall back to an exclusive lock subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1577
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1577,
                         @"[In Get Lock][Get Lock in SchemaLock Sub Request] If the coauthorable file is checked out on the server and it is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");

                bool isVerifyR385 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    schemaLockSubResponse.ErrorMessage);

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
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SchemaLockSubRequestDataType][AllowFallbackToExclusive] When shared locking on the file is not supported: An AllowFallbackToExclusive attribute value set to false indicates that a schema lock subrequest is not allowed to fall back to an exclusive lock subrequest.");

                bool isVerifyR385 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                  LogEntryKind.Debug,
                  "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                  this.UserName02,
                  schemaLockSubResponse.ErrorMessage);
                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get a schema lock on a file which is already locked with an exclusive lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC12_GetLock_FileAlreadyLockedOnServer_CurrentExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Get a schema lock with all valid parameters, expect the server returns the error code "FileAlreadyLockedOnServer".
            SchemaLockSubRequestType subRequestType = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestType });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1575
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1575,
                         @"[In Get Lock][Get Lock in SchemaLock Sub Request] If there is a current exclusive lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");

                bool isVerifyR293 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.ToUpper(CultureInfo.CurrentCulture).Contains(UserName01.ToUpper(CultureInfo.CurrentCulture));

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R293, the error message should contain the user name {0}, but actual value is {1}",
                    this.UserName01,
                    schemaLockSubResponse.ErrorMessage);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R293
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR293,
                         "MS-FSSHTTP",
                         293,
                         @"[In SubResponseType][ErrorMessage]If the error code value is set to ""FileAlreadyLockedOnServer"", the protocol server returns the user name of the client that is currently holding the lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Get Lock][Get Lock in SchemaLock Sub Request] If there is a current exclusive lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");

                bool isVerifyR293 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.ToUpper(CultureInfo.CurrentCulture).Contains(UserName01.ToUpper(CultureInfo.CurrentCulture));
                this.Site.Log.Add(
                 LogEntryKind.Debug,
                 "For MS-FSSHTTP_R293, the error message should contain the user name {0}, but actual value is {1}",
                 this.UserName01,
                 schemaLockSubResponse.ErrorMessage);
                Site.Assert.IsTrue(
                    isVerifyR293,
                    @"[In SubResponseType][ErrorMessage]If the error code value is set to ""FileAlreadyLockedOnServer"", the protocol server returns the user name of the client that is currently holding the lock on the file.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get a schema lock on a file which is already locked with a different schema lock identifier
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC13_GetLock_FileAlreadyLockedOnServer_DifferentSchemaLockId()
        {
            // User01 prepare a schema lock.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get an schema lock with different schema lock ID comparing the previous step, expect the server responses the error code "FileAlreadyLockedOnServer".
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            SchemaLockSubRequestType subRequestDifferentSchemaLockId = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequestDifferentSchemaLockId.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestDifferentSchemaLockId });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1576
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1576,
                         @"[In Get Lock][Get Lock in SchemaLock Sub Request] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Get Lock][Get Lock in SchemaLock Sub Request] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get a schema lock with the schema lock identifier set to null.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC14_GetLock_InvalidArgument()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with the SchemaLockId set to null, expect the server returns the error code "InvalidArgument".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest.SubRequestData.SchemaLockID = null;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);


        }

        /// <summary>
        /// A method used to verify the related requirements when get a schema lock on a file which is locked by more than the max number of users.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC15_GetLock_NumberOfCoauthorsReachedMax()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Set max number of coauthors to 2.
            bool isCheckOutSuccess = SutPowerShellAdapter.SetMaxNumOfCoauthUsers(2);
            Site.Assert.AreEqual(true, isCheckOutSuccess, "SetMaxNumOfCoauthUsers success!");

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            CellStorageResponse response = new CellStorageResponse();
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Get a schema lock with a different ClientId comparing with the previous step, expect the server returns the error code "Success".
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            SchemaLockSubRequestType subRequest2 = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest2.SubRequestData.ClientID = Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest2 });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest2.SubRequestData.ClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3107, this.Site))
            {
                this.InitializeContext(this.DefaultFileUrl, this.UserName03, this.Password03, this.Domain);
                int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
                int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

                // Get a schema lock with different ClientId comparing with the previous two steps, expect the server returns the error code "NumberOfCoauthorsReachedMax".
                SchemaLockSubRequestType subRequest3 = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
                subRequest3.SubRequestData.ClientID = Guid.NewGuid().ToString();

                while (retryCount > 0)
                {
                    response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest3 });
                    schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

                    if (SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site) == ErrorCodeType.NumberOfCoauthorsReachedMax)
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
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1161
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.NumberOfCoauthorsReachedMax,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1161,
                             @"[In Get Lock] The protocol server returns an error code value set to ""NumberOfCoauthorsReachedMax"" when all of the following conditions are true: 
                         1.The maximum number of coauthorable clients allowed to join a coauthoring session to edit a coauthorable file has been reached;
                         2.The current client is not allowed to edit the file because the limit has been reached.");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.NumberOfCoauthorsReachedMax,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                        @"[In Get Lock] The protocol server returns an error code value set to ""NumberOfCoauthorsReachedMax"" when all of the following conditions are true: 
                        1.The maximum number of coauthorable clients allowed to join a coauthoring session to edit a coauthorable file has been reached;
                        2.The current client is not allowed to edit the file because the limit has been reached.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when getting a schema lock successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC16_GetLock_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);

            if (errorCode == ErrorCodeType.Success)
            {
                this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3127
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3127,
                         @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] If the protocol server supports shared locking, the protocol server MUST support at least one of the following subrequests:
                         The schema lock subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Common Message Processing Rules and Events][The protocol server MUST follow the following common processing rules for all types of subrequests] If the protocol server supports shared locking, the protocol server MUST support at least one of the following subrequests:
                       The schema lock subrequest.");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1152
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         errorCode,
                         "MS-FSSHTTP",
                         1152,
                         @"[In SchemaLock Subrequest][The protocol server returns results based on the following conditions:] An ErrorCode value of ""Success"" indicates success in processing the schema lock request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1416
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         errorCode,
                         "MS-FSSHTTP",
                         1416,
                         @"[In SubResponseElementGenericType] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the cell storage service subrequest.");

                this.CaptureLockTypeRelatedRequirementsWhenGetLockSucceed(schemaLockSubResponse);

                // Verify the return value whether has LockType attribute.
                string lockType = schemaLockSubResponse.SubResponseData.LockType;

                // Add the log information.
                Site.Log.Add(LogEntryKind.Debug, "For the requirement MS-FSSHTTP_R401 and MS-FSSHTTP_R403, expect the LockType value is SchemaLock, actually the LockType of GetLockInSchemaLockSubRequest is :{0}", lockType);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R401
                Site.CaptureRequirementIfAreEqual<string>(
                         "SchemaLock",
                         lockType,
                         "MS-FSSHTTP",
                         401,
                         @"[In LockTypes] SchemaLock: The string value ""SchemaLock"", indicating a shared lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R403
                Site.CaptureRequirementIfAreEqual<string>(
                         "SchemaLock",
                         lockType,
                         "MS-FSSHTTP",
                         403,
                         @"[In LockTypes][SchemaLock or 1] In a cell storage service response message, a shared lock indicates that the current client is granted a shared lock on the file, which allows for coauthoring the file along with other clients.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    errorCode,
                    @"[In SchemaLock Subrequest][The protocol server returns results based on the following conditions:] An ErrorCode value of ""Success"" indicates success in processing the schema lock request.");

                Site.Assert.AreEqual<string>(
                   "SchemaLock",
                    schemaLockSubResponse.SubResponseData.LockType,
                    @"[In LockTypes] SchemaLock: The string value ""SchemaLock"", indicating a shared lock on the file.");
            }

            // Get a schema lock with the same parameters with previous step, expect the server returns the error code "Success".
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1152
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1156,
                         @"[In Get Lock][If the file already has a shared lock on the server with the given schema lock identifier and the client has already joined the coauthoring session, the protocol server does both of the following:] Return an error code value set to ""Success"".");

                // If this operation succeeds, it can approve that the implementation can support the schema lock feature.
                Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             93102,
                             @"[In Appendix B: Product Behavior] Implementation does support the schema lock subrequest. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Get Lock][If the file already has a shared lock on the server with the given schema lock identifier and the client has already joined the coauthoring session, the protocol server does both of the following:] Return an error code value set to ""Success"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when getting a schema lock successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC17_GetLock_Success_AnotherClientUseSameSchemaLockId()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get a schema lock with the same SchemaLockId and different ClientId comparing the previous step, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R688
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         688,
                         @"[In SchemaLockSubRequestDataType] The protocol server ensures that at any instant in time, only clients having the same schema lock identifier can lock the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1491
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1491,
                         @"[In SchemaLockSubRequestDataType] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SchemaLockSubRequestDataType] The protocol server ensures that at any instant in time, only clients having the same schema lock identifier can lock the file.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the Get Lock of SchemaLock on a file when the library's coauthoring feature is disabled.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC18_GetLock_ExclusiveLockReturnReason_CoauthoringDisabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Set the library turn off the coauthoring feature.
            if (!this.SutPowerShellAdapter.SwitchCoauthoringFeature(true))
            {
                this.Site.Assert.Fail("Cannot disable the coauthoring feature");
            }

            // Record the disable the coauthoring feature status.
            this.StatusManager.RecordDisableCoauth();

            // Waiting change takes effect
            System.Threading.Thread.Sleep(30 * 1000);

            // Get a schema lock with AllowFallbackToExclusive set to true, expect the server responses the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, true, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "The test case cannot continue unless the server responses Success when the user get a schema lock and the coauthoring feature is disabled.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Capture lock type related requirements.
                this.CaptureLockTypeRelatedRequirementsWhenGetLockSucceed(schemaLockSubResponse);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R404
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         404,
                         @"[In LockTypes] ExclusiveLock: The string value ""ExclusiveLock"", indicating an exclusive lock on the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R406
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         406,
                         @"[In LockTypes][ExclusiveLock or 2] In a cell storage service response message, an exclusive lock indicates that an exclusive lock is granted to the current client for that specific file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1157 and MS-FSSHTTP_R442
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         1157,
                         @"[In Get Lock] If the coauthoring feature is disabled on the protocol server, the server does one of the following: 
                         If the AllowFallbackToExclusive attribute is set to true, the protocol server gets an exclusive lock on the file.");

                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         schemaLockSubResponse.SubResponseData.LockType,
                         "MS-FSSHTTP",
                         442,
                         @"[In SubRequestDataOptionalAttributes] When shared locking on the file is not supported: An AllowFallbackToExclusive attribute value set to true indicates that a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is allowed to fall back to an exclusive lock subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R441
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         441,
                         @"[In SubRequestDataOptionalAttributes] AllowFallbackToExclusive: A Boolean value that specifies to a protocol server whether a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is allowed to fall back to an exclusive lock subrequest when shared locking on the file is not supported.");

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The ExclusiveLockReturnReason should be specified when the LockTypes equals ExclusiveLock, actually it is {0}.",
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified ? "specified" : "not specified");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1428
                Site.CaptureRequirementIfIsTrue(
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         1428,
                         @"[In SubResponseDataOptionalAttributes][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations when the LockType attribute in the subresponse is set to ""ExclusiveLock"": A schema lock subrequest of type ""Get lock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R717
                Site.CaptureRequirementIfIsTrue(
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         717,
                         @"[In SchemaLockSubResponseDataType] The ExclusiveLockReturnReason attribute MUST be specified in a schema lock subresponse that is generated in response to a schema lock subrequest of type ""Get lock"" when the LockType attribute in the subresponse is set to ""ExclusiveLock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R674
                Site.CaptureRequirementIfIsTrue(
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                         "MS-FSSHTTP",
                         674,
                         @"[In SchemaLockSubRequestDataType][AllowFallbackToExclusive] When shared locking on the file is not supported: An AllowFallbackToExclusive attribute value set to true indicates that a schema lock subrequest is allowed to fall back to an exclusive lock subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R349
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CoauthoringDisabled,
                         schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         349,
                         @"[In ExclusiveLockReturnReasonTypes] CoauthoringDisabled: The string value ""CoauthoringDisabled"", indicating that an exclusive lock is granted on a file because coauthoring is disabled.");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    "ExclusiveLock",
                    schemaLockSubResponse.SubResponseData.LockType,
                    @"[In LockTypes] ExclusiveLock: The string value ""ExclusiveLock"", indicating an exclusive lock on the file.");

                Site.Assert.IsTrue(
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReasonSpecified,
                    @"[In SubResponseDataOptionalAttributes][ExclusiveLockReturnReason] The ExclusiveLockReturnReason attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations when the LockType attribute in the subresponse is set to ""ExclusiveLock"": A schema lock subrequest of type ""Get lock"".");

                Site.Assert.AreEqual<ExclusiveLockReturnReasonTypes>(
                    ExclusiveLockReturnReasonTypes.CoauthoringDisabled,
                    schemaLockSubResponse.SubResponseData.ExclusiveLockReturnReason,
                    @"[In ExclusiveLockReturnReasonTypes] CoauthoringDisabled: The string value ""CoauthoringDisabled"", indicating that an exclusive lock is granted on a file because coauthoring is disabled.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when getting a schema lock using different schemaLockID after the current schema lock is released.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC19_GetLock_Success_UseDifferentSchemaLockIdAfterRelease()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Release a schema lock with all parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType releaseSubRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            releaseSubRequest.SubRequestData.ClientID = SharedTestSuiteHelper.DefaultClientID;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { releaseSubRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Release Lock of SchemaLock sub request succeeds.");
            this.StatusManager.CancelSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Get a schema lock with all different SchemaLockId comparing with the first step, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequestNew = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequestNew.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestNew });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequestNew.SubRequestData.ClientID, subRequestNew.SubRequestData.SchemaLockID);

            // Capture lock type related requirements.
            this.CaptureLockTypeRelatedRequirementsWhenGetLockSucceed(schemaLockSubResponse);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R451
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         451,
                         @"[In SubRequestDataOptionalAttributes][SchemaLockId] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R689
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         689,
                         @"[In SchemaLockSubRequestDataType] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SubRequestDataOptionalAttributes][SchemaLockId] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
            }
        }

        /// <summary>
        /// A method used to verify the schemaLock related requirements.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC20_GetLock_SchemaLockBlock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock on a file using the default schemaLockID.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "The client should get the schema lock successfully on the file {0}", this.DefaultFileUrl);
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Get a schema lock on the file using another client and the same schemaLockID.
            subRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1488
                this.Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                              ErrorCodeType.Success,
                              SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                              "MS-FSSHTTP",
                              1488,
                              @"[In SubRequestDataOptionalAttributes] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST  allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }
            else
            {
                this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In SubRequestDataOptionalAttributes] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow other protocol clients that specify the same schema lock identifier to share the file lock.");
            }

            // Get a schema lock on the file using another client and a different schemaLockID.
            subRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();
            subRequest.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R448
                this.Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                              ErrorCodeType.Success,
                              SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                              "MS-FSSHTTP",
                              448,
                              @"[In SubRequestDataOptionalAttributes][SchemaLockId] This schema lock identifier is used by the protocol server to block other clients that have different schema lock identifiers.");

                // If the previous requirements MS-FSSHTTP_R1488 and MS-FSSHTTP_R448 can be captured successfully, then MS-FSSHTTP_R450 can be captured.
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         450,
                         @"[In SubRequestDataOptionalAttributes][SchemaLockId] The protocol server ensures that at any instant of time, only clients having the same schema lock identifier can lock the file.");
            }
            else
            {
                this.Site.Assert.AreNotEqual<ErrorCodeType>(
                   ErrorCodeType.Success,
                   SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                   @"[In SubRequestDataOptionalAttributes][SchemaLockId] The protocol server ensures that at any instant of time, only clients having the same schema lock identifier can lock the file.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server's behavior when the coauthoring feature is disabled on the protocol server and the AllowFallbackToExclusive attribute is set to false.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC21_GetLock_CoauthDisabled_AllowFallbackToExclusiveIsFalse()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Set the library turn off the coauthoring feature.
            if (!this.SutPowerShellAdapter.SwitchCoauthoringFeature(true))
            {
                this.Site.Assert.Fail("Cannot disable the coauthoring feature");
            }

            // Record the disable the coauthoring feature status.
            this.StatusManager.RecordDisableCoauth();

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            SchemaLockSubResponseType subResponse = null;

            while (retryCount > 0)
            {
                // Get lock with AllowFallbackToExclusive set to false.
                SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(false);
                CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
                subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

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
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R443
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         443,
                         @"[In SubRequestDataOptionalAttributes][When shared locking on the file is not supported:] An AllowFallbackToExclusive attribute value set to false indicates that a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is not allowed to fall back to an exclusive lock subrequest.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3118, this.Site))
                {
                    // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_3118
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3118,
                             @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". <49> Section 3.1.4.4.1:  SharePoint Server 2010 will return an error code value ""FileNotLockedOnServerAsCoauthDisabled"", if the AllowFallbackToExclusive attribute is set to false.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 311801, this.Site))
                {
                    // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_311801
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             311801,
                             @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""LockNotConvertedAsCoauthDisabled"". <49> Section 3.1.4.4.1:  SharePoint Server 2019 and SharePoint Server Subscription Edition will return an error code value ""LockNotConvertedAsCoauthDisabled"", if the AllowFallbackToExclusive attribute is set to false.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3119, this.Site))
                {
                    // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_3119
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3119,
                             @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""FileNotLockedOnServer"". (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }
            }
            else
            {
                Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In SubRequestDataOptionalAttributes][When shared locking on the file is not supported:] An AllowFallbackToExclusive attribute value set to false indicates that a coauthoring subrequest of type ""Join coauthoring session"" or a schema lock subrequest of type ""Get lock"" is not allowed to fall back to an exclusive lock subrequest.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3118, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (Microsoft Office 2010 suites/Microsoft Office 2013/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 311801, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.LockNotConvertedAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""LockNotConvertedAsCoauthDisabled"". (SharePoint Server 2019 and SharePoint Server Subscription Edition follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3119, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServer,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the coauthoring feature is disabled on the protocol server, if the AllowFallbackToExclusive attribute is set to false when sending a Get Lock subrequest, the implementation does return an error code value set to ""FileNotLockedOnServer"". <42> Section 3.1.4.4.1:  SharePoint Server 2013 will return an error code value ""FileNotLockedOnServer"" if the AllowFallbackToExclusive attribute is set to false.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the timeout is refreshed when sending a Get lock subRequest if the file already has a shared lock on the server with the given schema lock identifier and the client has already joined the coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC22_GetLock_Success_RefreshTimeout()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with the value Timeout set to 60, expect the server responds the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(false);
            subRequest.SubRequestData.Timeout = "60";
            CellStorageResponse firstResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType firstSchemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(firstResponse, 0, 0, this.Site);
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3075
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(firstSchemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3075,
                         @"[In SchemaLockSubRequestDataType][Timeout] When the Timeout is set to a value ranging from 60 to 3600, the server also returns success [but sets the Timeout to an implementation-specific default value]. ");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(firstSchemaLockSubResponse.ErrorCode, this.Site),
                    @"[In SchemaLockSubRequestDataType][Timeout] When the Timeout is set to a value ranging from 60 to 3600, the server also returns success [but sets the Timeout to an implementation-specific default value]. ");
            }

            // Refresh the timeout through sending a Get Lock subRequest, expect the server responds the error code "Success".
            subRequest.SubRequestData.Timeout = "3600";
            CellStorageResponse secondResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType secondSchemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(secondResponse, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(secondSchemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the second Get Lock of SchemaLock sub request succeeds.");

            // Sleep 70 seconds which is larger than the first step defined time out value, but it is smaller than the second step specified time out value.
            SharedTestSuiteHelper.Sleep(70);

            // Check the schema lock availability on the file with a different SchemaLockId, if the server responds the error code "FileAlreadyLockedOnServer" which indicates the timeout of the schema lock associated with the default clientId was refreshed, then capture MS-FSSHTTP_R1155.
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
            subRequest.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1155
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1155,
                         @"[In Get Lock] If the file already has a shared lock on the server with the given schema lock identifier and the client has already joined the coauthoring session, the protocol server does both of the following: 
                         Refresh the timeout value associated with the ClientId in the file coauthoring tracker.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In Get Lock] If the file already has a shared lock on the server with the given schema lock identifier and the client has already joined the coauthoring session, the protocol server does both of the following: 
                        Refresh the timeout value associated with the ClientId in the file coauthoring tracker.");
            }
        }

        #endregion

        #region Refresh lock

        /// <summary>
        /// A method used to verify the related requirements when refresh the lock on a file which is checked out by a different user.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC23_RefreshLock_FileAlreadyCheckedOutOnServer()
        {
            // Check out one file by a specified user name.
            bool isCheckOutSuccess = SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            Site.Assert.AreEqual(true, isCheckOutSuccess, "Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellStorageResponse response = new CellStorageResponse();
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, false, null);
            subRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();

            // Refresh the schema lock with all valid parameters but different user account, expect server returns the error code "FileAlreadyCheckedOutOnServer".
            // Now the web service is initialized using the user01, so the user is different with the user who check the file out.
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1588, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1588
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyCheckedOutOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1588,
                             @"[In Refresh Lock] If the coauthorable file is checked out on the server and is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }

                bool isVerifyR385 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    schemaLockSubResponse.ErrorMessage);

                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1588, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If the coauthorable file is checked out on the server and is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }

                bool isVerifyR385 = schemaLockSubResponse.ErrorMessage != null && schemaLockSubResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    schemaLockSubResponse.ErrorMessage);
                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh lock on a file which is already locked with an exclusive lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC24_RefreshLock_FileAlreadyLockedOnServer_CurrentExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Refresh the schema lock with all valid parameters, expect server returns the error code "FileAlreadyLockedOnServer".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1586
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1586,
                         @"[In Refresh Lock] If there is a current exclusive lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If there is a current exclusive lock on the file, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refresh lock on a file which is locked with a different schema lock identifier.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC25_RefreshLock_FileAlreadyLockedOnServer_DifferentSchemaLockId()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock using schema lock ID
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Refresh the schema lock with different SchemaLockId comparing with previous step, expect server returns the error code "FileAlreadyLockedOnServer".
            SchemaLockSubRequestType subRequestDifferentSchemaLockId = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, false, null);
            subRequestDifferentSchemaLockId.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestDifferentSchemaLockId });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1587
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyLockedOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1587,
                         @"[In Refresh Lock] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If there is a shared lock on the file with a different schema lock identifier, the protocol server returns an error code value set to ""FileAlreadyLockedOnServer"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when refreshing lock succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC26_RefreshLock_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock using default schema lock id.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Refresh the schema lock with the default parameter values, expect server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the server responses the success.");

            // Verify the return value whether has LockType attribute.
            bool lockTypeSpecified = schemaLockSubResponse.SubResponseData.LockTypeSpecified;
            this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "For MS-FSSHTTP_R1423 and MS-FSSHTTP_R713, the LockType attribute should be specified when the ErrorCode attribute is set to \"Success\", actually the LockType attribute is {0}",
                        lockTypeSpecified ? "specified" : "not specified");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1423
                Site.CaptureRequirementIfIsTrue(
                         lockTypeSpecified,
                         "MS-FSSHTTP",
                         1423,
                         @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A schema lock subrequest of type ""Refresh lock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R713
                Site.CaptureRequirementIfIsTrue(
                         lockTypeSpecified,
                         "MS-FSSHTTP",
                         713,
                         @"[In SchemaLockSubResponseDataType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", LockType MUST be specified in a schema lock subresponse that is generated in response to a schema lock subrequest of type ""Refresh lock"".");
            }
            else
            {
                Site.Assert.IsTrue(
                    lockTypeSpecified,
                    @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A schema lock subrequest of type ""Refresh lock"".");
            }
        }

        /// <summary>
        /// This method aims to verify the related requirements when refresh lock on a file with special timeout value.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC27_RefreshLock_GreaterTimeoutValue()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with the value Timeout set to 60, expect the server responses the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest.SubRequestData.Timeout = "60";
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Refresh the schema lock with the value of Timeout set to 3600 seconds, expect server returns the error code "Success".
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, null, null);
            subRequest.SubRequestData.Timeout = "3600";
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the refresh Lock of SchemaLock sub request succeeds.");
        }

        /// <summary>
        /// This method aims to verify the related requirements when refresh lock on a file when the previous lock expired.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC28_RefreshLock_TimeExpired()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with the value Timeout set to 60, expect the server responses the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest.SubRequestData.Timeout = "60";
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Sleep 70s to try to wait the timeout expires, due to undefined behavior for time out less than 3600 seconds,
            // this lock will possible be not expired, but it will not affect the test result in the current situation.
            SharedTestSuiteHelper.Sleep(70);

            // Refresh the schema lock with the value of Timeout set to 3600, expect server returns the error code "Success".
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, null, null);
            subRequest.SubRequestData.Timeout = "3600";
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1602
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1602,
                         @"[In Refresh Lock] If the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, the protocol server does one of the following:
                         If the coauthoring feature is enabled on the protocol server, the server considers this a schema lock subrequest of type ""Get lock"" and gets a new shared lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Refresh Lock] If the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, the protocol server does one of the following:
                        If the coauthoring feature is enabled on the protocol server, the server considers this a schema lock subrequest of type ""Get lock"" and gets a new shared lock on the file.");
            }
        }

        /// <summary>
        /// This method aims to verify the protocol server returns error code "FileNotLockedOnServerAsCoauthDisabled" if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false when the timeout expires.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC29_RefreshLock_TimeExpired_CoauthDisabled()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a lock on a file using the default clientID and schemaLockID with timeout set to 60s.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequestForGetLock(null);
            subRequest.SubRequestData.Timeout = "60";
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "The client {0} should get a schemaLock with schemaLockID {1} on the file {2}", SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.DefaultFileUrl);

            // Sleep 70s to try to wait the timeout expires, due to undefined behavior for time out less than 3600 seconds,
            // this lock will possible be not expired, but it will not affect the test result in the current situation.
            SharedTestSuiteHelper.Sleep(70);

            // Disable the coauthoring feature
            if (!this.SutPowerShellAdapter.SwitchCoauthoringFeature(true))
            {
                this.Site.Assert.Fail("Cannot disable the coauthoring feature");
            }

            // Record the disable the coauthoring feature status.
            this.StatusManager.RecordDisableCoauth();

            // Waiting change takes effect
            System.Threading.Thread.Sleep(30 * 1000);

            // Refresh the schemaLock using the default clientID and schemaLockID and with the AllowFallbackToExclusive set to false.
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.RefreshLock, false, null);
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_R3121
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3121, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3121,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<52> Section 3.1.4.4.3:  SharePoint Server 2010 will return an error code value ""FileNotLockedOnServerAsCoauthDisabled"".)");
                }

                // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_R312101
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 312101, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             312101,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (<52> Section 3.1.4.4.3: SharePoint Server 2019 and SharePoint Server Subscription Edition will return an error code value ""Success"".)");
                }

                // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_R3122
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3122, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileNotLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3122,
                             @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServer"". (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3121, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServerAsCoauthDisabled,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServerAsCoauthDisabled"". (Microsoft Office 2010 suites/Microsoft Office 2013/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 312101, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""Success"". (SharePoint Server 2019 and SharePoint Server Subscription Edition follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3122, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileNotLockedOnServer,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] When the refresh of the shared lock on the file for that specific client fails because the file is no longer locked since the timeout value expired on the lock, if the coauthoring feature is disabled and if the AllowFallbackToExclusive attribute is set to false, the implementation does return an error code value set to ""FileNotLockedOnServer"". (<45> Section 3.1.4.4.3:  SharePoint Foundation 2013 and SharePoint Server 2013 will return an error code value ""FileNotLockedOnServer"".)");
                }
            }
        }

        #endregion

        #region Release schema lock

        /// <summary>
        /// A method used to release an already locked file using different schema lock id.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC30_ReleaseLock_DifferentSchemaLockWithOneClient()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock using a default schema lock.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Initialize the service using different user account.
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            SchemaLockSubRequestType subRequestDifferentSchemaLockId = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            subRequestDifferentSchemaLockId.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();

            // Release an schema lock with different schema lock ID comparing the previous step, expect the server responds the error code "FileAlreadyLockedOnServer".
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestDifferentSchemaLockId });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 158201, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R158201
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyLockedOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             158201,
                             @"[In Appendix B: Product Behavior] Implementation does return an error code of ""FileAlreadyLockedOnServer"" if there is a shared lock with a different shared lock identifier and a coauthoring session with one client in it when sending a release lock subrequest. (<49> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return an error code ""FileAlreadyLockedOnServer"" if there is a shared lock with a different shared lock identifier and a coauthoring session with one client in it.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 158201, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyLockedOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Appendix B: Product Behavior] Implementation does return an error code of ""FileAlreadyLockedOnServer"" if there is a shared lock with a different shared lock identifier and a coauthoring session with one client in it when sending a release lock subrequest. (<44> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return an error code ""FileAlreadyLockedOnServer"" if there is a shared lock with a different shared lock identifier and a coauthoring session with one client in it.)");
                }
            }
        }

        /// <summary>
        /// A method used to test release a schema lock on a file which has an exclusive lock. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC31_ReleaseLock_ExclusiveLock()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Release an exclusive lock when the file is exclusive locked, expect the server responses the error code "Success".
            SchemaLockSubRequestType subRequestDifferentSchemaLockId = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequestDifferentSchemaLockId });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1608
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1608,
                         @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" if there is a current exclusive lock on the file when sending a release lock subrequest. (<49> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return error code ""Success"" if there is an exclusive lock on the file and valid coauthoring session with more than one clients in it.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" if there is a current exclusive lock on the file when sending a release lock subrequest. (<44> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return error code ""Success"" if there is an exclusive lock on the file and valid coauthoring session with more than one clients in it.)");
            }
        }

        /// <summary>
        /// A method used to test the protocol server return "Success" if no client is present in the coauthoring session when sending a Release lock subRequest.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC32_ReleaseLock_NoClientInTheCoautuoringSession()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Release lock on the file using the default clientID.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1917
                // If the server responds the error code "Success", then capture MS-FSSHTTP_R1917
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1917,
                         @"[In Release Lock] If the current client is not already present in the coauthoring session, the protocol server does one of the following: 
                         Return ""Success"" if no clients are present in the coauthoring session.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In Release Lock] If the current client is not already present in the coauthoring session, the protocol server does one of the following: 
                        Return ""Success"" if no clients are present in the coauthoring session.");
            }
        }

        /// <summary>
        /// A method used to test the protocol server return "Success" if there is a shared lock with a different shared lock identifier and valid coauthoring session with more than one clients in it.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC33_ReleaseLock_DifferentShardLockWithMoreThanOneClients()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare two schemaLock on a file with two clientIDs.
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            this.PrepareSchemaLock(this.DefaultFileUrl, Guid.NewGuid().ToString(), SharedTestSuiteHelper.ReservedSchemaLockID);

            // Release the schemaLock using the default ClientId and a different schemaLockID with the previous step.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            subRequest.SubRequestData.SchemaLockID = Guid.NewGuid().ToString();
            subRequest.SubRequestData.ClientID = SharedTestSuiteHelper.DefaultClientID;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify the MS-FSSHTTP requirement: MS-FSSHTTP_R3120
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3120,
                         @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" if there is a shared lock with a different shared lock identifier and valid coauthoring session with more than one clients in it when sending a release lock subrequest. (<49> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return error code ""Success"" if there is a shared lock with a different shared lock identifier and valid coauthoring session with more than one clients in it.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In Appendix B: Product Behavior] Implementation does return an error code of ""Success"" if there is a shared lock with a different shared lock identifier and valid coauthoring session with more than one clients in it when sending a release lock subrequest. (<44> Section 3.1.4.4.2:  SharePoint Server 2013 and SharePoint Server 2010 return error code ""Success"" if there is a shared lock with a different shared lock identifier and valid coauthoring session with more than one clients in it.)");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when release lock when the client is not present in the current coauthoring session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC34_ReleaseLock_NotPresentInCoauthoringSession()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock 
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Release lock with different ClientId comparing with previous step, expect server returns the error code "Success" or "InvalidCoauthSession".
            SchemaLockSubRequestType releaseSubRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            releaseSubRequest.SubRequestData.ClientID = Guid.NewGuid().ToString();
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { releaseSubRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3129, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3129
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3129,
                             @"[In Appendix B: Product Behavior] The implementation does return an error code ""Success"" if the current client is not already present in the coauthoring session and if there are other clients present in the coauthoring session. (<48> Section 3.1.4.4.2:  SharePoint Server 2010 will return an error code of ""Success"" if there are other clients present in the coauthoring session.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3128, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3128 and MS-FSSHTTP_R2073
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3128,
                             @"[In Appendix B: Product Behavior] The implementation does return an error code ""InvalidCoauthSession"" if the current client is not already present in the coauthoring session and if there are other clients present in the coauthoring session. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidCoauthSession,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             2073,
                             @"[In LockAndCoauthRelatedErrorCodeTypes][InvalidCoauthSession indicates an error when one of the following conditions is true when a coauthoring subrequest or schema lock subrequest is sent:] The current client does not exist in the coauthoring session for the file.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3129, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] The implementation does return an error code ""Success"" if the current client is not already present in the coauthoring session and if there are other clients present in the coauthoring session. (<43> Section 3.1.4.4.2:  SharePoint Server 2010 will return an error code of ""Success"" if there are other clients present in the coauthoring session.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3128, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidCoauthSession,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] The implementation does return an error code ""InvalidCoauthSession"" if the current client is not already present in the coauthoring session and if there are other clients present in the coauthoring session. (Microsoft Office 2010 suites/Microsoft Office 2013/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010 follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when release lock succeeds.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC35_ReleaseLock_Success()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Prepare a schema lock 
            this.PrepareSchemaLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            // Release lock with same ClientId and SchemaLockId with first step, expect server responses the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1171
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1171,
                         @"[In Release Lock] If the coauthoring session has already been deleted, the protocol server returns an error code value set to ""Success"".");

                bool isLockExist = this.CheckSchemaLockExist(this.DefaultFileUrl, SharedTestSuiteHelper.ReservedSchemaLockID);
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For the requirement MS-FSSHTTP_R1128, the lock should be released, actually it {0}",
                    isLockExist ? "does not release." : "releases.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1128
                Site.CaptureRequirementIfIsFalse(
                         isLockExist,
                         "MS-FSSHTTP",
                         1128,
                         @"[In SchemaLock Subrequest] The protocol server also uses the ClientId sent in the schema lock subrequest to decide when to release the shared lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site),
                    @"[In Release Lock] If the coauthoring session has already been deleted, the protocol server returns an error code value set to ""Success"".");

                bool isLockExist = this.CheckSchemaLockExist(this.DefaultFileUrl, SharedTestSuiteHelper.ReservedSchemaLockID);
                Site.Assert.IsFalse(
                    isLockExist,
                    @"[In SchemaLock Subrequest] The protocol server also uses the ClientId sent in the schema lock subrequest to decide when to release the shared lock on the file.");
            }

            this.StatusManager.CancelSharedLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);
        }
        #endregion

        #region Common
        /// <summary>
        /// A method used to verify the protocol server returns ErrorCode "InvalidArgument" or "HighLevelExceptionThrown" when the attribute SchemaLockRequestType is not provided in the schema request.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S03_TC36_ExclusiveLock_MissingSchemaLockRequestType()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create an exclusive sub request without SchemaLockRequestType.
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, null, null);
            subRequest.SubRequestData.SchemaLockRequestTypeSpecified = false;

            // Call an exclusive sub request without SchemaLockRequestType, expect the server responds with error code "InvalidArgument" or "HighLevelExceptionThrown".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R6662
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 6662, this.Site))
                {
                    ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.InvalidArgument,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             6662,
                             @"[In Appendix B: Product Behavior]  If the specified attributes[SchemaLockRequestType] are not provided, the implementation does return error code. <30> Section 2.3.1.13:  SharePoint Server 2010 returns an ""InvalidArgument"" error code as part of the SubResponseData element associated with the schema lock subresponse(Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010 follow this behavior).");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3074
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3074, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.HighLevelExceptionThrown,
                             response.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             3074,
                             @"[In Appendix B: Product Behavior]  If the specified attributes[SchemaLockRequestType] are not provided, the implementation does return error code. &lt;30&gt; Section 2.3.1.13:  SharePoint Server 2013 and SharePoint Server 2016, return ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the schema lock subresponse(SharePoint Server 2013/Microsoft Office 2016/Microsoft SharePoint Server 2016).");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 6662, this.Site))
                {
                    ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidArgument,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] If the specified attributes[ExclusiveLockRequestType attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the exclusive lock subresponse. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft Office 2016/Microsoft SharePoint Server2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3074, this.Site))
                {
                    Site.Assert.AreEqual<GenericErrorCodeTypes>(
                        GenericErrorCodeTypes.HighLevelExceptionThrown,
                        response.ResponseVersion.ErrorCode,
                        @"[In Appendix B: Product Behavior] The implementation does return a ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the exclusive lock subresponse. <24> Section 2.3.1.9: In SharePoint Server 2013 [and Microsoft SharePoint Foundation 2013], if the ExclusiveLockRequestType attribute is not provided, a ""HighLevelExceptionThrown"" error code MUST be returned as part of the SubResponseData element associated with the exclusive lock subresponse.");
                }
            }
        }
        #endregion
        #endregion 

        #region private method

        /// <summary>
        /// A method used to capture LockType related requirements when getting lock on the file.
        /// </summary>
        /// <param name="response">A return value represents the SchemaLockResponse information.</param>
        private void CaptureLockTypeRelatedRequirementsWhenGetLockSucceed(SchemaLockSubResponseType response)
        {
            Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                "Test Case cannot continue unless the server responses the error code Success.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "When the request is join coauthoring session, for the requirement MS-FSSHTTP_R1134, MS-FSSHTTP_R1422, MS-FSSHTTP_R712, the LockTypeSpecified value MUST be true, but actual value is {0}",
                        response.SubResponseData.LockTypeSpecified);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1134
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         1134,
                         @"[In SchemaLock Subrequest] If the schema lock subrequest is of type ""Get lock"" or ""Refresh lock"", the protocol server MUST return the lock type granted to the client as part of the response message to the client—if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1422
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         1422,
                         @"[In SubResponseDataOptionalAttributes] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", the LockType attribute MUST be specified in a subresponse that is generated in response to one of the following types of cell storage service subrequest operations: A schema lock subrequest of type ""Get lock"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R712
                Site.CaptureRequirementIfIsTrue(
                         response.SubResponseData.LockTypeSpecified,
                         "MS-FSSHTTP",
                         712,
                         @"[In SchemaLockSubResponseDataType] If the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"", LockType MUST be specified in a schema lock subresponse that is generated in response to a schema lock subrequest of type ""Get lock"".");

                // If the above requirements are captured, then the MS-FSSHTTP_R1159 can be captured, because it indicates the lock type is returned by the lockType attribute.
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1159,
                         @"[In Get Lock] The result of the lock type obtained by the server MUST be sent as the LockType attribute in the SchemaLockSubResponseDataType.");
            }
            else
            {
                this.Site.Log.Add(
                       LogEntryKind.Debug,
                       "When the request is join coauthoring session, for the requirement MS-FSSHTTP_R1134, MS-FSSHTTP_R1422, MS-FSSHTTP_R712, the LockTypeSpecified value MUST be true, but actual value is {0}",
                       response.SubResponseData.LockTypeSpecified);

                Site.Assert.IsTrue(
                    response.SubResponseData.LockTypeSpecified,
                    @"If the schema lock subrequest is of type ""Get lock"" or ""Refresh lock"", the protocol server MUST return the lock type granted to the client as part of the response message to the client—if the ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");
            }
        }
        #endregion
    }
}