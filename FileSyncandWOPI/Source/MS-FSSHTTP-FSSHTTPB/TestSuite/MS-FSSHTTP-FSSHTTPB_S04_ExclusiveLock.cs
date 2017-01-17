namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the exclusive lock sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S04_ExclusiveLock : S04_ExclusiveLock
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S04_ExclusiveLock.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S04_ExclusiveLock.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to test the ErrorCode value when getting an exclusive lock and the file URL is not present.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC01_ExclusiveLock_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3006Verified = false;
            try
            {
                // Send a GetLock for exclusive subRequest to the protocol server without specifying URL attribute.
                response = this.Adapter.CellStorageRequest(null, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3006Verified = message.Contains("Duplicate attribute");
                isR3006Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3006
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3006Verified,
                             "MS-FSSHTTP",
                             3006,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3007
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3007,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3006, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.Assert.IsTrue(
                        isR3006Verified,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3007, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when get an exclusive lock on a file which is already checked out by the same client.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC02_GetExclusiveLock_CheckoutByCurrentClient()
        {
            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Change the authentication mode
                if (!this.SutPowerShellAdapter.SwitchClaimsAuthentication(false))
                {
                    this.Site.Assert.Fail("Cannot change the authentication mode to windows based.");
                }

                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }
                      
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CheckLockAvailability();

            // Get the exclusive lock using the same user name comparing with the previous step for the already checked out file, expect the server responds the error code "Success".
            // Now the service channel is initialized using the userName01 account by default.
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R1247
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1247,
                         @"[In Get Lock] If the checkout of the file has been done by the current client, the protocol server MUST allow an exclusive lock on the file.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Get Lock] If the checkout of the file has been done by the current client, the protocol server MUST allow an exclusive lock on the file.");
            }

            // Record the file exclusive lock status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);
        }

        /// <summary>
        /// A method used to verify the related requirements when the Check Lock Availability of ExclusiveLock sub request succeeds.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC03_CheckExclusiveLockAvailability_Success()
        {
            // Initialize the service.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);

            // Check the exclusive lock availability with all valid parameters on a file on which there is no lock, expect the server responds the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_RR13122
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         13122,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[the file is not checked out and there is no current exclusive lock and shared lock on the file], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");

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
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[the file is not checked out and there is no current exclusive lock and shared lock on the file], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }

            // Prepare an exclusive lock
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Check the exclusive lock availability with all valid parameters on a file which is locked by the same exclusive lock id, expect the server responds the error code "Success".
            exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R13121
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         13121,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[there is a current exclusive lock on the file with a same exclusive lock identifier with the one specified by the current client], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[there is a current exclusive lock on the file with a same exclusive lock identifier with the one specified by the current client], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }

            if (!this.StatusManager.RollbackStatus())
            {
                this.Site.Assert.Fail("Cannot release the exclusive lock {0}", SharedTestSuiteHelper.DefaultExclusiveLockID);
            }

            // Check out one file by a specified user name. 
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Change the authentication mode
                if (!this.SutPowerShellAdapter.SwitchClaimsAuthentication(false))
                {
                    this.Site.Assert.Fail("Cannot change the authentication mode to windows based.");
                }

                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }

            CheckLockAvailability();

            // Check the exclusive lock availability with all valid parameters on a file which is checked out by the same user, expect the server responds the error code "Success".
            exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { exclusiveLocksubRequest });
            exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "Success", then capture MS-FSSHTTP_R13123
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         13123,
                         @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[the file has been checked out by the current user and there is no current exclusive lock and shared lock on the file], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In Check Lock Availability] The protocol server returns error codes according to the following rules: In all other cases[the file has been checked out by the current user and there is no current exclusive lock and shared lock on the file], the protocol server returns an error code value set to ""Success"" to indicate the availability of the file for locking.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert an exclusive lock to coauthoring lock on a file which is checked out by the same client.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC04_ConvertToSchemaJoinCoauth_ConvertToSchemaFailedFileCheckedOutByCurrentUser()
        {
            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName01, this.Password01);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Change the authentication mode
                if (!this.SutPowerShellAdapter.SwitchClaimsAuthentication(false))
                {
                    this.Site.Assert.Fail("Cannot change the authentication mode to windows based.");
                }

                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CheckLockAvailability();

            // Get the exclusive lock using the same user name on the already checked out file, expect the server response the error code "Success".
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert current exclusive lock to a coauthoring shared lock with the same exclusive lock id as the previous step on the already checked out file, expect the server returns the error code "ConvertToSchemaFailedFileCheckedOutByCurrentUser".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchemaJoinCoauth);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals ConvertToSchemaFailedFileCheckedOutByCurrentUser, then capture MS-FSSHTTP requirement: MS-FSSHTTP_R1229, MS-FSSHTTP_R1279
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1229,
                         @"[In ExclusiveLock Subrequest][The protocol returns results based on the following conditions: ] If the protocol server gets an exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"" for a file, and the conversion fails because the file is checked out by the current client, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1279
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1279,
                         @"[In Convert to Schema Lock with Coauthoring Transition Tracked] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the exclusive lock to a shared lock on the file because the file is checked out by the current user, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R387
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         387,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] ConvertToSchemaFailedFileCheckedOutByCurrentUser indicates an error when converting to a shared lock fails because the file is checked out by the current client.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In ExclusiveLock Subrequest][The protocol returns results based on the following conditions: ] If the protocol server gets an exclusive lock subrequest of type ""Convert to schema lock with coauthoring transition tracked"" for a file, and the conversion fails because the file is checked out by the current client, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when convert an exclusive lock to schema lock on a file which is checked out by the same client.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC05_ConvertToSchema_ConvertToSchemaFailedFileCheckedOutByCurrentUser()
        {
            // Check out one file by a specified user name.
            if (!this.SutPowerShellAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password{2}", this.DefaultFileUrl, this.UserName01, this.UserName01);
            }

            // Record the file check out status.
            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3086, this.Site))
            {
                // Change the authentication mode
                if (!this.SutPowerShellAdapter.SwitchClaimsAuthentication(false))
                {
                    this.Site.Assert.Fail("Cannot change the authentication mode to windows based.");
                }

                this.StatusManager.RecordDisableClaimsBasedAuthentication();
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CheckLockAvailability();

            // Get the exclusive lock with same user name on the checked out file.
            this.PrepareExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            // Convert the current exclusive lock to a coauthoring shared lock on the already check out file, expect the server returns the error code "ConvertToSchemaFailedFileCheckedOutByCurrentUser".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.ConvertToSchema);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals "FailedFileCheckedOutByCurrentUser", then capture MS-FSSHTTP_R1230,MS-FSSHTTP_R1949
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1230,
                         @"[In ExclusiveLock Subrequest][If the protocol server gets an exclusive lock subrequest of type ""Convert to schema lock with] ""Convert to schema lock"" for a file, and the conversion fails because the file is checked out by the current client, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1299
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1299,
                         @"[In Convert to Schema Lock] The protocol server returns error codes according to the following rules: If the protocol server is unable to convert the exclusive lock to a shared lock on the file because the file is checked out by the current user, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R387
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         387,
                         @"[In LockAndCoauthRelatedErrorCodeTypes] ConvertToSchemaFailedFileCheckedOutByCurrentUser indicates an error when converting to a shared lock fails because the file is checked out by the current client.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.ConvertToSchemaFailedFileCheckedOutByCurrentUser,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                    @"[In ExclusiveLock Subrequest][If the protocol server gets an exclusive lock subrequest of type ""Convert to schema lock with] ""Convert to schema lock"" for a file, and the conversion fails because the file is checked out by the current client, the protocol server returns an error code value set to ""ConvertToSchemaFailedFileCheckedOutByCurrentUser"".");
            }
        }

        /// <summary>
        /// A method used to test the ErrorCode value when getting an exclusive lock on a non-existent file.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC06_ExclusiveLock_FileNotExistsOrCannotBeCreated()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with nonexistent file URL and expect the server returns error code "LockRequestFail" or "Unknown".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl.Substring(0, this.DefaultFileUrl.LastIndexOf('/')), new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            ErrorCodeType errorType = SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site);
            bool isR1926Verified = errorType == ErrorCodeType.LockRequestFail || errorType == ErrorCodeType.Unknown || errorType == ErrorCodeType.FileNotExistsOrCannotBeCreated;

            this.Site.Log.Add(
                LogEntryKind.Debug,
                "For requirement MS-FSSHTTP_R1926, expect the error code LockRequestFail or Unknown or FileNotExistsOrCannotBeCreated, but actually error code is " + errorType);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1926
                Site.CaptureRequirementIfIsTrue(
                         isR1926Verified,
                         "MS-FSSHTTP",
                         1926,
                         @"[In ExclusiveLock Subrequest][The protocol returns results based on the following conditions: ] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""LockRequestFail "" or ""Unknown"" or ""FileNotExistsOrCannotBeCreated""  in the ErrorCode attribute sent back in the SubResponse element.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isR1926Verified,
                    @"[In ExclusiveLock Subrequest][The protocol returns results based on the following conditions: ] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""LockRequestFail "" or ""Unknown"" or ""FileNotExistsOrCannotBeCreated""  in the ErrorCode attribute sent back in the SubResponse element.");
            }
        }

        /// <summary>
        /// A method used to test calling cell storage web service when this service is turned off.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC07_ExclusiveLock_CellStorageWebServiceDisabled()
        {
            if (!this.SutPowerShellAdapter.SwitchCellStorageService(false))
            {
                this.Site.Assert.Fail("Cannot disable the cell storage web service.");
            }

            this.StatusManager.RecordDisableCellStorageService();

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Call the GetExclusiveLock to expect fail.
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });

            this.Site.Log.Add(
                LogEntryKind.Debug,
                "When the cell storage service is turned off, the error code should be not null but actual {0}",
                response.ResponseVersion.ErrorCodeSpecified ? "is not null" : "null");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 15181, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R15181
                    Site.CaptureRequirementIfIsTrue(
                             response.ResponseVersion.ErrorCodeSpecified,
                             "MS-FSSHTTP",
                             15181,
                             @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if any one of the following is true.
                         This protocol is not enabled on the protocol server.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R368
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.WebServiceTurnedOff,
                             response.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             368,
                             @"[In GenericErrorCodeTypes] WebServiceTurnedOff indicates an error when the web service is turned off during the processing of the cell storage service request.");
                }
            }
            else
            {
                Site.Assert.IsTrue(
                    response.ResponseVersion.ErrorCodeSpecified,
                    @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if any one of the following is true.
                    This protocol is not enabled on the protocol server.");

                Site.Assert.AreEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.WebServiceTurnedOff,
                    response.ResponseVersion.ErrorCode,
                    @"[In GenericErrorCodeTypes] WebServiceTurnedOff indicates an error when the web service is turned off during the processing of the cell storage service request.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns ErrorCode "FileUnauthorizedAccess" when the user does not have permission to issue a cell storage service request to the file.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC08_ExclusiveLock_NoPermissionIssueCellStorageService()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, Common.GetConfigurationPropertyValue("NoPermisionToUseRemoteInterfaceUser", this.Site), Common.GetConfigurationPropertyValue("NoPermisionToUseRemoteInterfaceUserPwd", this.Site), this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });

            this.Site.Log.Add(
                LogEntryKind.Debug,
                "When the user {0} does not have permission to use remote service, the error code should be not null but actual {1}",
                Common.GetConfigurationPropertyValue("NoPermisionToUseRemoteInterfaceUser", this.Site),
                response.ResponseVersion.ErrorCodeSpecified ? "is not null" : "null");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R15171
                Site.CaptureRequirementIfIsTrue(
                         response.ResponseVersion.ErrorCodeSpecified,
                         "MS-FSSHTTP",
                         15171,
                         @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if any one of the following is true.
                         The user does not have permission to issue a cell storage service request to the file identified by the Url attribute of the Request element.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R359
                Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                         GenericErrorCodeTypes.FileUnauthorizedAccess,
                         response.ResponseVersion.ErrorCode,
                         "MS-FSSHTTP",
                         359,
                         @"[In GenericErrorCodeTypes] FileUnauthorizedAccess indicates an error when the targeted URL for the file specified as part of the Request element does not have correct authorization.");
            }
            else
            {
                Site.Assert.IsTrue(
                    response.ResponseVersion.ErrorCodeSpecified,
                    @"[In ResponseVersion] This attribute[ErrorCode] MUST be present if any one of the following is true.
                    The user does not have permission to issue a cell storage service request to the file identified by the URL attribute of the Request element.");

                Site.Assert.AreEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.FileUnauthorizedAccess,
                    response.ResponseVersion.ErrorCode,
                    @"[In GenericErrorCodeTypes] FileUnauthorizedAccess indicates an error when the targeted URL for the file specified as part of the Request element does not have correct authorization.");
            }
        }

        /// <summary>
        /// A method used to test the ErrorCode value when getting an exclusive lock and the file URL is empty.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S04_TC09_ExclusiveLock_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3008Verified = false;
            try
            {
                // Send a GetLock for exclusive subRequest to the protocol server with empty URL.
                response = this.Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { subRequest });
            }
            catch (System.Xml.XmlException exception)
            {
                string message = exception.Message;
                isR3008Verified = message.Contains("Duplicate attribute");
                isR3008Verified &= message.Contains("ErrorCode");
            }

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3008
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty string.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3008Verified,
                             "MS-FSSHTTP",
                             3008,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. <8> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty string.");

                    Site.Assert.IsTrue(
                        isR3008Verified,
                        "[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. &lt;3&gt; Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
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