namespace Microsoft.Protocols.TestSuites.MS_FSSHTTP_FSSHTTPB
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestSuites.SharedTestSuite;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to verify the co-authoring sub request operation.
    /// </summary>
    [TestClass]
    public sealed class MS_FSSHTTP_FSSHTTPB_S02_Coauth : S02_Coauth
    {
        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            S02_Coauth.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            S02_Coauth.ClassCleanup();
        }

        #endregion

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element doesn't exist.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC01_JoinCoauthoringSession_UrlNotSpecified()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a JoinCoauthoringSession subRequest with all valid parameters.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3006Verified = false;
            try
            {
                // Send a JoinCoauthoringSession subRequest to the protocol server without specifying URL attribute.
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
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element doesn't exist, the implementation does return two ErrorCode attributes in Response element. <11> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
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
                        @"[In Appendix B: Product Behavior] If the URL attribute of the corresponding Request element doesn't exist, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the URL attribute of the corresponding Request element is an empty string.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC02_JoinCoauthoringSession_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a JoinCoauthoringSession subRequest with all valid parameters.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            CellStorageResponse response = new CellStorageResponse();
            bool isR3008Verified = false;
            try
            {
                // Send a JoinCoauthoringSession subRequest to the protocol server with empty Url attribute.
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
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is empty.");

                    Site.CaptureRequirementIfIsTrue(
                             isR3008Verified,
                             "MS-FSSHTTP",
                             3008,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does return two ErrorCode attributes in Response element. <11> Section 2.2.3.5:  SharePoint Server 2010 will return 2 ErrorCode attributes in Response element.");
                }

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3009
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                             response.ResponseCollection,
                             "MS-FSSHTTP",
                             3009,
                             @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <11> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");

                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                        GenericErrorCodeTypes.InvalidUrl,
                        response.ResponseVersion.ErrorCode,
                        "MS-FSSHTTP",
                        11071,
                        @"[In GenericErrorCodeTypes] InvalidUrl indicates an error when the associated protocol server site URL is empty.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3008, this.Site))
                {
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");

                    Site.Assert.IsTrue(
                        isR3008Verified,
                        "SharePoint server 2010 and SharePoint Foundation responses two ErrorCode attributes when the URL is non exists.");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3009, this.Site))
                {
                    Site.Assert.IsNull(
                        response.ResponseCollection,
                        @"[In Appendix B: Product Behavior] If the Url attribute of the corresponding Request element is an empty string, the implementation does not return Response element. <8> Section 2.2.3.5:  SharePoint Server 2013 will not return Response element.");

                    Site.Assert.AreNotEqual<GenericErrorCodeTypes>(
                        GenericErrorCodeTypes.InvalidUrl,
                        response.ResponseVersion.ErrorCode,
                        @"[In GenericErrorCodeTypes] InvalidUrl indicates an error when the associated protocol server site URL is empty.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the response information when executing operation JoinCoauthoringSession, RefreshCoauthoringSession, CheckLockAvailability on a file which has been checked out by different user account.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC03_CoauthSubRequest_FileAlreadyCheckedOutOnServer()
        {
            // Use SUT method to check out the file.
            if (!this.SutManagedAdapter.CheckOutFile(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain))
            {
                this.Site.Assert.Fail("Cannot change the file {0} to check out status using the user name {1} and password {2}", this.DefaultFileUrl, this.UserName02, this.Password02);
            }

            this.StatusManager.RecordFileCheckOut(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session, if the coauthorable file is checked out on the server and is checked out by a client with a different user name, the server returns a FileAlreadyCheckedOutOnServer error code.
            // Now the web service is initialized using the user01, so the user is different with the user who check the file out.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1555
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileAlreadyCheckedOutOnServer,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1555,
                         @"[In Join Coauthoring Session]  If the coauthorable file is checked out on the server and checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    @"[When sending Join Coauthoring Session subrequest] If the coauthorable file is checked out on the server and checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
            }

            bool isVerifyR385 = joinResponse.ErrorMessage != null && joinResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                this.Site.Log.Add(
                LogEntryKind.Debug,
                "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                this.UserName02,
                joinResponse.ErrorMessage);

                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                this.Site.Log.Add(
                   LogEntryKind.Debug,
                   "The error message should contain the user name {0}, actual value is {1}",
                   this.UserName02,
                   joinResponse.ErrorMessage);

                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }

            // Refresh the Coauthoring Session, if the coauthorable file is checked out on the server and is checked out by a client with a different user name,
            // the server returns a FileAlreadyCheckedOutOnServer error code.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForRefreshCoauthoringSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType refreshResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            isVerifyR385 = refreshResponse.ErrorMessage != null && refreshResponse.ErrorMessage.IndexOf(this.UserName02, StringComparison.OrdinalIgnoreCase) >= 0;
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1566, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1566
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyCheckedOutOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(refreshResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1566,
                             @"[In Refresh Coauthoring Session][When sending Refresh Coauthoring Session subrequest] If the coauthorable file is checked out on the server and checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }
                this.Site.Log.Add(
                     LogEntryKind.Debug,
                     "For the requirement MS-FSSHTTP_R385, the error message should contain the user name {0}, actual value is {1}",
                     this.UserName02,
                     refreshResponse.ErrorMessage);

                Site.CaptureRequirementIfIsTrue(
                         isVerifyR385,
                         "MS-FSSHTTP",
                         385,
                         @"[In LockAndCoauthRelatedErrorCodeTypes][FileAlreadyCheckedOutOnServer] When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1566, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.FileAlreadyCheckedOutOnServer,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(refreshResponse.ErrorCode, this.Site),
                    @"[When sending Refresh Coauthoring Session subrequest]If the coauthorable file is checked out on the server and checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }

                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The error message should contain the user name {0}, actual value is {1}",
                    this.UserName02,
                    refreshResponse.ErrorMessage);

                Site.Assert.IsTrue(
                    isVerifyR385,
                    @"When the ""FileAlreadyCheckedOutOnServer"" error code is returned as the error code value in the SubResponse element, the protocol server returns the identity of the user who has currently checked out the file in the error message attribute.");
            }

            // Check Lock Availability of the Coauthoring Session, if the file is already checked out by a different user,
            // the server returns a FileAlreadyCheckedOutOnServer error code.
            subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType checkResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            // Add the log information.
            Site.Log.Add(LogEntryKind.Debug, "The errorCode for Check Lock Availability SubRequest is :{0}", checkResponse.ErrorCode);
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1082, this.Site))
            {
                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1082
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.FileAlreadyCheckedOutOnServer,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1082,
                             @"[In Check Lock Availability] If the coauthorable file is checked out on the server and it is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.FileAlreadyCheckedOutOnServer,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(checkResponse.ErrorCode, this.Site),
                        @"[When sending check lock availability subrequest]If the coauthorable file is checked out on the server and it is checked out by a client with a different user name than the current client, the protocol server returns an error code value set to ""FileAlreadyCheckedOutOnServer"".");
                }
            }
        }
        /// <summary>
        /// A method used to verify that only the clients which have the same schema lock identifier can lock the file.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC04_JoinCoauthSession_FileNotExistsOrCannotBeCreated()
        {
            // Join a coauthoring session on a nonexistent file, expect the server returns the error code "FileNotExistsOrCannotBeCreated".
            string fileUrlNotExit = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the service
            this.InitializeContext(fileUrlNotExit, this.UserName01, this.Password01, this.Domain);

            // Join a Coauthoring session with the non-exist URL.
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 3600);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(fileUrlNotExit, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the ErrorCode attribute returned equals "FileNotExistsOrCannotBeCreated", MS-FSSHTTP_R4003 can be covered.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.FileNotExistsOrCannotBeCreated,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         4003,
                         @"[In Coauth Subrequest][The protocol server returns results based on the following conditions:] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                     ErrorCodeType.FileNotExistsOrCannotBeCreated,
                     SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                     @"[In Coauth Subrequest][The protocol server returns results based on the following conditions:] If the protocol server was unable to find the URL for the file specified in the Url attribute, the protocol server reports a failure by returning an error code value set to ""FileNotExistsOrCannotBeCreated"" in the ErrorCode attribute sent back in the SubResponse element.");
            }
        }

        /// <summary>
        /// A method used to verify that the protocol server returns an ExclusiveLockReturnReasonTypes attribute value set to "CheckedOutByCurrentUser" when the file is checked out by the current user who sent the cell storage service request message.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC05_JoinCoauthoringSession_CheckedOutByCurrentUser()
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

            // Join a Coauthoring session
            CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, true, SharedTestSuiteHelper.DefaultExclusiveLockID);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site),
                        "The test case cannot continue unless the server response succeeds when the user join the coauthoring session with allowFallback set to true on an already checked out file by the same user.");
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID);

            this.Site.Assert.AreEqual(
                "ExclusiveLock",
                response.SubResponseData.LockType,
                "The test case cannot continue unless the server responses an exclusive lock type when the user join the coauthoring session with allowFallback set to true on an already checked out file by the same user.");

            this.Site.Assert.IsTrue(
                    response.SubResponseData.ExclusiveLockReturnReasonSpecified,
                    "The test case cannot continue unless the server returns ExclusiveLockReturnReason attribute when the user join the coauthoring session with allowFallback set to true on an already checked out file by the same user.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R350
                Site.CaptureRequirementIfAreEqual<ExclusiveLockReturnReasonTypes>(
                         ExclusiveLockReturnReasonTypes.CheckedOutByCurrentUser,
                         response.SubResponseData.ExclusiveLockReturnReason,
                         "MS-FSSHTTP",
                         350,
                         @"[In ExclusiveLockReturnReasonTypes] CheckedOutByCurrentUser: The string value ""CheckedOutByCurrentUser"", indicating that an exclusive lock is granted on the file because the file is checked out by the current user who sent the cell storage service request message.");
            }
            else
            {
                Site.Assert.AreEqual<ExclusiveLockReturnReasonTypes>(
                    ExclusiveLockReturnReasonTypes.CheckedOutByCurrentUser,
                    response.SubResponseData.ExclusiveLockReturnReason,
                    "The server should response when an exclusive lock is granted on the file because the file is checked out by the current user who sent the cell storage service request message.");
            }
        }

        /// <summary>
        /// A method used to verify the related requirements when the file path is not found.
        /// </summary>
        [TestCategory("MSFSSHTTP_FSSHTTPB"), TestMethod()]
        public void MSFSSHTTP_FSSHTTPB_S02_TC06_JoinCoauthoringSession_PathNotFound()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(11265, this.Site), "This test case only runs PathNotFound is returned when the file path is not found.");

            string fileUrlNotExit = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);
            string path = fileUrlNotExit.Substring(0, fileUrlNotExit.LastIndexOf('/'));

            // Initialize the service
            this.InitializeContext(path, this.UserName01, this.Password01, this.Domain);

            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null, 3600);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(path, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.PathNotFound,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         11265,
                         @"[In Appendix B: Product Behavior] Implementation does return the value ""PathNotFound"" of GenericErrorCodeTypes when the file path is not found. (Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");

                // Verify MS-FSHTTP requirement: MS-FSSHTTP_R11072
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.PathNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    "MS-FSHTTP",
                    11072,
                    @"[In GenericErrorCodeTypes] PathNotFound indicates an error when the file path is not found.<20>");

            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                     ErrorCodeType.PathNotFound,
                     SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                     @"Error ""PathNotFound"" should be returned when the file path is not found.");
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