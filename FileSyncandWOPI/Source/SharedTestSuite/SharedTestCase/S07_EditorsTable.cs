namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with EditorsTable operation.
    /// </summary>
    [TestClass]
    public abstract class S07_EditorsTable : SharedTestSuiteBase
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
        public void S07_EditorsTableInitialization()
        {
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                // Initialize the default file URL, for this scenario, the target file URL should be unique for each test case
                this.DefaultFileUrl = this.PrepareFile();
            }
        }

        #endregion

        #region Test Cases

        /// <summary>
        /// A method used to verify JoinEditingSession can execute successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC01_EditorsTable_JoinEditSession()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("The implementation does not support EditorsTable. It is determined using PTFConfig property named R9001Enabled_MS-FSSHTTP-FSSHTTPB.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Call protocol adapter operation CellStorageRequest using user01 to join the editing session.
            EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(System.Guid.NewGuid().ToString(), 3600);
            DateTime time = DateTime.UtcNow;
            CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { join });
            EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);
            string firstClientId = join.SubRequestData.ClientID;

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the ErrorCode attribute is not null, MS-FSSHTTP_R1969 could be covered.
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "For MS-FSSHTTP_R1969, the ErrorCode attribute should be returned by the protocol server, actually it is {0}.",
                    subResponseJoin.ErrorCode == null ? "not returned" : "returned");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1969
                Site.CaptureRequirementIfIsNotNull(
                         subResponseJoin.ErrorCode,
                         "MS-FSSHTTP",
                         1969,
                         @"[In EditorsTable Subrequest] The protocol server returns results based on the following conditions: 
                         Depending on the type of error, the ErrorCode is returned as an attribute of the SubResponse element.");

                // If the ErrorCode attribute returned equals "Success", then MS-FSSHTTP_R1771 and MS-FSSHTTP_R1961 could be covered.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1771,
                         @"[In EditorsTableSubResponseType][ErrorCode] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the editors table subrequest.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1961
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1961,
                         @"[In EditorsTable Subrequest][The protocol client sends an editors table SubRequest message, which is of type EditorsTableSubRequestType] The protocol server responds with an editors table SubResponse message, which is of type EditorsTableSubResponseType as specified in section 2.3.1.25.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1979
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1979,
                         @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] The protocol server returns an error code value set to ""Success"" to indicate success in processing the EditorsTable request.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3081
                Site.CaptureRequirementIfIsNotNull(
                         subResponseJoin.SubResponseData,
                         "MS-FSSHTTP",
                         3081,
                         @"[In EditorsTableSubResponseType] As part of processing the editors table subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true: The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");
            }
            else
            {
                Site.Assert.IsNotNull(
                    subResponseJoin.ErrorCode,
                    @"[In EditorsTable Subrequest] The protocol server returns results based on the following conditions: 
                        Depending on the type of error, the ErrorCode is returned as an attribute of the SubResponse element.");

                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                    @"[In EditorsTableSubResponseType][ErrorCode] The protocol server sets the value of the ErrorCode attribute to ""Success"" if the protocol server succeeds in processing the editors table subrequest.");

                Site.Assert.IsNotNull(
                    subResponseJoin.SubResponseData,
                    @"[In EditorsTableSubResponseType] As part of processing the editors table subrequest, the SubResponseData element MUST be sent as part of the SubResponse element in a cell storage service response message if the following condition is true: The ErrorCode attribute that is part of the SubResponse element is set to a value of ""Success"".");
            }

            this.StatusManager.RecordEditorTable(this.DefaultFileUrl, join.SubRequestData.ClientID);

            // Call protocol adapter operation CellStorageRequest using user02 to join the editing session.
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(System.Guid.NewGuid().ToString(), 3600);
            cellStorageResponseJoin = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { join });
            subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                        "Test case cannot continue unless the user {0} join the editors table succeeds",
                        this.UserName02);
            this.StatusManager.RecordEditorTable(this.DefaultFileUrl, join.SubRequestData.ClientID, this.UserName02, this.Password02, this.Domain);
            string secondClientId = join.SubRequestData.ClientID;

            DateTime time2 = DateTime.UtcNow;
            EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
            Editor editor1 = this.FindEditorById(editorsTable, firstClientId);
            Editor editor2 = this.FindEditorById(editorsTable, secondClientId);

            this.Site.Assert.IsNotNull(
                        editor1,
                        "For the requirement MS-FSSHTTP_R693, the client id {0} should exist in the editor table.",
                        firstClientId);

            this.Site.Assert.IsNotNull(
                        editor2,
                        "For the requirement MS-FSSHTTP_R693, the client id {0} should exist in the editor table.",
                        secondClientId);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1982
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    1982,
                    @"[In Join Editing Session] The protocol server processes this request[the EditorsTableRequestType attribute is set to ""JoinEditingSession""] to add an entry to the editors table associated with the coauthorable file by adding the client’s associated ClientId, Timeout, and AsEditor status in an entry.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R693
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    693,
                    @"[In SchemaLockSubRequestDataType] When more than one client is editing the file, the protocol server MUST maintain a separate timeout value for each client.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R9001
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    9001,
                    @"[In Appendix B: Product Behavior] Implementation does support EditorsTable operation. (Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4075
                Site.CaptureRequirementIfIsTrue(
                    editor1 != null && editor2 != null,
                    "MS-FSSHTTPB",
                    4075,
                    @"[In EditorElement] CacheID: A string that serves to uniquely identify each client that has access to an editors table on a coauthorable file.");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4076
                Site.CaptureRequirementIfIsTrue(
                    editor1 != null && editor2 != null,
                    "MS-FSSHTTPB",
                    4076,
                    @"[In EditorElement] This MUST be the same as the ClientID attribute of the EditorsTableSubRequestDataType ([MS-FSSHTTP] section 2.3.1.23), the CoauthSubRequestDataType [MS-FSSHTTP] section 2.3.1.5), the ExclusiveLockSubRequestDataType ([MS-FSSHTTP] section 2.3.1.9) or the SchemaLockSubRequestDataType ([MS-FSSHTTP] section 2.3.1.13).");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4077
                Site.CaptureRequirementIfAreEqual<string>(
                    this.UserName01.ToLower(),
                    editor1.FriendlyName.ToLower(),
                    "MS-FSSHTTPB",
                    4077,
                    @"[In EditorElement] FriendlyName: A UserNameType that specifies the user name for the client.");

                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"^[a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*");
                System.Text.RegularExpressions.Regex regex2 = new System.Text.RegularExpressions.Regex(@"([a-zA-Z]([a-zA-Z0-9\-_])*\\[a-zA-Z]([a-zA-Z0-9])*)$");
                System.Text.RegularExpressions.Match match = regex2.Match(editor1.LoginName);

                bool isVerifiedR4079 = editor1.LoginName.ToLower().Contains(this.UserName01.ToLower()) && (regex.IsMatch(editor1.LoginName)
                    || (regex2.IsMatch(editor1.LoginName) && match.Success && match.Index > 0));

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4079
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR4079,
                    "MS-FSSHTTPB",
                    4079,
                    @"[In EditorElement] LoginName: A UserLoginType that specifies the user login alias of the client.");

                if (!string.IsNullOrEmpty(editor2.EmailAddress))
                {
                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4082
                    Site.CaptureRequirementIfAreEqual<string>(
                        (this.UserName02 + "@" + this.Domain).ToLower(),
                        editor2.EmailAddress.ToLower(),
                        "MS-FSSHTTPB",
                        4082,
                        @"[In EditorElement] EmailAddress: A string that specifies the email address associated with the client.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4083
                    Site.CaptureRequirementIfAreEqual<string>(
                        (this.UserName02 + "@" + this.Domain).ToLower(),
                        editor2.EmailAddress.ToLower(),
                        "MS-FSSHTTPB",
                        4083,
                        @"[In EditorElement] The format of the email address MUST be as specified in [RFC2822] section 3.4.1.");
                }
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4084
                // If can fetch the editors table, then the following requirement can be directly captured.
                Site.CaptureRequirement(
                    "MS-FSSHTTPB",
                    4084,
                    @"[In EditorElement] HasEditorPermission: A string that specifies if the editor has permission to make edits to the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTPB_R4085
                Site.CaptureRequirementIfIsTrue(
                    editor1.Timeout > 0,
                    "MS-FSSHTTPB",
                    4085,
                    @"[In EditorElement] Timeout: A positive integer that specifies the time when the editor’s entry will expire and the editor will no longer be considered an active editor, which is expressed as a tick count.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTPB_R4086
                Site.CaptureRequirementIfIsTrue(
                    editor1.Timeout > 0,
                    "MS-FSSHTTPB",
                    4086,
                    @"[In EditorElement] A single tick represents 100 nanoseconds, or one ten-millionth of a second.");
            }
        }

        /// <summary>
        /// A method used to verify LeaveEditingSession can execute successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC02_EditorsTable_LeaveEditSession()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // User01 to join the editor table session
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID);

            // User02 to join the editor table session
            this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, System.Guid.NewGuid().ToString(), this.UserName02, this.Password02, this.Domain);

            // User01 leaves the editors table session.
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            EditorsTableSubRequestType leaveEditorsTable = SharedTestSuiteHelper.CreateEditorsTableSubRequestForLeaveSession(SharedTestSuiteHelper.DefaultClientID);
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { leaveEditorsTable });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                "The user {0} should leave the editors table on the file {1} succeed.",
                this.UserName01,
                this.DefaultFileUrl);

            // Fetch the editors table
            EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);

            Editor firstEditor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);
            this.Site.Assert.IsNull(
                firstEditor,
                "For requirement MS-FSSHTTP_R1987 and MS-FSSHTTP_R1997, When the user {0} with the client id {1} leaves the editors table, the client id should not exist in the editors tables in the server.",
                this.UserName01,
                SharedTestSuiteHelper.DefaultClientID);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the above asserts are valid, then MS-FSSHTTP_R1987, MS-FSSHTTP_R1997 can be captured.
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1987,
                         @"[In Leave Editing Session] The protocol server processes this request[the EditorsTableRequestType attribute is set to ""LeaveEditingSession""] to remove the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1997
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         1997,
                         @"[In Remove Editor Metadata] The protocol server processes this request[Remove Editor Metadata] to remove the client-supplied key/value pair for the given key in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
        }

        /// <summary>
        /// A method used to verify when the specified attributes are not provided, the protocol server returns error code "InvalidArgument".
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC03_EditorsTable_NoClientId()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a join editor session with the ClientId set to null.
            EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(null, SharedTestSuiteHelper.DefaultTimeOut);

            // Call protocol adapter operation CellStorageRequest using user01 to join the editing session.
            CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { join });
            EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the ErrorCode attribute returned equals "InvalidArgument", then MS-FSSHTTP_R1730 and MS-FSSHTTP_R2248 can be covered.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidArgument,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1730,
                         @"[In EditorsTableSubRequestDataType] If the specified attributes are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the editors table subresponse.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2248
                Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2248,
                         @"[In EditorsTableSubResponseType][ErrorCode] The protocol server must not set the value of the ErrorCode attribute to ""Success"" if the protocol server fails in processing the editors table subrequest.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidArgument,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                    @"[In EditorsTableSubRequestDataType] If the specified attributes are not provided, an ""InvalidArgument"" error code MUST be returned as part of the SubResponseData element associated with the editors table subresponse.");

                Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                    @"[In EditorsTableSubResponseType][ErrorCode] The protocol server must not set the value of the ErrorCode attribute to ""Success"" if the protocol server fails in processing the editors table subrequest.");
            }
        }

        /// <summary>
        /// A method used to verify RefreshEditSession can execute successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC04_EditorsTable_RefreshEditSession()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Join the editors table using default client id
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, this.UserName03, this.Password03, this.Domain);

            EditorsTable editorsTableBefore = this.FetchEditorTable(this.DefaultFileUrl);
            var editorBefore = editorsTableBefore.Editors.First();

            // Sleep 10 seconds to avoid concurrency issue.
            SharedTestSuiteHelper.Sleep(10);

            // Refresh the editor entry timeout value to 7200 seconds.
            EditorsTableSubRequestType refreshSubRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForRefreshSession(SharedTestSuiteHelper.DefaultClientID, 7200);
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { refreshSubRequest });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(
                     ErrorCodeType.Success,
                     SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                     "Test case cannot continue unless the refresh editors table time out operation succeeds.");

            EditorsTable editorsTableAfter = this.FetchEditorTable(this.DefaultFileUrl);
            var editorAfter = editorsTableAfter.Editors.First();

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreNotEqual<long>(
                         editorBefore.Timeout,
                         editorAfter.Timeout,
                         "MS-FSSHTTP",
                         1990,
                         @"[In Refresh Editing Session] The protocol server processes this request[the EditorsTableRequestType attribute is set to ""RefreshEditingSession""] to refresh the Timeout value in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
            else
            {
                Site.Assert.AreNotEqual<long>(
                     editorBefore.Timeout,
                     editorAfter.Timeout,
                    @"[In Refresh Editing Session] The protocol server processes this request[the EditorsTableRequestType attribute is set to ""RefreshEditingSession""] to refresh the Timeout value in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
        }

        /// <summary>
        /// A method used to verify if the ClientId does not exist in the editors table currently when updating the table, the protocol server returns an error code value set to "EditorClientIdNotFound".
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC05_EditorsTable_UpdateClientIdNotExistIn()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a update editor session with ClientId which is not in the editors table.
            byte[] content = System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4));
            EditorsTableSubRequestType update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(System.Guid.NewGuid().ToString(), "key", content);

            // Call protocol adapter operation CellStorageRequest to update the editing session.
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the ErrorCode attribute returned equals "EditorClientIdNotFound", then MS-FSSHTTP_R1974, MS-FSSHTTP_R3037 can be covered.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.EditorClientIdNotFound,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1974,
                         @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] If the ClientId does not currently exist in the editors table, the protocol server returns an error code value set to ""EditorClientIdNotFound"" for the Update editor metadata request.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.EditorClientIdNotFound,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3037,
                         @"[In NewEditorsTableCategoryErrorCodeTypes] The value ""EditorClientIdNotFound"" indicates an error when the specify client does not currently exist in the editors table.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.EditorClientIdNotFound,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] If the ClientId does not currently exist in the editors table, the protocol server returns an error code value set to ""EditorClientIdNotFound"" for the Update editor metadata request.");
            }
        }

        /// <summary>
        /// A method used to verify UpdateEditSession can execute successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC06_EditorsTable_UpdateEditSession()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join the editors table using default client id
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID);

            // Create a update editor session object.
            byte[] content = System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10));
            EditorsTableSubRequestType update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, "testKey", content);

            // Call protocol adapter operation CellStorageRequest to update the editing session.
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        "testKey",
                        "testValue",
                        this.DefaultFileUrl);

            // Get the RefreshEditSession editors table.
            EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
            Editor editor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);

            bool isUpdated = editor.Metadata.ContainsKey("testKey");
            this.Site.Assert.IsTrue(
                    isUpdated,
                    "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                    "testKey",
                    "testValue",
                    this.DefaultFileUrl);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1993
                Site.CaptureRequirementIfIsTrue(
                         isUpdated,
                         "MS-FSSHTTP",
                         1993,
                         @"[In Update Editor Metadata] The protocol server processes this request[Update Editor Metadata] to add the client-supplied key/value pair in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isUpdated,
                    @"[In Update Editor Metadata] The protocol server processes this request[Update Editor Metadata] to add the client-supplied key/value pair in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
        }

        /// <summary>
        /// A method used to verify if the key exceeds the server's length limit, the protocol server returns error code set to "EditorMetadataStringExceedsLengthLimit".
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC07_EditorsTable_UpdateEditSession_ExceedKeyLength()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join the editors table using default client id
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID);

            // Generate a long key, the key exceeds the server limit.
            string longKey = SharedTestSuiteHelper.GenerateRandomString(280);

            // Create a update editor session object with the long key.
            byte[] content = System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10));
            EditorsTableSubRequestType update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, longKey, content);

            // Call protocol adapter operation CellStorageRequest to update the editing session.
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1977, MS-FSSHTTP_R3036 and MS-FSSHTTP_R3031
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.EditorMetadataStringExceedsLengthLimit,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1977,
                         @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] The protocol server returns an error code value set to ""EditorMetadataStringExceedsLengthLimit"" for an ""Update editor metadata"" request if the key exceeds the server’s length limit.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.EditorMetadataStringExceedsLengthLimit,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3036,
                         @"[In NewEditorsTableCategoryErrorCodeTypes] The value ""EditorMetadataStringExceedsLengthLimit"" indicates an error when the key and value exceeds the server’s length limit.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.EditorMetadataStringExceedsLengthLimit,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3031,
                         @"[In Appendix B: Product Behavior] Implementation does return NewEditorsTableCategoryErrorCodeTypes when the error occurs during the processing of an EditorsTable subrequest. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.EditorMetadataStringExceedsLengthLimit,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] The protocol server returns an error code value set to ""EditorMetadataStringExceedsLengthLimit"" for an ""Update editor metadata"" request if the key exceeds the server’s length limit.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server adds the client-supplied key/value pair in the entry in the editors table.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC08_EditorsTable_UpdateEditSession_AddKeyValuePair()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join the editors table using default client id
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID);

            // Generate a key.
            string firstKey = SharedTestSuiteHelper.GenerateRandomString(4);

            // Create a update editor session object with the key.
            byte[] content = System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(10));
            EditorsTableSubRequestType update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, firstKey, content);

            // Call protocol adapter operation CellStorageRequest to add the key/value pair.
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        firstKey,
                        "firstTestValue",
                        this.DefaultFileUrl);

            // Get the RefreshEditSession editors table to verify the key/value is added successfully.
            EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
            Editor editor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);
            bool isUpdated = editor.Metadata.ContainsKey(firstKey);
            this.Site.Assert.IsTrue(
                       isUpdated,
                       "The Key/Value pair should be added successfully.",
                       firstKey,
                       "firstTestValue",
                       this.DefaultFileUrl);

            // Create another update editor session to update the key/value pair.
            string secondKey = SharedTestSuiteHelper.GenerateRandomString(4);
            update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, secondKey, content);
            cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        secondKey,
                        "secondTestValue",
                        this.DefaultFileUrl);

            // Get the RefreshEditSession editors table to verify the key/value is updated.
            editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
            editor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);
            isUpdated = editor.Metadata.ContainsKey(secondKey);

            this.Site.Assert.IsTrue(
                   isUpdated,
                   "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                   secondKey,
                   "secondTestValue",
                   this.DefaultFileUrl);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1994
                Site.CaptureRequirementIfIsTrue(
                         isUpdated,
                         "MS-FSSHTTP",
                         1994,
                         @"[In Update Editor Metadata][or] The protocol server processes this request[Update Editor Metadata] to update the client-supplied key/value pair in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isUpdated,
                    @"[In Update Editor Metadata][or] The protocol server processes this request[Update Editor Metadata] to update the client-supplied key/value pair in the entry in the editors table associated with the coauthorable file corresponding to the client with the given ClientId.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to EditorMetadataQuotaReached when the client has already exceeded its quota for key/value pairs.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC09_EditorsTable_ExceedQuota()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join the editors table using default client id
            this.PrepareJoinEditorsTable(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID);

            // Add the first Key/value pair.
            string key = SharedTestSuiteHelper.GenerateRandomString(4);
            byte[] value = System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4));
            EditorsTableSubRequestType update = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, key, value);
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        update.SubRequestData.Key,
                        update.SubRequestData.Text[0],
                        this.DefaultFileUrl);

            EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
            Site.Assert.AreEqual<int>(1, editorsTable.Editors.Length, "Should only one editor exists.");
            Editor editor = editorsTable.Editors[0];

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4088
            Site.CaptureRequirementIfIsTrue(
                editorsTable.Editors[0].Metadata.Count >= 1,
                "MS-FSSHTTPB",
                4088,
                "[In EditorElement] Metadata: An element that specifies any arbitrary key-value pairs that the protocol client has provided. ");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4089
            Site.CaptureRequirementIfIsTrue(
                editorsTable.Editors[0].Metadata.Count >= 1,
                "MS-FSSHTTPB",
                4089,
                "[In EditorElement] Each contained element represents one such key-value pair.");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4090
            Site.CaptureRequirementIfIsTrue(
                editorsTable.Editors[0].Metadata.ContainsKey(key),
                "MS-FSSHTTPB",
                4090,
                "[In EditorElement] The name of the element is the key.");

            // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4091
            Site.CaptureRequirementIfIsTrue(
                AdapterHelper.ByteArrayEquals(value, System.Text.Encoding.Unicode.GetBytes(editorsTable.Editors[0].Metadata[key])),
                "MS-FSSHTTPB",
                4091,
                "[In EditorElement] The binary content is the value.");

            // Add the second Key/value pair.
            EditorsTableSubRequestType update2 = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.GenerateRandomString(4), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4)));
            CellStorageResponse cellStorageResponse2 = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update2 });
            EditorsTableSubResponseType subResponse2 = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse2, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse2.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        update2.SubRequestData.Key,
                        update2.SubRequestData.Text[0],
                        this.DefaultFileUrl);

            // Add the third Key/value pair.
            EditorsTableSubRequestType update3 = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.GenerateRandomString(4), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4)));
            CellStorageResponse cellStorageResponse3 = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update3 });
            EditorsTableSubResponseType subResponse3 = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse3, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse3.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        update3.SubRequestData.Key,
                        update3.SubRequestData.Text[0],
                        this.DefaultFileUrl);

            // Add the fourth Key/value pair.
            EditorsTableSubRequestType update4 = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.GenerateRandomString(4), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4)));
            CellStorageResponse cellStorageResponse4 = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update4 });
            EditorsTableSubResponseType subResponse4 = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse4, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse4.ErrorCode, this.Site),
                        "Update the editor table with the key {0} and value {1} on the file {2} should succeed.",
                        update4.SubRequestData.Key,
                        update4.SubRequestData.Text[0],
                        this.DefaultFileUrl);

            // Add the fifth Key/value pair.
            EditorsTableSubRequestType update5 = SharedTestSuiteHelper.CreateEditorsTableSubRequestForUpdateSessionMetadata(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.GenerateRandomString(4), System.Text.Encoding.Unicode.GetBytes(SharedTestSuiteHelper.GenerateRandomString(4)));
            CellStorageResponse cellStorageResponse5 = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { update5 });
            EditorsTableSubResponseType subResponse5 = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponse5, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1976, this.Site))
                {
                    // If the ErrorCode attribute returned equals "EditorMetadataQuotaReached", then MS-FSSHTTP_R1976,MS-FSSHTTP_R3035 and MS-FSSHTTP_R3031 can be captured.
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.EditorMetadataQuotaReached,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse5.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1976,
                             @"[In Appendix B: Product Behavior] The implementation does return an error code value set to ""EditorMetadataQuotaReached"" for an ""Update editor metadata"" request if the client has already exceeded 4 key/value pairs. (<49> Section 3.1.4.8: Only 4 key/value pairs can be associated with an editor on servers running Office 2013.)");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.EditorMetadataQuotaReached,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse5.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3035,
                             @"[In NewEditorsTableCategoryErrorCodeTypes] The value ""EditorMetadataQuotaReached"" indicates an error when the protocol client has already exceeded its quota for number of key/value pairs.");

                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.EditorMetadataQuotaReached,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse5.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3031,
                             @"[In Appendix B: Product Behavior] Implementation does return NewEditorsTableCategoryErrorCodeTypes when the error occurs during the processing of an EditorsTable subrequest. (Microsoft Office 2013/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2013 and above follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1976, this.Site))
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.EditorMetadataQuotaReached,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse5.ErrorCode, this.Site),
                        @"[In Appendix B: Product Behavior] The implementation does return an error code value set to ""EditorMetadataQuotaReached"" for an ""Update editor metadata"" request if the client has already exceeded 4 key/value pairs. (<49> Section 3.1.4.8: Only 4 key/value pairs can be associated with an editor on servers running Office 2013.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the EditorTable is supported when the minor version value is 2.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC10_EditorsTable_Support()
        {
            // Initialize the service
            string url = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
            this.InitializeContext(url, this.UserName01, this.Password01, this.Domain);

            // Send a serverTimeSubRequest to get the current MinorVersion value.
            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(url, new SubRequestType[] { serverTimeSubRequest });
            ServerTimeSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the serverTime request succeeds.");

            if (cellStoreageResponse.ResponseVersion.MinorVersion == 2)
            {
                // Create a join editor session object.
                EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultTimeOut);

                // Call protocol adapter operation CellStorageRequest to join the editing session.
                CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(url, new SubRequestType[] { join });

                // Get the subResponse of EditorsTableSubResponseType
                EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If the ErrorCode attribute returned equals "Success", then MS-FSSHTTP_R1703 can be covered.
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1703,
                             @"[In MinorVersionNumberType][The value of MinorVersionNumberType] 2: In responses, indicates that the protocol server is capable of managing the editors table.");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                        @"[In MinorVersionNumberType][The value of MinorVersionNumberType] 2: In responses, indicates that the protocol server is capable of managing the editors table.");
                }

                this.StatusManager.RecordEditorTable(url, SharedTestSuiteHelper.DefaultClientID);
            }
            else
            {
                Site.Assume.Inconclusive(string.Format("This test case is only valuable when the MinorVersion value is 2, but the actual MinorVersion is {0}", cellStoreageResponse.ResponseVersion.MinorVersion));
            }
        }

        /// <summary>
        /// A method used to verify the EditorTable is not supported when the minor version value is 0.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC11_EditorsTable_NotSupport_MinorVersionIsZero()
        {
            // Send a serverTimeSubRequest to get the current MinorVersion value.
            // Initialize the service
            string fileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            ServerTimeSubRequestType serverTimeSubRequest = SharedTestSuiteHelper.CreateServerTimeSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { serverTimeSubRequest });
            ServerTimeSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<ServerTimeSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the serverTime request succeeds.");

            if (cellStoreageResponse.ResponseVersion.MinorVersion == 0)
            {
                // Create a join editor session object.
                EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultTimeOut);

                // Call protocol adapter operation CellStorageRequest to join the editing session.
                CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { join });

                // Get the subResponse of EditorsTableSubResponseType
                EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If the ErrorCode attribute returned doesn't equal to "Success", then MS-FSSHTTP_R1701 can be covered.
                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1701,
                             @"[In MinorVersionNumberType][The value of MinorVersionNumberType] 0: In responses, indicates that the protocol server is not capable of managing the editors table and expects the protocol client to do so through PutChanges requests.");
                }
                else
                {
                    Site.Assert.AreNotEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                        @"[In MinorVersionNumberType][The value of MinorVersionNumberType] 0: In responses, indicates that the protocol server is not capable of managing the editors table and expects the protocol client to do so through PutChanges requests.");
                }
            }
            else
            {
                Site.Assume.Inconclusive(string.Format("This test case is only valuable when the MinorVersion value is 0, but the actual MinorVersion is {0}", cellStoreageResponse.ResponseVersion.MinorVersion));
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns an error code value set to "InvalidSubRequest" if server does not support this request type.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC12_EditorsTable_NotSupport()
        {
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("This test case only run when Editors Table is not supported.");
            }

            string fileUrl = this.PrepareFile();

            // Initialize the service
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            // Send a EditorsTable subRequest with all valid parameters, expect the protocol server returns error code "InvalidSubRequest".
            EditorsTableSubRequestType subRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, 3600);
            CellStorageResponse response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { subRequest });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1978
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.InvalidSubRequest,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1978,
                         @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] The protocol server returns an error code value set to ""InvalidSubRequest"" if server does not support this request type[EditorsTable Subrequest].");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidSubRequest,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In EditorsTable Subrequest][The protocol server returns results based on the following conditions:] The protocol server returns an error code value set to ""InvalidSubRequest"" if server does not support this request type[EditorsTable Subrequest].");
            }
        }

        /// <summary>
        /// A method used to verify read-only user can't join the editing session.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC13_EditorsTable_ReadOnlyUser()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            string readOnlyUser = Common.GetConfigurationPropertyValue("ReadOnlyUser", this.Site);
            string readOnlyUserPassword = Common.GetConfigurationPropertyValue("ReadOnlyUserPwd", this.Site);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, readOnlyUser, readOnlyUserPassword, this.Domain);

            // Create a join editor session object.
            EditorsTableSubRequestType join = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.DefaultTimeOut);

            // Call protocol adapter operation CellStorageRequest to join the editing session.
            CellStorageResponse cellStorageResponseJoin = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { join });
            EditorsTableSubResponseType subResponseJoin = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(cellStorageResponseJoin, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1738, this.Site))
                {
                    // If the ErrorCode attribute returned does not equal "Success", then MS-FSSHTTP_R1738 and MS-FSSHTTP_R3050 can be covered.
                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1738,
                             @"[In EditorsTableSubRequestDataType][AsEditor] The server MUST NOT allow a user with read-only access to join the editing session as a reader.");

                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             3050,
                             @"[In SubRequestDataOptionalAttributes][AsEditor] The server MUST NOT allow a user with read-only access to join the editing session as a reader.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1738, this.Site))
                {
                    Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponseJoin.ErrorCode, this.Site),
                    @"[In EditorsTableSubRequestDataType][AsEditor] The server MUST NOT allow a user with read-only access to join the editing session as a reader.");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server automatically adds a client to the editors table when it takes a coauthoring lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC14_EditorsTable_CoauthSession()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Join a coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, null, null);
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
            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 2050, this.Site))
            {
                // Get the RefreshEditSession editors table.
                EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
                Editor editor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);

                this.Site.Assert.IsNotNull(
                        editor,
                        "For the requirement MS-FSSHTTP_R2050, the client id {0} should exist in the editor table.",
                        SharedTestSuiteHelper.DefaultClientID);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2050
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             2050,
                             @"[In Appendix B: Product Behavior] Implementation does automatically add a client to the editors table when sending a EditorsTable subrequest. (<48> Section 3.1.4.8: Servers running Office 2013 automatically add a client to the editors table when it takes a coauthoring lock—if the client protocol version is 2.2 or higher as described in section 2.2.5.10.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server automatically adds a client to the editors table when it takes a schema lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC15_EditorsTable_SchemaLock()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);

            // Get a schema lock with all valid parameters, expect the server responses the error code "Success".
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                errorCode,
                "Test case cannot continue unless the user {0} Get schema lock id {1} with client id {3} sub request succeeds on the file {2}.",
                this.UserName01,
                SharedTestSuiteHelper.ReservedSchemaLockID,
                this.DefaultFileUrl,
                SharedTestSuiteHelper.DefaultClientID);

            this.StatusManager.RecordCoauthSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 2051, this.Site))
            {
                // Get the editors table.
                EditorsTable editorsTable = this.FetchEditorTable(this.DefaultFileUrl);
                Editor editor = this.FindEditorById(editorsTable, SharedTestSuiteHelper.DefaultClientID);

                this.Site.Assert.IsNotNull(
                        editor,
                        "For the requirement MS-FSSHTTP_R2051, the client id {0} should exist in the editor table.",
                        SharedTestSuiteHelper.DefaultClientID);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2051
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             2051,
                             @"[In Appendix B: Product Behavior] Implementation does automatically add a client to the editors table when sending a EditorsTable subrequest. (<48> Section 3.1.4.8: Servers running Office 2013 automatically add a client to the editors table when it takes a schema lock—if the client protocol version is 2.2 or higher as described in section 2.2.5.10.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify the protocol server returns Success when setting the timeout ranging from 60 to 3600.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S07_TC16_EditorsTable_Timeout()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9001, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the editors table.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Call protocol adapter operation CellStorageRequest to join the editing session with timeout set to 60.
            string clientId = Guid.NewGuid().ToString();
            EditorsTableSubRequestType subRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(clientId, 60);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(response, 0, 0, this.Site);

            // Expect the protocol server returns Success
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site), "The protocol server should return Success when the timeout is set to 60.");
            this.StatusManager.RecordEditorTable(this.DefaultFileUrl, clientId);

            // Sleep 10 seconds to avoid concurrency issue.
            SharedTestSuiteHelper.Sleep(10);

            // Call protocol adapter operation CellStorageRequest to join the editing session with timeout set to 1000.
            subRequest = SharedTestSuiteHelper.CreateEditorsTableSubRequestForRefreshSession(clientId, 1000);
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3077
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3077,
                         @"[In EditorsTableSubRequestDataType][Timeout] When the Timeout is set to a value ranging from 60 to 3600, the server also returns success [but sets the Timeout to an implementation-specific default value].");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    @"[In EditorsTableSubRequestDataType][Timeout] When the Timeout is set to a value ranging from 60 to 3600, the server also returns success [but sets the Timeout to an implementation-specific default value].");
            }
        }

        #endregion 

        #region Private Helper Method

        /// <summary>
        /// A method used to find editor table by table's id.
        /// </summary>
        /// <param name="editorsTable">A parameter represents an editor table object.</param>
        /// <param name="id">A parameter represents editor table's id.</param>
        /// <returns>A return value represents a editor table object.</returns>
        private Editor FindEditorById(EditorsTable editorsTable, string id)
        {
            return editorsTable.Editors.FirstOrDefault(e => new System.Guid(e.CacheID) == new System.Guid(id));
        }
        #endregion
    }
}