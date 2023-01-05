namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is the base class for all the test classes in the MS-FSSHTTP-FSSHTTPB test suite.
    /// </summary>
    [TestClass]
    public abstract class SharedTestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// A value indicate performing the merge PTF configuration file once.
        /// </summary>
        private static bool isPerformMergeOperation;

        /// <summary>
        /// Gets or sets adapter instance.
        /// </summary>
        protected IMS_FSSHTTP_FSSHTTPBAdapter Adapter { get; set; }

        /// <summary>
        /// Gets or sets  the FileStatusManager instance.
        /// </summary>
        protected StatusManager StatusManager { get; set; }

        /// <summary>
        /// Gets or sets the default file URL.
        /// </summary>
        protected string DefaultFileUrl { get; set; }

        /// <summary>
        /// Gets or sets SUT PowerShell adapter instance.
        /// </summary>
        protected IMS_FSSHTTP_FSSHTTPBSUTControlAdapter SutPowerShellAdapter { get; set; }

        /// <summary>
        /// Gets or sets SUT managed adapter instance.
        /// </summary>
        protected IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter SutManagedAdapter { get; set; }

        /// <summary>
        /// Gets or sets the userName01.
        /// </summary>
        protected string UserName01 { get; set; }
        
        /// <summary>
        /// Gets or sets the password01.
        /// </summary>
        protected string Password01 { get; set; }
        
        /// <summary>
        /// Gets or sets the userName02.
        /// </summary>
        protected string UserName02 { get; set; }
        
        /// <summary>
        /// Gets or sets the password02.
        /// </summary>
        protected string Password02 { get; set; }

        /// <summary>
        /// Gets or sets the userName03.
        /// </summary>
        protected string UserName03 { get; set; }

        /// <summary>
        /// Gets or sets the password03.
        /// </summary>
        protected string Password03 { get; set; }

        /// <summary>
        /// Gets or sets the domain.
        /// </summary>
        protected string Domain { get; set; }

        #endregion 

        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test Case Initialization

        /// <summary>
        /// A test case's level initialization method for TestSuiteBase class. It will perform before each test case.
        /// </summary>
        [TestInitialize]
        public void MSFSSHTTP_FSSHTTPBTestInitialize()
        {
            if (!isPerformMergeOperation)
            {
                this.MergeConfigurationFile(this.Site);
                isPerformMergeOperation = true;
            }

            // If the shared test code are executed in MS-WOPI mode, try to verify whether run in support products by check the MS-WOPI_Supported property of MS-WOPI.
            if ("MS-WOPI".Equals(this.Site.DefaultProtocolDocShortName, System.StringComparison.OrdinalIgnoreCase))
            {
                if (!Common.GetConfigurationPropertyValue<bool>("MS-WOPI_Supported", this.Site))
                {
                    SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", this.Site);
                    this.Site.Assume.Inconclusive(@"The server does not support this specification [MS-WOPI]. It is determined by ""MS-WOPI_Supported"" SHOULDMAY property of the [{0}_{1}_SHOULDMAY.deployment.ptfconfig] configuration file.", this.Site.DefaultProtocolDocShortName, currentSutVersion);
                }
            }

            this.UserName01 = Common.GetConfigurationPropertyValue("UserName1", this.Site);
            this.Password01 = Common.GetConfigurationPropertyValue("Password1", this.Site);
            this.UserName02 = Common.GetConfigurationPropertyValue("UserName2", this.Site);
            this.Password02 = Common.GetConfigurationPropertyValue("Password2", this.Site);
            this.UserName03 = Common.GetConfigurationPropertyValue("UserName3", this.Site);
            this.Password03 = Common.GetConfigurationPropertyValue("Password3", this.Site);
            this.Domain = Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Initialize the web service using the configured UserName/Domain and password.
            this.Adapter = Site.GetAdapter<IMS_FSSHTTP_FSSHTTPBAdapter>();

            // Initialize the SUT Control Adapter instance
            this.SutPowerShellAdapter = Site.GetAdapter<IMS_FSSHTTP_FSSHTTPBSUTControlAdapter>();
            this.SutManagedAdapter = Site.GetAdapter<IMS_FSSHTTP_FSSHTTPBManagedCodeSUTControlAdapter>();

            // Initialize the file status manager to handle the environment clean up.
            this.StatusManager = new StatusManager(this.Site, this.InitializeContext);
        }

        /// <summary>
        /// A test case's level clean up. It will perform after each test case.
        /// </summary>
        [TestCleanup]
        public void MSFSSHTTP_FSSHTTPBTestCleanup()
        {
            System.Collections.Generic.Dictionary<StatusManager.ServerStatus, Action> status =
                new System.Collections.Generic.Dictionary<StatusManager.ServerStatus, Action>();

            foreach (System.Collections.Generic.KeyValuePair<StatusManager.ServerStatus, Action> pairs in this.StatusManager.DocumentLibraryStatusRollbackFunctions)
            {
                status.Add(pairs.Key, pairs.Value);
            }

            if (!this.StatusManager.RollbackStatus())
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The error message report for the clean up environments {0}",
                    this.StatusManager.GenerateErrorMessageReport());
            }

            foreach (System.Collections.Generic.KeyValuePair<StatusManager.ServerStatus, Action> pairs in status)
            {
                if (pairs.Key == StatusManager.ServerStatus.DisableCoauth)
                {
                    string url = this.PrepareFile();

                    int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
                    int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);
                    int temp = retryCount;

                    while (retryCount > 0)
                    {
                        // Join a Coauthoring session with AllowFallbackToExclusive attribute set to true
                        CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, true, SharedTestSuiteHelper.DefaultExclusiveLockID);
                        CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(url, new SubRequestType[] { subRequest });
                        CoauthSubResponseType firstJoinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
                        if (SharedTestSuiteHelper.ConvertToErrorCodeType(firstJoinResponse.ErrorCode, this.Site) != ErrorCodeType.Success)
                        {
                            break;
                        }
                        if (temp == retryCount)
                        {
                            this.StatusManager.RecordExclusiveLock(url, SharedTestSuiteHelper.DefaultExclusiveLockID);
                        }

                        if (firstJoinResponse.SubResponseData.ExclusiveLockReturnReasonSpecified
                            && firstJoinResponse.SubResponseData.ExclusiveLockReturnReason == ExclusiveLockReturnReasonTypes.CoauthoringDisabled)
                        {
                            System.Threading.Thread.Sleep(waitTime);
                            retryCount--;

                            // Release lock
                            ExclusiveLockSubRequestType coauthLockRequest = new ExclusiveLockSubRequestType();
                            coauthLockRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
                            coauthLockRequest.SubRequestData = new ExclusiveLockSubRequestDataType();
                            coauthLockRequest.SubRequestData.ExclusiveLockRequestType = ExclusiveLockRequestTypes.ReleaseLock;
                            coauthLockRequest.SubRequestData.ExclusiveLockRequestTypeSpecified = true;
                            coauthLockRequest.SubRequestData.ExclusiveLockID = SharedTestSuiteHelper.DefaultExclusiveLockID;

                            CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(
                                this.Adapter.CellStorageRequest(url, new SubRequestType[] { coauthLockRequest }), 0, 0, this.Site);

                            Site.Assert.AreEqual<string>(
                                "success",
                                subResponse.ErrorCode.ToLower(),
                                "Failed to release the exclusive lock.");
                        }
                        else
                        {
                            // Release lock
                            SchemaLockSubRequestType schemaLockRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
                            CellStorageResponse response = Adapter.CellStorageRequest(url, new SubRequestType[] { schemaLockRequest });
                            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
                            Site.Assert.AreEqual<string>(
                                "success",
                                schemaLockSubResponse.ErrorCode.ToLower(),
                                "Failed to release the schema lock.");
                            break;
                        }
                    }
                }
            }

            if (!this.StatusManager.CleanUpFiles())
            {
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "The error message report for the clean up environments {0}",
                    this.StatusManager.GenerateErrorMessageReport());
            }

            this.Adapter.Reset();
            SharedContext.Current.Clear();
        }

        #endregion 

        #region Helper Methods
        /// <summary>
        /// This method is used to prepare an exclusive lock on the file with the specified exclusive lock id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="exclusiveLockId">Specify the exclusive lock id.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareExclusiveLock(string fileUrl, string exclusiveLockId, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.PrepareExclusiveLock(fileUrl, exclusiveLockId, this.UserName01, this.Password01, this.Domain, timeout);
        }

        /// <summary>
        /// This method is used to prepare an exclusive lock on the file with the specified exclusive lock id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="exclusiveLockId">Specify the exclusive lock id.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareExclusiveLock(string fileUrl, string exclusiveLockId, string userName, string password, string domain, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.InitializeContext(fileUrl, userName, password, domain);
            
            // Get an exclusive lock
            ExclusiveLockSubRequestType getExclusiveLockSubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            getExclusiveLockSubRequest.SubRequestData.ExclusiveLockID = exclusiveLockId;
            getExclusiveLockSubRequest.SubRequestData.Timeout = timeout.ToString();
            CellStorageResponse response = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { getExclusiveLockSubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                "Get an exclusive lock with identifier {0} by the user {1}@{2} on the file {3} should succeed.",
                exclusiveLockId,
                userName,
                domain,
                fileUrl);

            this.StatusManager.RecordExclusiveLock(fileUrl, exclusiveLockId, userName, password, domain);
        }

        /// <summary>
        /// This method is used to prepare a coauthoring session on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        /// <returns>Return the current coauthoring status when the user join the coauthoring session.</returns>
        protected CoauthStatusType PrepareCoauthoringSession(string fileUrl, string clientId, string schemaLockId, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            return this.PrepareCoauthoringSession(fileUrl, clientId, schemaLockId, this.UserName01, this.Password01, this.Domain, timeout);
        }

        /// <summary>
        /// This method is used to prepare a coauthoring session on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        /// <returns>Return the current coauthoring status when the user join the coauthoring session.</returns>
        protected CoauthStatusType PrepareCoauthoringSession(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.InitializeContext(fileUrl, userName, password, domain);

            // Join a Coauthoring session
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForJoinCoauthSession(clientId, schemaLockId, null, null, timeout);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType joinResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(
                    ErrorCodeType.Success, 
                    SharedTestSuiteHelper.ConvertToErrorCodeType(joinResponse.ErrorCode, this.Site),
                    "Join coauthoring session with client ID {0} and schema lock ID {4} by the user {1}@{2} on the file {3} should succeed.",
                    clientId,
                    userName,
                    domain,
                    fileUrl,
                    schemaLockId);

            this.StatusManager.RecordCoauthSession(fileUrl, clientId, schemaLockId, userName, password, domain);

            this.Site.Assert.IsTrue(
                    joinResponse.SubResponseData.CoauthStatusSpecified,
                    "When join the coauthoring session succeeds, the coauth status should be returned.");

            return joinResponse.SubResponseData.CoauthStatus;
        }

        /// <summary>
        /// This method is used to prepare a schema lock on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareSchemaLock(string fileUrl, string clientId, string schemaLockId, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.PrepareSchemaLock(fileUrl, clientId, schemaLockId, this.UserName01, this.Password01, this.Domain, timeout);
        }

        /// <summary>
        /// This method is used to prepare a schema lock on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareSchemaLock(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.InitializeContext(fileUrl, userName, password, domain);

            // Get a schema
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            subRequest.SubRequestData.SchemaLockID = schemaLockId;
            subRequest.SubRequestData.ClientID = clientId;
            subRequest.SubRequestData.Timeout = timeout.ToString();
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success, 
                SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), 
                "Test case cannot continue unless the user {0} Get schema lock id {1} with client id {3} sub request succeeds on the file {2}.",
                this.UserName01,
                schemaLockId,
                this.DefaultFileUrl,
                clientId);
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID, userName, password, domain);
        }

        /// <summary>
        /// This method is used to prepare a schema lock or coauthoring on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareSharedLock(string fileUrl, string clientId, string schemaLockId, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.PrepareSharedLock(fileUrl, clientId, schemaLockId, this.UserName01, this.Password01, this.Domain, timeout);
        }

        /// <summary>
        /// This method is used to prepare a schema lock or coauthoring on the file with the specified schema lock id and client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareSharedLock(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93101, this.Site))
            {
                this.PrepareCoauthoringSession(fileUrl, clientId, schemaLockId, userName, password, domain, timeout);
            }
            else if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93102, this.Site))
            {
                this.PrepareSchemaLock(fileUrl, clientId, schemaLockId, userName, password, domain, timeout);
            }
            else
            {
                this.Site.Assert.Fail("The server should at least support one of operations: Coauthoring session and Schema Lock");
            }
        }

        /// <summary>
        /// This method is used to prepare to join a editors table on the file using the specified client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file.</param>
        /// <param name="clientId">Specify the client ID.</param>
        protected void PrepareJoinEditorsTable(string fileUrl, string clientId)
        {
            this.PrepareJoinEditorsTable(fileUrl, clientId, this.UserName01, this.Password01, this.Domain);
        }

        /// <summary>
        /// This method is used to prepare to join a editors table on the file using the specified client id.
        /// </summary>
        /// <param name="fileUrl">Specify the file.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <param name="timeout">Specify the timeout value.</param>
        protected void PrepareJoinEditorsTable(string fileUrl, string clientId, string userName, string password, string domain, int timeout = SharedTestSuiteHelper.DefaultTimeOut)
        {
            this.InitializeContext(fileUrl, userName, password, domain);

            // Create join editor session object.
            EditorsTableSubRequestType joinEditorTable = SharedTestSuiteHelper.CreateEditorsTableSubRequestForJoinSession(clientId, timeout);

            // Call protocol adapter operation CellStorageRequest with  EditorsTableRequestType JoinEditingSession.
            CellStorageResponse response = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { joinEditorTable });
            EditorsTableSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<EditorsTableSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "The user {0} uses the client id {1} to join the editor table should succeed.",
                        userName,
                        clientId);

            this.StatusManager.RecordEditorTable(this.DefaultFileUrl, clientId);
        }

        /// <summary>
        /// This method is used to test whether the lock with the specified id exist on the file.
        /// </summary>
        /// <param name="file">Specify the file.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <returns>Return true if exist, otherwise return false.</returns>
        protected bool CheckSchemaLockExist(string file, string schemaLockId)
        {
            return this.CheckSchemaLockExist(file, schemaLockId, this.UserName01, this.Password01, this.Domain);
        }

        /// <summary>
        /// This method is used to test whether the lock with the specified id exist on the file.
        /// </summary>
        /// <param name="file">Specify the file.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <returns>Return true if exist, otherwise return false.</returns>
        protected bool CheckSchemaLockExist(string file, string schemaLockId, string userName, string password, string domain)
        {
            this.InitializeContext(file, userName, password, domain);

            // Generate a new schema lock id value which is different with the given one.
            System.Guid newId;
            do
            {
                newId = System.Guid.NewGuid();
            }
            while (newId == new System.Guid(schemaLockId));

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93101, this.Site))
            {
                // Check the schema lock availability using the new schema lock id.
                SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
                subRequest.SubRequestData.SchemaLockID = newId.ToString();
                CellStorageResponse response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { subRequest });
                SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

                ErrorCodeType error = SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site);
                if (error == ErrorCodeType.FileAlreadyLockedOnServer)
                {
                    // Now there could be kind of two conditions:
                    //  1) There is an exclusive lock
                    //  2) There is a schema lock
                    // So it is needed to check the schema lock with the given schema lock id should exist.
                    subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.CheckLockAvailability, null, null);
                    subRequest.SubRequestData.SchemaLockID = schemaLockId;
                    response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { subRequest });
                    schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

                    return SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site) == ErrorCodeType.Success;
                }
                else
                {
                    if (error != ErrorCodeType.Success)
                    {
                        this.Site.Assert.Fail(
                            "If the schema lock {0} does not exist, check the schema lock using the id {0} should success, but actual the result is {2}",
                            schemaLockId,
                            newId.ToString(),
                            error.ToString());
                    }

                    return false;
                }
            }
            else if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93102, this.Site))
            {
                // Check the schema lock availability using the new schema lock id.
                CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, newId.ToString());
                CellStorageResponse response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { subRequest });
                CoauthSubResponseType coauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);

                ErrorCodeType error = SharedTestSuiteHelper.ConvertToErrorCodeType(coauthSubResponse.ErrorCode, this.Site);
                if (error == ErrorCodeType.FileAlreadyLockedOnServer)
                {
                    // Now there could be kind of two conditions:
                    //  1) There is an exclusive lock
                    //  2) There is a schema lock
                    // So it is needed to check the schema lock with the given schema lock id should exist.
                    subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, schemaLockId);
                    response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { subRequest });
                    coauthSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(response, 0, 0, this.Site);

                    return SharedTestSuiteHelper.ConvertToErrorCodeType(coauthSubResponse.ErrorCode, this.Site) == ErrorCodeType.Success;
                }
                else
                {
                    if (error != ErrorCodeType.Success)
                    {
                        this.Site.Assert.Fail(
                            "If the schema lock {0} does not exist, check the schema lock using the id {0} should success, but actual the result is {2}",
                            schemaLockId,
                            newId.ToString(),
                            error.ToString());
                    }

                    return false;
                }
            }

            this.Site.Assert.Fail("The server should at least support one of operations: Coauthoring session and Schema Lock");
            return false;
        }

        /// <summary>
        /// This method is used to test whether the exclusive lock with the specified id exist on the file.
        /// </summary>
        /// <param name="file">Specify the file.</param>
        /// <param name="exclusiveLockId">Specify the exclusive lock.</param>
        /// <returns>Return true if exist, otherwise return false.</returns>
        protected bool CheckExclusiveLockExist(string file, string exclusiveLockId)
        {
            return this.CheckExclusiveLockExist(file, exclusiveLockId, this.UserName01, this.Password01, this.Domain);
        }

        /// <summary>
        /// This method is used to test whether the exclusive lock with the specified id exist on the file.
        /// </summary>
        /// <param name="file">Specify the file.</param>
        /// <param name="exclusiveLockId">Specify the exclusive lock.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <returns>Return true if exist, otherwise return false.</returns>
        protected bool CheckExclusiveLockExist(string file, string exclusiveLockId, string userName, string password, string domain)
        {
            this.InitializeContext(file, userName, password, domain);

            // Generate a new schema lock id value which is different with the given one.
            System.Guid newId;
            do
            {
                newId = System.Guid.NewGuid();
            }
            while (newId == new System.Guid(exclusiveLockId));

            ExclusiveLockSubRequestType exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
            exclusiveLocksubRequest.SubRequestData.ExclusiveLockID = newId.ToString();
            CellStorageResponse response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { exclusiveLocksubRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

            ErrorCodeType error = SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site);
            if (error == ErrorCodeType.FileAlreadyLockedOnServer)
            {
                // Now there could be kind of two conditions:
                //  1) There is an exclusive lock
                //  2) There is a schema lock
                // So it is needed to check the exclusive lock with the given exclusive lock id should exist.
                exclusiveLocksubRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.CheckLockAvailability);
                exclusiveLocksubRequest.SubRequestData.ExclusiveLockID = exclusiveLockId;
                response = this.Adapter.CellStorageRequest(file, new SubRequestType[] { exclusiveLocksubRequest });
                exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);

                return SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site) == ErrorCodeType.Success;
            }
            else
            {
                if (error != ErrorCodeType.Success)
                {
                    this.Site.Assert.Fail(
                        "If the exclusive lock {0} does not exist, check the schema lock using the id {0} should success, but actual the result is {2}",
                        exclusiveLockId,
                        newId.ToString(),
                        error.ToString());
                }

                return false;
            }
        }

        /// <summary>
        /// This method is used to check the user with specified client whether exist in a coauthoring session.
        /// </summary>
        /// <param name="fileUrl">Specify the file which is locked.</param>
        /// <param name="clientId">Specify the client ID.</param>
        /// <param name="schemaLockId">Specify the schemaLock ID.</param>
        /// <param name="userName">Specify the user name of the user who calls cell storage service.</param>
        /// <param name="password">Specify the password of the user who calls cell storage service.</param>
        /// <param name="domain">Specify the domain of the user who calls cell storage service.</param>
        /// <returns>Return true is exist in the coauthoring session, otherwise return false.</returns>
        protected bool IsPresentInCoauthSession(string fileUrl, string clientId, string schemaLockId, string userName, string password, string domain)
        {
            this.InitializeContext(fileUrl, userName, password, domain);
            CoauthSubRequestType subRequest = SharedTestSuiteHelper.CreateCoauthSubRequestForMarkTransitionComplete(clientId, schemaLockId);
            CellStorageResponse cellResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { subRequest });
            CoauthSubResponseType response = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(cellResponse, 0, 0, this.Site);

            return SharedTestSuiteHelper.ConvertToErrorCodeType(response.ErrorCode, this.Site) == ErrorCodeType.Success;
        }

        /// <summary>
        /// This method is used to fetch editor table on the specified file.
        /// </summary>
        /// <param name="file">Specify the file URL.</param>
        /// <returns>Return the editors table on the file if the server support the editor tables, otherwise return null.</returns>
        protected EditorsTable FetchEditorTable(string file)
        {
            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 2053, this.Site))
            {
                // Call QueryEditorsTable to get editors table
                this.InitializeContext(file, this.UserName01, this.Password01, this.Domain);
                CellSubRequestType cellSubRequestEditorsTable = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryEditorsTable(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
                CellStorageResponse cellStorageResponseEditorsTable = this.Adapter.CellStorageRequest(file, new SubRequestType[] { cellSubRequestEditorsTable });
                CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponseEditorsTable, 0, 0, this.Site);
                FsshttpbResponse queryEditorsTableResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

                // Get EditorsTable
                EditorsTable editorsTable = EditorsTableUtils.GetEditorsTableFromResponse(queryEditorsTableResponse, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    this.Site.CaptureRequirement(
                            "MS-FSSHTTP",
                            2053,
                            @"[In Appendix B: Product Behavior]  Implementation does represent the editors table as a compressed XML fragment. (<50> Section 3.1.4.8: On servers running Office 2013, the editors table is represented as a compressed XML fragment.)");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4061
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4061,
                             @"[In Query Changes to Editors Table Partition] If the Partition Id GUID of the Target Partition Id of the sub-request (section 2.2.2.1.1) of the associated QueryChangesRequest is the value of the Guid ""7808F4DD - 2385 - 49d6 - B7CE - 37ACA5E43602"", the protocol server will interpret the QueryChangesRequest as a download of the editors table specified in [MS-FSSHTTP] section 3.1.4.8.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4063
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4063,
                             @"[In Query Changes to Editors Table Partition] The protocol server MUST return the editors table in the associated data element collection using the following process:
                           Construct a byte array representation of the editors table using UTF-8 encoding.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4064
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4064,
                             @"[In Query Changes to Editors Table Partition] [The protocol server MUST return the editors table in the associated data element collection using the following process:] Compress the byte array using the DEFLATE compression algorithm.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4065
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4065,
                             @"[In Query Changes to Editors Table Partition] [The protocol server MUST return the editors table in the associated data element collection using the following process:] Prepend the compressed byte array with the Editors Table Zip Stream Header (section 2.2.3.1.2.1.1).");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4066
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4066,
                             @"[In Query Changes to Editors Table Partition] [The protocol server MUST return the editors table in the associated data element collection using the following process:] Treat the byte array as a file stream and chunk it using the schema described in [MS-FSSHTTPD]. ");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4067
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4067,
                             @"[In Query Changes to Editors Table Partition] [The protocol server MUST return the editors table in the associated data element collection using the following process:] Sync the resulting cell using this protocol as you would any other cell.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4068
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4068,
                             @"[In Editors Table Zip Stream Header] A header that is prepended to the compressed EditorsTable byte array.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4069
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4069,
                             @"[In Editors Table Zip Stream Header] This header is an array of eight bytes.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4070
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4070,
                             @"[In Editors Table Zip Stream Header] The value and order of these bytes MUST be as follows: 0x1A, 0x5A, 0x3A, 0x30, 0x00, 0x00, 0x00, 0x00.");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4072
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4072,
                             @" [In EditorsTable] [EditorsTable schema is:]
                             <xs:element name=""EditorsTable"">
                                < xs:complexType >
                                  < xs:sequence >
                                    < xs:element ref= ""tns:EditorElement"" name = ""Editor"" minOccurs = ""0"" maxOccurs = ""unbounded"" />
                                  </ xs:sequence >
                                </ xs:complexType >
                              </ xs:element > ");

                    // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4074
                    // If can fetch the editors table, then the following requirement can be directly captured.
                    Site.CaptureRequirement(
                             "MS-FSSHTTPB",
                             4074,
                             @"[In EditorElement] [EditorElement schema is:] 
                             <xs:element name=""EditorElement"">
                                < xs:complexType >
                                  < xs:sequence >
                                    < xs:element name = ""CacheID"" minOccurs = ""1"" maxOccurs = ""1"" type = ""xs:string"" />
                                    < xs:element name = ""FriendlyName"" minOccurs = ""0"" maxOccurs = ""1"" type=""tns: UserNameType"" />
                                    < xs:element name = ""LoginName"" minOccurs = ""0"" maxOccurs = ""1"" type = ""tns:UserLoginType"" />
                                    < xs:element name = ""SIPAddress"" minOccurs = ""0"" maxOccurs = ""1"" type = ""xs:string"" />
                                    < xs:element name = ""EmailAddress"" minOccurs = ""0"" maxOccurs = ""1"" type = ""xs:string"" />
                                    < xs:element name = ""HasEditorPermission"" minOccurs = ""0"" maxOccurs = ""1"" type = ""xs:boolean"" />
                                    < xs:element name = ""Timeout"" minOccurs = ""1"" maxOccurs = ""1"" type = ""xs:positiveInteger"" />
                                    < xs:element name = ""Metadata"" minOccurs = ""0"" maxOccurs = ""1"" >
                                      < xs:complexType >
                                        < xs:sequence >
                                          < xs:any minOccurs = ""0"" maxOccurs = ""unbounded"" type = ""xs:binary"" />
                                        </ xs:sequence >
                                      </ xs:complexType >
                                    </ xs:element >
                                  </ xs:sequence >
                                </ xs:complexType >
                              </ xs:element > ");
                }

                return editorsTable;
            }

            return null;
        }

        /// <summary>
        /// This method is used to upload a new Txt file to a random generated URL which is based on the current test case name.
        /// </summary>
        /// <returns>Return the full URL of the Txt file target location.</returns>
        protected string PrepareFile()
        {
            string fileUri = SharedTestSuiteHelper.GenerateNonExistFileUrl(Site);
            string fileUrl = fileUri.Substring(0, fileUri.LastIndexOf("/", StringComparison.OrdinalIgnoreCase));
            string fileName = fileUri.Substring(fileUri.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);

            bool ret = this.SutPowerShellAdapter.UploadTextFile(fileUrl, fileName);

            if (ret == false)
            {
                this.Site.Assert.Fail("Cannot upload a file in the URL {0}", fileUri);
            }

            this.StatusManager.RecordFileUpload(fileUri);

            return fileUri;
        }

        /// <summary>
        /// Check if a file is available to take a shared lock or exclusive lock.
        /// </summary>
        protected void CheckLockAvailability()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 93101, this.Site))
            {
                return;
            }

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            while (retryCount > 0)
            {
                retryCount--;

                CoauthSubRequestType request = SharedTestSuiteHelper.CreateCoauthSubRequestForCheckLockAvailability(SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
                CellStorageResponse storageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { request });
                CoauthSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CoauthSubResponseType>(storageResponse, 0, 0, this.Site);

                if (SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site) == ErrorCodeType.Success)
                {
                    break;
                }
                else
                {
                    System.Threading.Thread.Sleep(waitTime);
                }

                if (retryCount == 0)
                {
                    this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                        "Check lock availability failed after checkout file.");
                }
            }
        }
        #endregion

        #region Abstract Methods
        /// <summary>
        /// Override to initialize the shared context based on the specified request file URL, user name, password and domain.
        /// </summary>
        /// <param name="requestFileUrl">Specify the request file URL.</param>
        /// <param name="userName">Specify the user name.</param>
        /// <param name="password">Specify the password.</param>
        /// <param name="domain">Specify the domain.</param>
        protected abstract void InitializeContext(string requestFileUrl, string userName, string password, string domain);

        /// <summary>
        /// Override to merge the common configuration and should/man configuration file. 
        /// </summary>
        /// <param name="site">An instance of interface ITestSite which provides logging, assertions,
        /// and adapters for test code onto its execution context.</param>
        protected abstract void MergeConfigurationFile(ITestSite site);
        #endregion
    }
}