namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using System.Collections.Generic;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with CellSubRequest operation.
    /// </summary>
    [TestClass]
    public abstract class S01_Cell : SharedTestSuiteBase
    {

        #region Test Suite Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            SharedTestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
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
        public void S01_CellInitialization()
        {
            // Initialize the default file URL
            this.DefaultFileUrl = this.PrepareFile();

        }

        #endregion

        #region Test Cases

        /// <summary>
        /// This method is used to test query file contents successfully.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC01_DownloadContents_Success()
        {
            Thread.Sleep(6000);
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Upload random file content to the server.
            byte[] fileContent = SharedTestSuiteHelper.GenerateRandomFileContent(this.Site);
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), fileContent);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                     ErrorCodeType.Success,
                     SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                     "Test case cannot continue unless the upload file operation succeeds.");
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site), this.Site);

            // Query the updated file content.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R97602
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         97602,
                         @"[In Cell Subrequest][The protocol returns results based on the following conditions:] ErrorCode includes ""Success"" to indicate success in processing the file [upload or ]download request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest][The protocol returns results based on the following conditions:] ErrorCode includes ""Success"" to indicate success in processing the file [upload or ]download request.");
            }

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            byte[] downloadBytes = new IntermediateNodeObject.RootNodeObjectBuilder().Build(fsshttpbResponse.DataElementPackage.DataElements, fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID)
                                                                             .GetContent()
                                                                             .ToArray();

            bool isContentMatch = AdapterHelper.ByteArrayEquals(fileContent, downloadBytes);
            this.Site.Assert.IsTrue(
                        isContentMatch,
                        "The download file contents should be equal to the updated file contents.");
        }

        /// <summary>
        /// This method is used to test whether the LastModifiedTime and the ModifiedBy is null when the GetFileProps is true/false and query changes content which is embedded in the cell sub response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC02_DownloadContents_GetFileProps()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Update the file content using the specified user1.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(Site));
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of updating file content succeeds.");

            // Query file content with the GetFileProps attribute is true, expect the server responses the LastModifiedTime and ModifiedBy.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.GetFilePropsSpecified = true;
            queryChange.SubRequestData.GetFileProps = true;
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds.");
            this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.LastModifiedTime, "The LastModifiedTime attribute cannot be null when GetFileProps in the request is true.");
            this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.CreateTime, "The LastModifiedTime attribute cannot be null when GetFileProps in the request is true.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If Create time is not null, then capture R850.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R850
                Site.CaptureRequirementIfIsNotNull(
                         cellSubResponse.SubResponseData.CreateTime,
                         "MS-FSSHTTP",
                         850,
                         @"[In CellSubResponseDataOptionalAttributes][CreateTime] The protocol server MUST return and specify the CreateTime attribute in the cell SubResponseData element only when the GetFileProps attribute is set to true in the cell subrequest.");

                // If LastModifiedTime is not null, then capture R854.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R854
                Site.CaptureRequirementIfIsNotNull(
                         cellSubResponse.SubResponseData.LastModifiedTime,
                         "MS-FSSHTTP",
                         854,
                         @"[In CellSubResponseDataOptionalAttributes][LastModifiedTime] The protocol server MUST return and specify the LastModifiedTime attribute in the cell SubResponseData element only when the GetFileProps attribute is set to true in the cell subrequest.");

                // when the LastModifiedTime and the ModifiedBy attribute is not null
                // capture requirement R804 and R855
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R804
                Site.CaptureRequirementIfIsTrue(
                         cellSubResponse.SubResponseData.LastModifiedTime != null && cellSubResponse.SubResponseData.CreateTime != null,
                         "MS-FSSHTTP",
                         804,
                         @"[In CellSubRequestDataOptionalAttributes] When [GetFileProps]set to true, file properties have been requested as part of the cell subrequest.");

                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Expect the ModifiedBy attribute contains the user name {0}, actual value is {1}",
                    this.UserName01,
                    cellSubResponse.SubResponseData.ModifiedBy);

                // If the modification contains the user name string then capture requirement R855.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R855
                Site.CaptureRequirementIfIsTrue(
                         cellSubResponse.SubResponseData.ModifiedBy.IndexOf(this.UserName01, StringComparison.OrdinalIgnoreCase) >= 0,
                         "MS-FSSHTTP",
                         855,
                         @"[In CellSubResponseDataOptionalAttributes] ModifiedBy: A UserNameType that specifies the user name for the protocol client that last modified the file.");
            }
            else
            {
                Site.Assert.IsNotNull(
                    cellSubResponse.SubResponseData.CreateTime,
                    @"[In CellSubResponseDataOptionalAttributes][CreateTime] The protocol server MUST return and specify the CreateTime attribute in the cell SubResponseData element only when the GetFileProps attribute is set to true in the cell subrequest.");

                Site.Assert.IsNotNull(
                    cellSubResponse.SubResponseData.LastModifiedTime,
                    @"[In CellSubResponseDataOptionalAttributes][LastModifiedTime] The protocol server MUST return and specify the LastModifiedTime attribute in the cell SubResponseData element only when the GetFileProps attribute is set to true in the cell subrequest.");

                Site.Log.Add(
                    LogEntryKind.Comment,
                    "Expect the ModifiedBy attribute contains the user name {0}, actual value is {1}",
                    this.UserName01,
                    cellSubResponse.SubResponseData.ModifiedBy);

                Site.Assert.IsTrue(
                    cellSubResponse.SubResponseData.ModifiedBy.IndexOf(this.UserName01, StringComparison.OrdinalIgnoreCase) >= 0,
                    @"[In CellSubResponseDataOptionalAttributes] ModifiedBy: A UserNameType that specifies the user name for the client that last modified the file.");
            }

            queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.GetFileProps = false;
            queryChange.SubRequestData.GetFilePropsSpecified = true;

            // Query file content with the GetFileProps attribute is false, expect the server does not respond the LastModifiedTime and ModifiedBy.
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.IsNull(cellSubResponse.SubResponseData.LastModifiedTime, "The LastModifiedTime attribute MUST be null when ExpectNoFileExists in the request is false.");
            this.Site.Assert.IsNull(cellSubResponse.SubResponseData.CreateTime, "The create time attribute MUST be null when ExpectNoFileExists in the request is false.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // when the LastModifiedTime and the ModifiedBy attribute is null
                // capture requirement R805
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R805
                Site.CaptureRequirementIfIsTrue(
                         cellSubResponse.SubResponseData.LastModifiedTime == null && cellSubResponse.SubResponseData.CreateTime == null,
                         "MS-FSSHTTP",
                         805,
                         @"[In CellSubRequestDataOptionalAttributes] When [GetFileProps]set to false, file properties have not been requested as part of the cell subrequest.");
            }
            else
            {
                Site.Assert.IsTrue(
                    cellSubResponse.SubResponseData.LastModifiedTime == null && cellSubResponse.SubResponseData.CreateTime == null,
                    @"[In CellSubRequestDataOptionalAttributes] When [GetFileProps]set to false, file properties have not been requested as part of the cell subrequest.");
            }
        }

        /// <summary>
        /// This method is used to test BinaryDataSize attribute will be ignored by the protocol server.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC03_DownloadContents_DifferentBinaryDataSize()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());

            // Make the BinaryDataSize is different with the actual value.
            queryChange.SubRequestData.BinaryDataSize = 1;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server returns the error code "Success", then it indicates that the value of BinaryDataSize does not affect the server behavior.
                // In this case, the requirement MS-FSSHTTP_R15231 can be captured.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R15231
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         15231,
                         @"[In CellSubRequestDataType] The server returns ""Success"" error code when the client sends a BinaryDataSize value which is different with the actual value.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataType] The server returns ""Success"" error code when the client sends a BinaryDataSize value which is different with the actual value.");
            }
        }

        /// <summary>
        /// This method is used to test the ByPassLockID will not be specified when the file is locked and retrieving the file content.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC04_DownloadContents_SchemaLock_ByPassLockIDNotSpecified()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Query the file without the BypassLockID, expect the server responds success.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds when the file is locked and ByPassLockID is not specified.");

            FsshttpbResponse fsshttpResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            Site.Assert.IsFalse(fsshttpResponse.Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary response status should be false, which indicates the response succeed.");
            Site.Assert.IsFalse(fsshttpResponse.CellSubResponses[0].Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary sub response status should be false, which indicates the query changes sub response succeed.");
        }

        /// <summary>
        /// This method is used to test the ByPassLockID will be specified when the file is locked and retrieving the file content.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC05_DownloadContents_SchemaLock_ByPassLockIDSpecified()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.BypassLockID = subRequest.SubRequestData.SchemaLockID;
            queryChange.SubRequestData.CoalesceSpecified = true;
            queryChange.SubRequestData.Coalesce = true;

            // Query the file content with the BypassLockID same as the schema lock ID used in the previous step, expect the server responds success.
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds when the file is locked and ByPassLockID is not specified.");

            FsshttpbResponse fsshttpResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            Site.Assert.IsFalse(fsshttpResponse.Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary response status should be false, which indicates the response succeed.");
            Site.Assert.IsFalse(fsshttpResponse.CellSubResponses[0].Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary sub response status should be false, which indicates the query changes sub response succeed.");
        }

        /// <summary>
        /// This method is used to test the ByPassLockID will not be specified when the file is locked by exclusive lock and retrieving the file content.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC06_DownloadContents_ExclusiveLock_ByPassLockIDNotSpecified()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with all valid parameters, expect the server responses the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Get Lock of ExclusiveLock sub request succeeds.");

            // Record the current file status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            // Query the file without the BypassLockID, expect the server responses success.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds when the file is locked and ByPassLockID is not specified.");

            FsshttpbResponse fsshttpResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            Site.Assert.IsFalse(fsshttpResponse.Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary response status should be false, which indicates the response succeed.");
            Site.Assert.IsFalse(fsshttpResponse.CellSubResponses[0].Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary sub response status should be false, which indicates the query changes sub response succeed.");
        }

        /// <summary>
        /// This method is used to test the ByPassLockID be specified when the file is locked by exclusive lock and retrieving the file content.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC07_DownloadContents_ExclusiveLock_ByPassLockIDSpecified()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);

            // Get the exclusive lock with all valid parameters, expect the server responses the error code "Success".
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Get Lock of ExclusiveLock sub request succeeds.");

            // Record the current file status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            // Query the file content with the BypassLockID same as the exclusive lock ID used in the previous step, expect the server responses success.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds when the file is locked and ByPassLockID is not specified.");

            FsshttpbResponse fsshttpResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            Site.Assert.IsFalse(fsshttpResponse.Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary response status should be false, which indicates the response succeed.");
            Site.Assert.IsFalse(fsshttpResponse.CellSubResponses[0].Status, "When the file is locked and ByPassLockID is not specified, the operation of retrieving file binary sub response status should be false, which indicates the query changes sub response succeed.");
        }

        /// <summary>
        /// This method is used to test retrieving the contents when the ETag value specified by the client is valid or invalid.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC08_DownloadContents_ETag()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query the file content with ETag value is not specified, expect the server responses success.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of retrieving file content succeeds");

            // Query the file content with valid ETag
            queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.Etag = cellSubResponse.SubResponseData.Etag;
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                                ErrorCodeType.Success,
                                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                                "Expect the error code 'Success' when the ETag specified by the client equals to the ETag stored in the server");

            this.Site.Assert.IsFalse(fsshttpbResponse.Status, "FSSHTTPB binary response cannot contain error data.");
            this.Site.Assert.IsFalse(fsshttpbResponse.CellSubResponses[0].Status, "Query changes sub response cannot contain error data.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // For the above asserts succeed, which indicates the server will check the ETag value and returns success when the ETag specified by the client equals the value stored by the server.
                // In this case the requirement R818 can be captured.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R818
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         818,
                         @"[In CellSubRequestDataOptionalAttributes] Any time the protocol client specifies the Etag attribute in a cell subrequest, the server MUST check to ensure that the Etag sent by the protocol client matches the Etag specified for that file on the server.");
            }

            // Query the file content with invalid ETag
            queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.Etag = System.Guid.NewGuid().ToString();
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals CellRequestFail, then capture requirement MS-FSSHTTP_R819.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R819
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         819,
                         @"[In CellSubRequestDataOptionalAttributes] If there is a mismatch[Etag sent by the client doesn't match the Etag specified for that file on the server], the protocol server MUST send an error code value set to ""CellRequestFail"" in the cell subresponse message.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.CellRequestFail,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataOptionalAttributes] If there is a mismatch[Etag sent by the client doesn't match the Etag specified for that file on the server], the protocol server MUST send an error code value set to ""CellRequestFail"" in the cell subresponse message.");
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents succeeds. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC09_UploadContents_Success()
        {
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R97601
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         97601,
                         @"[In Cell Subrequest][The protocol returns results based on the following conditions:] ErrorCode includes ""Success"" to indicate success in processing the file upload [or download] request.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest][The protocol returns results based on the following conditions:] ErrorCode includes ""Success"" to indicate success in processing the file upload [or download] request.");
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents succeeds when the file has a shared lock and the ByPassLockID is specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC10_UploadContents_SchemaLockSuccess()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Update contents using the same SchemaLockId and ByPassLockID as the previous step's SchemaLockId when the coalesce is true.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.SchemaLockID = subRequest.SubRequestData.SchemaLockID;
            putChange.SubRequestData.BypassLockID = subRequest.SubRequestData.SchemaLockID;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by schema lock and the ByPassLockID is specified by the valid schema lock id, the server returns the error code success.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server responds with the error code "Success", 
                // when the above steps show that the client has got a schema lock and the PutChange subrequest was sent with BypassLockID equal to SchemaLockId, 
                // then requirement MS-FSSHTTP_R834 is captured.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         834,
                         @"[In CellSubRequestDataOptionalAttributes][BypassLockID] If a client has got a shared lock, this value[BypassLockID] MUST be the same as the value of SchemaLockId, as specified in section 2.3.1.1.");

                // If the server responds with the error code "Success", 
                // when the above steps show that the schema lock identifier is specified in the cell subrequest and the file currently has a shared lock with the specified schema lock identifier,
                // then requirement MS-FSSHTTP_R2212 is captured.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2212,
                         @"[In Cell Subrequest] When the schema lock identifier is specified in the cell subrequest, the protocol server returns a success error code if the file currently has a shared lock with the specified schema lock identifier.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataOptionalAttributes][BypassLockID] If a client has got a shared lock, this value[BypassLockID] MUST be the same as the value of SchemaLockId, as specified in section 2.3.1.1.");
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents succeeds when the file has a shared lock and the ByPassLockID is not specified or specified by incorrect value.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC11_UploadContents_SchemaLockFail()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server responds with the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Update contents without ByPassLockID.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.BypassLockID = null;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by schema lock and the ByPassLockID is not specified, the server returns the error code not equals to success.");

            // Update contents with ByPassLockID specified invalid.
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.BypassLockID = System.Guid.NewGuid().ToString();
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by schema lock and the ByPassLockID is not specified, the server returns the error code not equals to success.");
        }

        /// <summary>
        /// This method is used to test uploading file contents succeeds when the file has an exclusive lock and the specified by incorrect value.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC12_UploadContents_ExclusiveLockFail()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock with all valid parameters, expect the server returns the error code "Success".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Get Lock of ExclusiveLock sub request succeeds.");

            // Record the current file status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            // Update contents with ByPassLockID specified invalid.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.BypassLockID = System.Guid.NewGuid().ToString();
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by schema lock and the ByPassLockID is not specified, the server returns the error code not equals to success.");
        }

        /// <summary>
        /// This method is used to test retrieving the content when sending the partition id in MS-FSSHTTP format for the server version is 2.2 or higher;
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC13_DownloadContents_Success_PartitionID()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // A client using protocol version 2.2 uploads content with PartitionID attribute specified.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryChange.SubRequestData.PartitionID = "00000000-0000-0000-0000-000000000000";
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "The test case cannot continue unless the update file contents succeed.");

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
                queryChange.SubRequestData.PartitionID = "00000000-0000-0000-0000-000000000000";
                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange }, "1", 2, 0, null, null);
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If the server responds with "Success" to both client 2.0 and 2.2 (the above step), then requirement MS-FSSHTTP_R1829 is captured.
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1829,
                             @"[In CellSubRequestDataOptionalAttributes][PartitionID] A protocol server that has a version number of 2.2 MUST accept a PartitionID attribute from clients using protocol numbers of 2.0 and 2.2.[The value of PartitionID for the file contents is ""00000000-0000-0000-0000-000000000000""]");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"[In CellSubRequestDataOptionalAttributes][PartitionID] A protocol server that has a version number of 2.2 MUST accept a PartitionID attribute from clients using protocol numbers of 2.0 and 2.2.[The value of PartitionID for the file contents is ""00000000-0000-0000-0000-000000000000""]");
                }
            }

            // If the server version is 2.2, it also should support the partition id defined in the MS-FSSHTTPB.
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                ExGuid storageIndexExGuid;
                List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(Site), out storageIndexExGuid);
                QueryChangesCellSubRequest queryChangeFsshttpb = new QueryChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
                cellRequest.AddSubRequest(queryChangeFsshttpb, dataElements);
                queryChangeFsshttpb.IsPartitionIDGUIDUsed = true;
                queryChangeFsshttpb.PartitionIdGUID = Guid.Empty;

                // Do not send the partition id defined in the MS-FSSHTTP
                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                cellSubRequest.SubRequestData.PartitionID = null;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "For the server version is number of 2.2 or greater, the server should accept the Partition id attribute defined in the MS-FSSHTTPB");
            }
        }

        /// <summary>
        /// This test case is used to test uploading file contents when the ETag value is valid.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC14_UploadContents_ValidEtag()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke the query change to get the valid ETag value.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                   ErrorCodeType.Success,
                   SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                   "The test case cannot continue until the query change error code equals success.");

            // Save the ETag value 
            string previousETagValue = cellSubResponse.SubResponseData.Etag;
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.Etag = previousETagValue;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                                ErrorCodeType.Success,
                                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                                "Expect the error code 'Success' when the ETag specified by the client equals to the ETag stored in the server");

            this.Site.Assert.IsFalse(fsshttpbResponse.Status, "FSSHTTPB binary response cannot contain error data.");
            this.Site.Assert.IsFalse(fsshttpbResponse.CellSubResponses[0].Status, "Put changes sub response cannot contain error data.");

            // Save the current Etag value.
            string currentETagValue = cellSubResponse.SubResponseData.Etag;

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // For the above asserts succeed, which indicates the server will check the ETag value and response succeeds when the ETag specified by the client equals the value stored by the server. In this case the requirement R818 can be captured.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R818
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         818,
                         @"[In CellSubRequestDataOptionalAttributes] Any time the protocol client specifies the Etag attribute in a cell subrequest, the server MUST check to ensure that the Etag sent by the protocol client matches the Etag specified for that file on the server.");

                // If both the Etag are not equal, then capture requirement R2087, R2088
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2087
                Site.CaptureRequirementIfAreNotEqual<string>(
                         previousETagValue,
                         currentETagValue,
                         "MS-FSSHTTP",
                         2087,
                         @"[In CellSubResponseDataOptionalAttributes] Etag: The string value is different from the previous one after the file contents are changed.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2088
                Site.CaptureRequirementIfAreNotEqual<string>(
                         previousETagValue,
                         currentETagValue,
                         "MS-FSSHTTP",
                         2088,
                         @"[In CellSubResponseDataOptionalAttributes][Etag] The string value is different from the previous one irrespective of which protocol client updated the file contents in a coauthorable file.");

                // If the ETage value is different in the first QueryChange response and second PutChanges response, then the Etag is the value which uniquely specifies the file current version.
                // R842 and R843 can be directly capture in this case.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R842
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         842,
                         @"[In CellSubResponseDataOptionalAttributes] Etag defines the file version and allows for the protocol client to know the version of the file.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R843
                Site.CaptureRequirement(
                         "MS-FSSHTTP",
                         843,
                         @"[In CellSubResponseDataOptionalAttributes] When the Etag attribute is specified as part of a response to a cell subrequest, the Etag attribute value specifies the updated file version.");
            }
            else
            {
                Site.Assert.AreNotEqual<string>(
                    previousETagValue,
                    currentETagValue,
                    @"[In CellSubResponseDataOptionalAttributes] Etag: The string value is different from the previous one after the file contents are changed.");
            }
        }

        /// <summary>
        /// This test case is used to test uploading file contents when the ETag value does not match the value stored by the server.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC15_UploadContents_InvalidEtag()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));

            // Random generate the ETag value
            putChange.SubRequestData.Etag = SharedTestSuiteHelper.GenerateRandomETag();
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error code equals CellRequestFail, then capture requirements R818 and R819.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R818
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         818,
                         @"[In CellSubRequestDataOptionalAttributes] Any time the protocol client specifies the Etag attribute in a cell subrequest, the server MUST check to ensure that the Etag sent by the protocol client matches the Etag specified for that file on the server.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R819
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         819,
                         @"[In CellSubRequestDataOptionalAttributes] If there is a mismatch[Etag sent by the client doesn't match the Etag specified for that file on the server], the protocol server MUST send an error code value set to ""CellRequestFail"" in the cell subresponse message.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.CellRequestFail,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataOptionalAttributes] Any time the protocol client specifies the Etag attribute in a cell subrequest, the server MUST check to ensure that the Etag sent by the client matches the Etag specified for that file on the server.");
            }
        }

        /// <summary>
        /// This test case is used to test when uploading contents the GetFileProps attribute does not affect the response for the attribute LastModifiedTime and CreateTime.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC16_UploadContents_GetFileProps()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Upload file content with the GetFileProps value false and Coalesce value true, expect the server does not response the LastModifiedTime and CreateTime.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.GetFilePropsSpecified = true;
            putChange.SubRequestData.GetFileProps = false;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless uploading file contents succeeds");
            this.Site.Assert.IsNull(cellSubResponse.SubResponseData.LastModifiedTime, "The LastModifiedTime attribute should be null when GetFileProps in the request is false and uploading file contents.");
            this.Site.Assert.IsNull(cellSubResponse.SubResponseData.CreateTime, "The CreateTime attribute should be null when GetFileProps in the request is false and uploading file contents.");

            // Upload file content with the GetFileProps value true and Coalesce value true, expect the server responses the LastModifiedTime and CreateTime in SharePoint Server 2010..
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.GetFilePropsSpecified = true;
            putChange.SubRequestData.GetFileProps = true;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless uploading file contents succeeds");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3089, this.Site))
                {
                    this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.LastModifiedTime, "The LastModifiedTime attribute should not be null when GetFileProps in the request is true and uploading file contents.");
                    this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.CreateTime, "The CreateTime attribute should not be null when GetFileProps in the request is true and uploading file contents.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3089
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             3089,
                             @"[In Appendix B: Product Behavior] When [GetFileProps is] set to true in Put Changes subrequest, the implementation does return CreateTime and LastModifiedTime as attributes in the cell SubResponseData element. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 3089, this.Site))
                {
                    this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.LastModifiedTime, "The LastModifiedTime attribute should not be null when GetFileProps in the request is true and uploading file contents.");
                    this.Site.Assert.IsNotNull(cellSubResponse.SubResponseData.CreateTime, "The CreateTime attribute should not be null when GetFileProps in the request is true and uploading file contents.");
                }
            }
        }

        /// <summary>
        /// This test case is used to test CoauthVersion attribute when the coauthoring status is Alone or coauthoring.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC17_UploadContents_CoauthVersion()
        {
            // User1 Join the coauthoring session.
            CoauthStatusType coauthStatus = this.PrepareCoauthoringSession(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID);
            this.Site.Assert.AreEqual<CoauthStatusType>(
                        CoauthStatusType.Alone,
                        coauthStatus,
                        "After the first user join the coauthoring session, the coauthoring status should be Alone.");

            // Uploading the file contents with CoauthVersioning false when the coauthoring status is "Alone". 
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoauthVersioningSpecified = true;
            putChange.SubRequestData.CoauthVersioning = false;
            putChange.SubRequestData.BypassLockID = SharedTestSuiteHelper.ReservedSchemaLockID;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the operation of uploading file content succeeds when the CoauthVersioning is false and the coauthoring status is Alone.");

            // User2 Join the coauthoring session.
            string secondClientId = System.Guid.NewGuid().ToString();
            coauthStatus = this.PrepareCoauthoringSession(this.DefaultFileUrl, secondClientId, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName02, this.Password02, this.Domain);
            this.Site.Assert.AreEqual<CoauthStatusType>(
                         CoauthStatusType.Coauthoring,
                         coauthStatus,
                         "After the second user join the coauthoring session, the coauthoring status should be Coauthoring.");

            // Uploading the file contents with CoauthVersioning false when the coauthoring status is "Coauthoring". 
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoauthVersioningSpecified = true;
            putChange.SubRequestData.CoauthVersioning = true;
            putChange.SubRequestData.BypassLockID = SharedTestSuiteHelper.ReservedSchemaLockID;
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the operation of uploading file content succeeds when the CoauthVersioning is true and the coauthoring status is Coauthoring.");
        }

        /// <summary>
        /// This method is used to test uploading file contents with the attribute of PartitionID. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC18_UploadContents_Success_PartitionID()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // A client using protocol version 2.2 uploads content with PartitionID attribute specified.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.PartitionID = "00000000-0000-0000-0000-000000000000";
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "The test case cannot continue unless the update file contents succeed.");

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.PartitionID = "00000000-0000-0000-0000-000000000000";
                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange }, "1", 2, 0, null, null);
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If the server responds with "Success" to both client 2.0 and 2.2 (the above step), then requirement MS-FSSHTTP_R1829 is captured.
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             1829,
                             @"[In CellSubRequestDataOptionalAttributes][PartitionID] A protocol server that has a version number of 2.2 MUST accept a PartitionID attribute from clients using protocol numbers of 2.0 and 2.2.[The value of PartitionID for the file contents is ""00000000-0000-0000-0000-000000000000""]");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"[In CellSubRequestDataOptionalAttributes][PartitionID] A protocol server that has a version number of 2.2 MUST accept a PartitionID attribute from clients using protocol numbers of 2.0 and 2.2.[The value of PartitionID for the file contents is ""00000000-0000-0000-0000-000000000000""]");
                }
            }

            // If the server version is 2.2, it also should support the partition id defined in the MS-FSSHTTPB.
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                ExGuid storageIndexExGuid;
                List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(Site), out storageIndexExGuid);
                PutChangesCellSubRequest putChangeFsshttpb = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
                cellRequest.AddSubRequest(putChangeFsshttpb, dataElements);
                putChangeFsshttpb.IsPartitionIDGUIDUsed = true;
                putChangeFsshttpb.PartitionIdGUID = Guid.Empty;

                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                cellSubRequest.SubRequestData.PartitionID = null;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "For the server version is number of 2.2 or greater, the server should accept the Partition id attribute defined in the MS-FSSHTTPB");
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 213802, this.Site))
            {
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                ExGuid storageIndexExGuid;
                List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
                PutChangesCellSubRequest putChangeFsshttpb = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
                cellRequest.AddSubRequest(putChangeFsshttpb, dataElements);
                putChangeFsshttpb.IsPartitionIDGUIDUsed = true;
                putChangeFsshttpb.PartitionIdGUID = Guid.Empty;

                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                cellSubRequest.SubRequestData.PartitionID = null;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTPB",
                             213802,
                             @"[In Appendix B: Product Behavior]  Implementation does support the Target Partition Id field. (<9> Section 2.2.2.1:  SharePoint Server 2013 and above support the Target Partition Id field.)");
                }
                else
                {
                    Site.Assert.AreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             @"[In Appendix B: Product Behavior]  Implementation does support the Target Partition Id field. (<9> Section 2.2.2.1:  SharePoint Server 2013 and above support the Target Partition Id field.)");
                }
            }
        }

        /// <summary>
        /// This test case is used to test the CoalesceHResult attribute when uploading file contents.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC19_UploadContents_CoalesceHResult()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = false;

            // Put changes to the server with the Coalesce is false to not fully saving changes, expect the server response the error code "Success" and CoalesceHResult equals 0.
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the server put changes to server success.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            this.Site.Assert.IsFalse(
                fsshttpbResponse.Status,
                "Test case cannot continue unless the server put changes to server success.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4001, this.Site))
                {
                    this.Site.Assert.IsNotNull(
                                    cellSubResponse.SubResponseData.CoalesceHResult,
                                    "Test case cannot continue unless the CoalesceHResult attribute is specified.");

                    // If the CoalesceHResult equals 0, then capture the requirement R1526 and R4001.
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1526
                    Site.CaptureRequirementIfAreEqual<int>(
                             0,
                             Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                             "MS-FSSHTTP",
                             1526,
                             @"[In CellSubResponseDataOptionalAttributes] CoalesceHResult: An integer that MUST be 0 except when the protocol server attempts to fully save all the changes in the underlying store. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4001
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             4001,
                             @"[In Appendix B: Product Behavior] Attribute[CoalesceHResult] is supported by Office 2010.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4001, this.Site))
                {
                    Site.Assert.AreEqual<int>(
                        0,
                        Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                        @"[In CellSubResponseDataOptionalAttributes] CoalesceHResult: An integer that MUST be 0 except when the protocol server attempts to fully save all the changes in the underlying store.");
                }
            }

            // Put changes to the server with the Coalesce is true, expect the server returns the error code "Success" and CoalesceHResult equals 0.
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the server put changes to server success.");

            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            this.Site.Assert.IsFalse(
                fsshttpbResponse.Status,
                "Test case cannot continue unless the server put changes to server success.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4001, this.Site))
                {
                    this.Site.Assert.IsNotNull(
                                cellSubResponse.SubResponseData.CoalesceHResult,
                                "Test case cannot continue unless the CoalesceHResult attribute is specified.");

                    // If the CoalesceHResult equals 0, then capture the requirement R1527 and R1529.
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1527
                    Site.CaptureRequirementIfAreEqual<int>(
                             0,
                             Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                             "MS-FSSHTTP",
                             1527,
                             @"[In CellSubResponseDataOptionalAttributes] It[CoalesceHResult] specifies the HRESULT when the protocol server attempts to fully save all the changes in the underlying store. (Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010 follow this behavior.)");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1529
                    Site.CaptureRequirementIfAreEqual<int>(
                             0,
                             Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                             "MS-FSSHTTP",
                             1529,
                             @"[In CellSubResponseDataOptionalAttributes][CoalesceHResult] A CoalesceHResult value of 0 indicates success.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4001, this.Site))
                {
                    Site.Assert.AreEqual<int>(
                         0,
                         Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                         @"[In CellSubResponseDataOptionalAttributes] It[CoalesceHResult] specifies the HRESULT when the protocol server attempts to fully save all the changes in the underlying store.");
                }
            }

            // Prepare a shared Lock.
            this.PrepareSharedLock(this.DefaultFileUrl, SharedTestSuiteHelper.DefaultClientID, SharedTestSuiteHelper.ReservedSchemaLockID, this.UserName01, this.Password01, this.Domain);

            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            putChange.SubRequestData.BypassLockID = null;

            // Put changes to the server with the Coalesce set to true and without specifying ByPassLockID, expect the server does not returns the error code "Success" and CoalesceHResult does not equal 0.
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        "Test case cannot constituent unless the server put changes to server fails.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4002, this.Site))
                {
                    this.Site.Assert.IsNotNull(
                        cellSubResponse.SubResponseData.CoalesceHResult,
                        "When the put changes fail and the Coalesce is specified to true in the request, the CoalesceHResult should be not null.");

                    this.Site.Assert.IsTrue(
                                Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult) != 0,
                                "When the put changes fail and the Coalesce is specified to true in the request, the CoalesceHResult should not equal to 0, the value is {0}.",
                                cellSubResponse.SubResponseData.CoalesceHResult);

                    Site.CaptureRequirementIfAreNotEqual<int>(
                             0,
                             Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult),
                             "MS-FSSHTTP",
                             3097,
                             @"[In CellSubResponseDataOptionalAttributes][CoalesceHResult] If CoalesceHResult is not equal to 0, it indicates an exception or failure condition that occurred. <32>");

                    bool isR1528Verified = Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult) >= -2147483648 && Convert.ToInt32(cellSubResponse.SubResponseData.CoalesceHResult) <= 2147483647;
                    Site.Log.Add(
                        TestTools.LogEntryKind.Debug,
                        string.Format("The CoalesceHResult should in the range -2147483648 to 2147483647, the actual value {0}", cellSubResponse.SubResponseData.CoalesceHResult));

                    Site.CaptureRequirementIfIsTrue(
                             isR1528Verified,
                             "MS-FSSHTTP",
                             1528,
                             @"[In CellSubResponseDataOptionalAttributes][CoalesceHResult] CoalesceHResult MUST be set to a value ranging from -2,147,483,648 through 2,147,483,647.");

                    // If the CoalesceHResult is greater than 0, and the CoalesceErrorMessage is not null, then capture R3095.
                    Site.CaptureRequirementIfIsNotNull(
                             cellSubResponse.SubResponseData.CoalesceErrorMessage,
                             "MS-FSSHTTP",
                             3095,
                             @"[In CellSubResponseDataOptionalAttributes][CoalesceErrorMessage] CoalesceErrorMessage MUST be sent only when the CoalesceHResult attribute is set to an integer value which is not equal to 0. <31>");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R4002
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             4002,
                             @"[In Appendix B: Product Behavior] Attribute[CoalesceErrorMessage] is supported by Office 2010.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4002, this.Site))
                {
                    Site.Assert.IsNotNull(
                        cellSubResponse.SubResponseData.CoalesceErrorMessage,
                        @"[In Appendix B: Product Behavior] The implementation does send CoalesceErrorMessage when the CoalesceHResult attribute is set to an integer value which is not equal to 0. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Foundation 2013/Microsoft SharePoint Server 2010/Microsoft SharePoint Server 2013/Microsoft SharePoint Workspace 2010 follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// This method is used to test BinaryDataSize attribute will be ignored by the protocol server when uploading file contents.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC20_UploadContents_DifferentBinaryDataSize()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));

            // Make the BinaryDataSize is different with the actual value.
            putChange.SubRequestData.BinaryDataSize = 1;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server returns the error code "Success", then it indicates that the value of BinaryDataSize does not affect the server behavior. In this case, the requirement MS-FSSHTTP_R15231 can be captured.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R15231
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         15231,
                         @"[In CellSubRequestDataType] The server returns ""Success"" error code when the client sends a BinaryDataSize value which is different with the actual value.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataType] The server returns ""Success"" error code when the client sends a BinaryDataSize value which is different with the actual value.");
            }

            // Make the BinaryDataSize different with the value in last step.
            putChange.SubRequestData.BinaryDataSize = 2;
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server also responds with the error code "Success" the same as the last step, then it indicates that different value of BinaryDataSize does not affect the server behavior. In this case, the requirement MS-FSSHTTP_R3044 can be captured.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3044
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3044,
                         @"[In SubRequestDataOptionalAttributes][BinaryDataSize] When the BinaryDataSize is set to two different values, the server responds the same.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In SubRequestDataOptionalAttributes][BinaryDataSize] When the BinaryDataSize is set to two different values, the server responds the same.");
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents with the Lock ID attribute defined in MS-FSSHTTPB for the server of the version is 2.2 or higher.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC21_UploadContents_Success_LockID()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                // Update contents using the same SchemaLockId and ByPassLockID as the previous step's SchemaLockId when the coalesce is true.
                CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = subRequest.SubRequestData.SchemaLockID;
                putChange.SubRequestData.BypassLockID = subRequest.SubRequestData.SchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        "When the file is locked by schema lock and the ByPassLockID is specified by the valid schema lock id, the server returns the error code success.");

                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                ExGuid storageIndexExGuid;
                List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(Site), out storageIndexExGuid);
                PutChangesCellSubRequest putChangeFsshttpb = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
                cellRequest.AddSubRequest(putChangeFsshttpb, dataElements);
                putChangeFsshttpb.IsLockIdUsed = true;
                putChangeFsshttpb.LockID = new Guid(subRequest.SubRequestData.SchemaLockID);

                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
                cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "For the server version is number of 2.2 or greater, the server should accept the lock id attribute defined in the MS-FSSHTTPB");

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1834
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             1834,
                             @"[In CellSubRequestDataOptionalAttributes][BypassLockID] A protocol server that has a version number of 2.2 MUST accept both LockID and BypassLockID.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R518
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             518,
                             @"[In CellSubRequestDataType][SchemaLockID] After a protocol client is able to get a shared lock for a file with a specific schema lock identifier, the server MUST allow only other protocol clients that specify the same schema lock identifier to share the file lock.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R519
                    Site.CaptureRequirement(
                             "MS-FSSHTTP",
                             519,
                             @"[In CellSubRequestDataType] The protocol server ensures that at any instant in time, only clients having the same schema lock identifier can lock the file.");
                }
            }
        }

        /// <summary>
        /// This test case is used to test the ExpectNoFileExists attribute when uploading file contents. 
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC22_UploadContents_ExpectNoFileExists()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.ExpectNoFileExistsSpecified = true;
            putChange.SubRequestData.ExpectNoFileExists = true;
            putChange.SubRequestData.Etag = string.Empty;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange }, "1",
                2, 2, null, null, null, null, true);
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1869
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1869,
                         @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST cause the cell subrequest to fail with a coherency error if the file already exists on the server.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1102401, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1101401
                    Site.CaptureRequirementIfIsNotNull(
                             response.ResponseCollection.Response[0].SuggestedFileName,
                             "MS-FSSHTTP",
                             1101401,
                             @"[In Request] ShouldReturnDisambiguatedFileName: If an upload request fails with a coherency failure, this flag [is true] specifies the host should return a suggested/available file name that the client can try instead.");

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1102401
                    Site.CaptureRequirementIfIsNotNull(
                        response.ResponseCollection.Response[0].SuggestedFileName,
                        "MS-FSSHTTP",
                        1102401,
                        @"[In Appendix B: Product Behavior] Implementation does support SuggestedFileName to specify that the suggested filename that the host returns if the ShouldReturnDisambiguatedFileName flag is set on the Request. (Microsoft SharePoint Server 2016 and above support this behavior.)");
                }
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.CellRequestFail,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST cause the cell subrequest to fail with a coherency error if the file already exists on the server.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1102401, this.Site))
                {
                    Site.Assert.IsNotNull(
                    response.ResponseCollection.Response[0].SuggestedFileName,
                    "[In Request] ShouldReturnDisambiguatedFileName: If an upload request fails with a coherency failure, this flag [is true] specifies the host should return a suggested/available file name that the client can try instead.");
                }
            }

            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange }, "1",
                2, 2, null, null, null, null, false);
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1869
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         1869,
                         @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST cause the cell subrequest to fail with a coherency error if the file already exists on the server.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1102401, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R1101402
                    Site.CaptureRequirementIfIsNull(
                         response.ResponseCollection.Response[0].SuggestedFileName,
                         "MS-FSSHTTP",
                         1101402,
                         @"[In Request] ShouldReturnDisambiguatedFileName: If an upload request fails with a coherency failure, this flag [is false] specifies the host should not return a suggested/available file name that the client can try instead.");
                }

            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.CellRequestFail,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST cause the cell subrequest to fail with a coherency error if the file already exists on the server.");

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 1102401, this.Site))
                {
                    Site.Assert.IsNull(
                    response.ResponseCollection.Response[0].SuggestedFileName,
                    "[In Request] ShouldReturnDisambiguatedFileName: If an upload request fails with a coherency failure, this flag [is false] specifies the host should not return a suggested/available file name that the client can try instead.");
                }
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents when the file has a shared lock and the SchemaLockID is other schema lock id.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC23_UploadContents_DifferentSchemaLockID()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                // Initialize the context using user02 and defaultFileUrl.
                this.InitializeContext(this.DefaultFileUrl, this.UserName02, this.Password02, this.Domain);

                string otherSchemaLockID = Guid.NewGuid().ToString();
                // Update contents using the same SchemaLockId and ByPassLockID as the previous step's SchemaLockId when the coalesce is true.
                CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = otherSchemaLockID;
                putChange.SubRequestData.BypassLockID = otherSchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // If the server does not return the error code "Success", then it indicates that the protocol server to block the other client with different schema lock identifier. In this case, the requirement MS-FSSHTTP_R517 can be captured.
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R517
                    Site.CaptureRequirementIfAreNotEqual<ErrorCodeType>(
                             ErrorCodeType.Success,
                             SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             "MS-FSSHTTP",
                             517,
                             @"[In CellSubRequestDataType][SchemaLockID] This schema lock identifier is used by the protocol server to block other clients with a different schema lock identifier.");
                }
                else
                {
                    Site.Assert.AreNotEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"[In CellSubRequestDataType][SchemaLockID] This schema lock identifier is used by the protocol server to block other clients with a different schema lock identifier.");
                }
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents when the file has a shared lock and the ByPassLockID is not set or not same with this schema lock identified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC24_UploadContents_SchemaLockIDIgnored()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                string otherSchemaLockID = Guid.NewGuid().ToString();
                // Update contents and the ByPassLockID not same with SchemaLockId.
                CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = subRequest.SubRequestData.SchemaLockID;
                putChange.SubRequestData.BypassLockID = otherSchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponseNotSameWithSchemaLockId1 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = subRequest.SubRequestData.SchemaLockID;
                putChange.SubRequestData.BypassLockID = otherSchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponseNotSameWithSchemaLockId2 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseNotSameWithSchemaLockId1.ErrorCode, this.Site),
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseNotSameWithSchemaLockId2.ErrorCode, this.Site),
                         "If the ByPassLockID is not same with this schema lock identified, the SchemaLockID will be ignored by the server.");

                // Update contents and the ByPassLockID not set.
                putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = subRequest.SubRequestData.SchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponseNotSetByPassLockID1 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                // Update contents and the ByPassLockID not set.
                putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponseNotSetByPassLockID2 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseNotSetByPassLockID1.ErrorCode, this.Site),
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseNotSetByPassLockID2.ErrorCode, this.Site),
                         "If the ByPassLockID is not set, the SchemaLockID will be ignored by the server.");

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11278
                    this.Site.CaptureRequirement(
                             11278,
                            "[In CellSubRequestDataType][SchemaLockID] if the ByPassLockID is not set or not same with this schema lock identified, the SchemaLockID will be ignored by the server.");
                }
            }
        }

        /// <summary>
        /// This method is used to test uploading file contents after all the protocol clients have released their lock for that file.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC25_UploadContents_AfterReleaseSchemaLock()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get a schema lock with all valid parameters, expect the server returns the error code "Success".
            SchemaLockSubRequestType subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.GetLock, false, null);
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            SchemaLockSubResponseType schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);

            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.RecordSchemaLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // Release lock with same ClientId and SchemaLockId with first step, expect server responses the error code "Success".
            subRequest = SharedTestSuiteHelper.CreateSchemaLockSubRequest(SchemaLockRequestTypes.ReleaseLock, null, null);
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            schemaLockSubResponse = SharedTestSuiteHelper.ExtractSubResponse<SchemaLockSubResponseType>(response, 0, 0, this.Site);
            Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(schemaLockSubResponse.ErrorCode, this.Site), "Test case cannot continue unless the Get Lock of SchemaLock sub request succeeds.");
            this.StatusManager.CancelSharedLock(this.DefaultFileUrl, subRequest.SubRequestData.ClientID, subRequest.SubRequestData.SchemaLockID);

            // If the server version is 2.2, it should also accept the client version 2.0
            if (response.ResponseVersion.Version >= 2 && response.ResponseVersion.MinorVersion >= 2)
            {
                string otherSchemaLockID = Guid.NewGuid().ToString();
                // Update contents using the different schema lock ID.
                CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
                putChange.SubRequestData.SchemaLockID = otherSchemaLockID;
                putChange.SubRequestData.BypassLockID = otherSchemaLockID;
                putChange.SubRequestData.CoalesceSpecified = true;
                putChange.SubRequestData.Coalesce = true;

                response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
                CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

                if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R520
                    this.Site.CaptureRequirementIfAreEqual<ErrorCodeType>(                          
                            ErrorCodeType.Success,
                            SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                             520,
                            "[In CellSubRequestDataType] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
                }
                else
                {
                    this.Site.Assert.AreEqual<ErrorCodeType>(
                            ErrorCodeType.Success,
                            SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                            "[In CellSubRequestDataType] After all the protocol clients have released their lock for that file, the protocol server MUST allow a protocol client with a different schema lock identifier to get a shared lock for that file.");
                }
            }
        }

        /// <summary>
        /// This method is used to test whether protocol server save the file with LastModifiedTime value in CellSubRequest as the LastModifiedTime instead of the current time.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S01_TC26_UploadContents_LastModifiedTime()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Update the file content using the specified user1.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(Site));
            putChange.SubRequestData.GetFilePropsSpecified = true;
            putChange.SubRequestData.GetFileProps = true;

            //set LastModifiedTime to null 
            CellStorageResponse response1 = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse1 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response1, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse1.ErrorCode, this.Site), "Test case cannot continue unless the operation of updating file content succeeds.");
            string lastModifiedTime1 = cellSubResponse1.SubResponseData.LastModifiedTime;

            //set LastModifiedTime to lastModifiedTime1 
            CellStorageResponse response2 = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange }, "1", 2, 2, null, null, lastModifiedTime1, null, null, null, null);
            CellSubResponseType cellSubResponse2 = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response2, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse2.ErrorCode, this.Site), "Test case cannot continue unless the operation of updating file content succeeds.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If lastModifiedTime1 equals cellSubResponse2.SubResponseData.LastModifiedTime, then capture R11215.
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11215
                Site.CaptureRequirementIfIsTrue(
                         lastModifiedTime1 == cellSubResponse2.SubResponseData.LastModifiedTime,
                         "MS-FSSHTTP",
                         11215,
                         @"[In CellSubRequestDataOptionalAttributes][LastModifiedTime] The protocol server MUST save the file with this value as the LastModifiedTime instead of the current time.");
            }
            else
            {
                Site.Assert.IsTrue(
                   lastModifiedTime1 == cellSubResponse2.SubResponseData.LastModifiedTime,
                   @"[In CellSubRequestDataOptionalAttributes][LastModifiedTime] The protocol server MUST save the file with this value as the LastModifiedTime instead of the current time.");
            }
        }

        #endregion
    }
}