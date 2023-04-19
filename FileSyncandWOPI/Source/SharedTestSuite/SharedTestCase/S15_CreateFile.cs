namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System.Collections.Generic;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with CellSubRequest operation for creating new file on the server. 
    /// </summary>
    [TestClass]
    public abstract class S15_CreateFile : SharedTestSuiteBase
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

        #region Test Case Initialization

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S15_CreateFileInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
            
        }

        #endregion

        /// <summary>
        /// This method is used to test uploading a file for the first time.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S15_TC01_CreateFile()
        {
            string randomFileUrl = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(randomFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            CellStorageResponse response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When uploading contents if the protocol server was unable to find the URL for the file specified in the Url attribute,, the server returns success and create a new file using the specified contents in the file URI.");

            // Query the updated file content.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R3103
                // If queryChange subrequest can download a file from randomFileUrl, then capture MS-FSSHTTP_R3103.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         3103,
                         @"[In Cell Subrequest] But for Put Changes subrequest, as described in [MS-FSSHTTPB] section 2.2.2.1.4, [If the protocol server was unable to find the URL for the file specified in the Url attribute] the protocol server creates a new file using the specified Url.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest] But for Put Changes subrequest, as described in [MS-FSSHTTPB] section 2.2.2.1.4, [If the protocol server was unable to find the URL for the file specified in the Url attribute] the protocol server creates a new file using the specified Url.");
            }

            this.StatusManager.RecordFileUpload(randomFileUrl);

            // Re-generate the non-exist file URL again.
            randomFileUrl = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);
            this.InitializeContext(randomFileUrl, this.UserName01, this.Password01, this.Domain);

            putChange.SubRequestData.ExpectNoFileExistsSpecified = true;
            putChange.SubRequestData.ExpectNoFileExists = true;
            putChange.SubRequestData.Etag = string.Empty;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2252
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2252,
                         @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST NOT cause the cell subrequest to fail with a coherency error if the file does not exist on the server.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Cell Subrequest] In this case[If the ExpectNoFileExists attribute is set to true in a file content upload cell subrequest, the Etag attribute MUST be an empty string], the protocol server MUST NOT cause the cell subrequest to fail with a coherency error if the file does not exist on the server.");
            }

            this.StatusManager.RecordFileUpload(randomFileUrl);
        }

        /// <summary>
        /// This method is used to test uploading a file for the first time with an exclusive lock.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S15_TC02_UploadContents_CreateFile_ExclusiveLock()
        {
            string randomFileUrl = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(randomFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            putChange.SubRequestData.ExpectNoFileExistsSpecified = true;
            putChange.SubRequestData.ExpectNoFileExists = true;
            putChange.SubRequestData.ExclusiveLockID = SharedTestSuiteHelper.DefaultExclusiveLockID;
            putChange.SubRequestData.Timeout = "3600";

            CellStorageResponse response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When uploading contents if the protocol server was unable to find the URL for the file specified in the Url attribute,, the server returns success and create a new file using the specified contents in the file URI.");

            bool isExclusiveLock = this.CheckExclusiveLockExist(randomFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID, this.UserName01, this.Password01, this.Domain);

            Site.Log.Add(
                TestTools.LogEntryKind.Debug,
                "When creating a new file with ExclusiveLockID specified, the server will lock the new created file, but actually it {0}",
                isExclusiveLock ? "locks" : "does not lock");

            Site.Assert.IsTrue(
                isExclusiveLock,
                "When creating a new file with ExclusiveLockID specified, the server will lock the new created file");

            this.StatusManager.RecordExclusiveLock(randomFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID, this.UserName01, this.Password01, this.Domain);
            this.StatusManager.RecordFileUpload(randomFileUrl);
        }

        /// <summary>
        /// This method is used to test uploading a new file incrementally by the contents which is downloading from the server.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S15_TC03_Download_UploadPartial()
        {
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);
            string uploadFileUrl = SharedTestSuiteHelper.GenerateNonExistFileUrl(Site);
            bool partial = false;

            Knowledge knowledge = null;

            // Set the limit number of upload tries, this will allow 500000 * 10 bytes size file to be download and complete upload.
            int limitNumberOfPartialUpload = 10;
            do
            {
                this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

                // Create query changes request with allow fragments flag with the value false.
                FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                QueryChangesCellSubRequest queryChange = SharedTestSuiteHelper.BuildFsshttpbQueryChangesSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), 0, false, false, true, 0, true, true, 0, null, 500000, null, knowledge);
                cellRequest.AddSubRequest(queryChange, null);
                CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

                CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
                CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the query changes succeed.");

                FsshttpbResponse queryResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
                SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryResponse, this.Site);
                QueryChangesSubResponseData data = queryResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>();
                partial = data.PartialResult;
                knowledge = data.Knowledge;

                this.InitializeContext(uploadFileUrl, this.UserName01, this.Password01, this.Domain);
                cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
                PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), null);
                putChange.Partial = partial ? 1 : 0;
                putChange.PartialLast = partial ? 0 : 1;
                putChange.StorageIndexExtendedGUID = partial ? null : data.StorageIndexExtendedGUID;

                if (partial)
                {
                    var storageIndex = queryResponse.DataElementPackage.DataElements.FirstOrDefault(e => e.DataElementType == DataElementType.StorageIndexDataElementData);
                    if (storageIndex != null)
                    {
                        queryResponse.DataElementPackage.DataElements.Remove(storageIndex);
                    }
                }

                cellRequest.AddSubRequest(putChange, queryResponse.DataElementPackage.DataElements);
                cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

                cellStorageResponse = this.Adapter.CellStorageRequest(uploadFileUrl, new SubRequestType[] { cellSubRequest });
                subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
                this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site),
                    "Test case cannot continue unless the query changes succeed.");

                FsshttpbResponse putResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);
                SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putResponse, this.Site);

                // Decrease the number of upload tries.
                limitNumberOfPartialUpload--;
            }
            while (partial && limitNumberOfPartialUpload > 0);

            this.StatusManager.RecordFileUpload(uploadFileUrl);
        }

        /// <summary>
        /// This method is used to test uploading file contents succeeds when the file has an exclusive lock and the ByPassLockID is specified or not specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S15_TC04_UploadContents_ExclusiveLockSuccess()
        {
            string randomFileUrl = SharedTestSuiteHelper.GenerateNonExistFileUrl(this.Site);

            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(randomFileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;
            putChange.SubRequestData.ExpectNoFileExistsSpecified = true;
            putChange.SubRequestData.ExpectNoFileExists = true;
            putChange.SubRequestData.ExclusiveLockID = SharedTestSuiteHelper.DefaultExclusiveLockID;
            putChange.SubRequestData.BypassLockID = putChange.SubRequestData.ExclusiveLockID;
            putChange.SubRequestData.Timeout = "3600";

            CellStorageResponse response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by exclusive lock and the ByPassLockID is specified by the valid exclusive lock id, the server returns the error code success.");

            this.StatusManager.RecordExclusiveLock(randomFileUrl, SharedTestSuiteHelper.DefaultExclusiveLockID, this.UserName01, this.Password01, this.Domain);
            this.StatusManager.RecordFileUpload(randomFileUrl);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the server responds with the error code "Success", 
                // when the above steps show that the client has got an exclusive lock and the PutChange subrequest was sent with BypassLockID equal to ExclusiveLockID, 
                // then requirement MS-FSSHTTP_R833 is captured.
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         833,
                         @"[In CellSubRequestDataOptionalAttributes][BypassLockID] If a client has got an exclusive lock, this value[BypassLockID] MUST be the same as the value of ExclusiveLockID, as specified in section 2.3.1.1.");

                // If the server responds with "ExclusiveLock" in LockType attribute, the requirement MS-FSSHTTP_R1533 is captured. 
                Site.CaptureRequirementIfAreEqual<string>(
                         "ExclusiveLock",
                         cellSubResponse.SubResponseData.LockType.ToString(),
                         "MS-FSSHTTP",
                         1533,
                         @"[In CellSubResponseDataType] The LockType attribute MUST be set to ""ExclusiveLock"" in the cell subresponse if the ExclusiveLockID attribute is sent in the cell subrequest and the protocol server is successfully able to take an exclusive lock.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In CellSubRequestDataOptionalAttributes][BypassLockID] If a client has got an exclusive lock, this value[BypassLockID] MUST be the same as the value of ExclusiveLockID, as specified in section 2.3.1.1.");

                Site.Assert.AreEqual<string>(
                    "ExclusiveLock",
                    cellSubResponse.SubResponseData.LockType.ToString(),
                    @"[In CellSubResponseDataType] The LockType attribute MUST be set to ""ExclusiveLock"" in the cell subresponse if the ExclusiveLockID attribute is sent in the cell subrequest and the protocol server is successfully able to take an exclusive lock.");
            }

            // Update contents without the ByPassLockID and coalesce true.
            putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), SharedTestSuiteHelper.GenerateRandomFileContent(this.Site));
            putChange.SubRequestData.BypassLockID = null;
            putChange.SubRequestData.CoalesceSpecified = true;
            putChange.SubRequestData.Coalesce = true;

            response = Adapter.CellStorageRequest(randomFileUrl, new SubRequestType[] { putChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    "When the file is locked by exclusive lock and the ByPassLockID is not specified, the server returns the error code success.");
        }
    }
}