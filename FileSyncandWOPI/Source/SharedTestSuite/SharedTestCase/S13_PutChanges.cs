namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with PutChanges operation.
    /// </summary>
    [TestClass]
    public abstract class S13_PutChanges : SharedTestSuiteBase
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
        public void S13_PutChangesInitialization()
        {
            // Initialize the default file URL.
            this.DefaultFileUrl = this.PrepareFile();
          
        }

        #endregion

        #region Test Cases
        /// <summary>
        /// A method used to test the Put Changes subRequest processing.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC01_PutChanges_ValidExpectedStorageIndexExtendedGUID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            ExGuid storageIndex = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            dataElements.AddRange(fsshttpbResponse.DataElementPackage.DataElements);
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            // Expect the operation succeeds
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R936
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.Success,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTPB",
                         936,
                         @"[In Put Changes] Expected Storage Index Extended GUID (variable): If the expected Storage Index was specified and the key that is to be updated in the protocol server’s StorageIindex exists in the expected storage index, the corresponding values in the protocol server’s storage index and the expected Storage Index MUST match.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.Success,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[In Put Changes] Expected Storage Index Extended GUID (variable): If the expected storage index was specified and the key that is to be updated in the protocol server’s storage index exists in the expected storage index, the corresponding values in the protocol server’s storage index and the expected storage index MUST match.");
            }

            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            
            // Create a putChanges cellSubRequest
            ExGuid storageIndexExGuid2;
            List<DataElement> dataElements1 = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid2);
            putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid2);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            cellRequest.AddSubRequest(putChange,  null);
            CellSubRequestType cellSubRequest1 = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest1 });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R4050
                Site.CaptureRequirementIfAreEqual<CellErrorCode>(
                         CellErrorCode.Referenceddataelementnotfound,
                         fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                         "MS-FSSHTTPB",
                         4050,
                         @"[In Put Changes] If the Extended GUID does not have the corresponding data element in the Data Element Package of the request, the protocol server MUST return a Cell Error failure value of 16 indicating the referenced data element not found failure, as specified in section 2.2");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server processes the PutChanges subRequest successfully when the ExpectedStorageIndexExtendedGUID attribute is specified and Imply Null Expected if No Mapping flag set to one.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC02_PutChanges_InvalidExpectedStorageIndexExtendedGUID()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse queryChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryChangeResponse, this.Site);

            // Put changes to upload the content once to change the server state. In this case, the server expected the storage index value is different with previous returned.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges((ulong)SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)));
            CellStorageResponse putResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType putSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(putResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putSubResponse.ErrorCode, this.Site), "The operation PutChanges should succeed.");
            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(putSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            // Create a putChanges cellSubRequest with specified ExpectedStorageIndexExtendedGUID value as the step 1 returned.
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)), out storageIndexExGuid);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Specify ExpectedStorageIndexExtendedGUID
            putChangeSecond.ExpectedStorageIndexExtendedGUID = queryChangeResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;
            dataElements.AddRange(queryChangeResponse.DataElementPackage.DataElements);

            cellRequestSecond.AddSubRequest(putChangeSecond, dataElements);

            // Put changes to the protocol server 
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Due to the Invalid Expected StorageIndexExtendedGUID, the put changes should fail.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R937
                Site.CaptureRequirementIfAreEqual<CellErrorCode>(
                         CellErrorCode.Coherencyfailure,
                         fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                         "MS-FSSHTTPB",
                         937,
                         @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): otherwise[If the expected Storage Index was specified and the key that is to be updated in the protocol server’s Storage Index exists in the expected Storage Index but the corresponding values in the protocol server’s Storage Index and the expected Storage Index doesn't match] the protocol server MUST return a Cell Error Coherency failure value of 12 indicating a coherency failure as specified in section 2.2.3.2.1.");
            }
            else
            {
                Site.Assert.AreEqual<CellErrorCode>(
                    CellErrorCode.Coherencyfailure,
                    fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                    @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): otherwise[If the expected storage index was specified and the key that is to be updated in the protocol server’s storage index exists in the expected storage index but the corresponding values in the protocol server’s storage index and the expected storage index doesn't match] the protocol server MUST return a Cell Error Coherency failure value of 12 indicating a coherency failure as specified in section 2.2.3.2.1.");
            }
        }

        /// <summary>
        /// A method used to test the protocol server apply the change successfully when the ExpectedStorageIndexExtendedGUID attribute is not specified and Imply Null Expected if No Mapping flag set to zero.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC03_PutChanges_NotSpecifiedExpectedStorageIndexExtendedGUID_ImplyFlagZero()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified Expected Storage Index Extended GUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.Unicode.GetBytes(Common.GenerateResourceName(this.Site, "FileContent")), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Assign the ImplyNullExpectedIfNoMapping equal to 0 to make apply the change without checking the expected storage index value. 
            putChange.ImplyNullExpectedIfNoMapping = 0;
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the PutChanges operation succeeds, then capture MS-FSSHTTPB_R939
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         939,
                         @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): If this flag[Imply Null Expected if No Mapping] is zero, the protocol server MUST apply the change without checking the current value.");
            }
        }

        /// <summary>
        /// A method used to test the protocol server apply the change successfully when a mapping exists and Imply Null Expected if No Mapping flag set to zero.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC04_PutChanges_MappingExist_ImplyFlagZero()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server to get the current server storage index extended guid for the specified file.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse queryChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryChangeResponse, this.Site);

            // Restore the current server storage index extended guid. 
            ExGuid storageIndex = queryChangeResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest specified the Expected Storage Index Extended GUID tribute and Imply Null Expected if No Mapping flag set to one
            // Also use the server returns the data elements, in this case the key will exist.
            FsshttpbCellRequest cellRequestFirst = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndex);

            // Assign the ImplyNullExpectedIfNoMapping to 0.
            putChange.ImplyNullExpectedIfNoMapping = 0;
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            cellRequestFirst.AddSubRequest(putChange, queryChangeResponse.DataElementPackage.DataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestFirst.ToBase64());

            // Put changes to the protocol server to expect the server responds the success.
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The first PutChanges operation should succeed.");
        }

        /// <summary>
        /// A method used to test the protocol server returns a coherency failure when a mapping exists and Imply Null Expected if No Mapping flag set to one.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC05_PutChanges_MappingExist_ImplyFlagOne()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server to get the current server storage index extended guid for the specified file.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse queryChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryChangeResponse, this.Site);

            // Restore the current server storage index extended guid. 
            ExGuid storageIndex = queryChangeResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest specified the Expected Storage Index Extended GUID tribute and Imply Null Expected if No Mapping flag set to one
            // Also use the server returns the data elements, in this case the key will exist.
            FsshttpbCellRequest cellRequestFirst = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndex);

            // Assign the ImplyNullExpectedIfNoMapping to 1.
            putChange.ImplyNullExpectedIfNoMapping = 1;
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            cellRequestFirst.AddSubRequest(putChange, queryChangeResponse.DataElementPackage.DataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestFirst.ToBase64());

            // Put changes to the protocol server to expect the server responds the Coherency failure error.
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");

        }

        /// <summary>
        /// A method used to test the protocol server apply the change failed when the ExpectedStorageIndexExtendedGUID attribute is not specified and Imply Null Expected if No Mapping flag set to zero.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC06_PutChanges_NotSpecifiedExpectedStorageIndexExtendedGUID_ImplyFlagOne()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified Expected Storage Index Extended GUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.Unicode.GetBytes(Common.GenerateResourceName(this.Site, "FileContent")), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Assign the ImplyNullExpectedIfNoMapping equal to 1 to make apply the change with checking the expected storage index value. 
            putChange.ImplyNullExpectedIfNoMapping = 1;
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} failed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            CellError cellError = (CellError)putChangeResponse.CellSubResponses[0].ResponseError.ErrorData;
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTPB_R941
                Site.CaptureRequirementIfAreEqual<CellErrorCode>(
                         CellErrorCode.Coherencyfailure,
                         cellError.ErrorCode,
                         "MS-FSSHTTPB",
                         941,
                         @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): [if Imply Null Expected if No Mapping flag specifies 1,] If a mapping exists, the protocol server MUST return a Cell Error failure value of 12 indicating a coherency failure as specified in section 2.2.3.2.1.");
            }
            else
            {
                Site.Assert.AreEqual<CellErrorCode>(
                        CellErrorCode.Coherencyfailure,
                        cellError.ErrorCode,
                        @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): [if Imply Null Expected if No Mapping flag specifies 1,] If the Expected Storage Index Extended GUID is not specified, the protocol server returns a Cell Error failure value of 12 indicating a coherency failure as specified in section 2.2.3.2.1.");
            }
        }

        /// <summary>
        /// A method used to verify the requirements related with Priority.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC07_PutChanges_Prioriy()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create the first Put changes subRequest with Priority attribute value set to 0.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuidFirst;
            List<DataElement> dataElementsFirst = DataElementUtils.BuildDataElements(System.Text.Encoding.UTF8.GetBytes("First Content"), out storageIndexExGuidFirst);
            PutChangesCellSubRequest putChangeFirst = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuidFirst);
            putChangeFirst.Priority = 0;
            cellRequest.AddSubRequest(putChangeFirst, dataElementsFirst);

            // Create the second Put changes subRequest with Priority attribute value set to 100.
            ExGuid storageIndexExGuidSecond;
            List<DataElement> dataElementsSecond = DataElementUtils.BuildDataElements(System.Text.Encoding.UTF8.GetBytes("Second Content"), out storageIndexExGuidSecond);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuidSecond);
            putChangeSecond.Priority = 100;
            cellRequest.AddSubRequest(putChangeSecond, dataElementsSecond);

            // Send PutChanges subRequest to the protocol server.
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes succeed.");

            // Query changes from the protocol server.
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes succeed.");

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4110, this.Site))
            {
                Site.CaptureRequirement(
                    "MS-FSSHTTPB",
                    4110,
                    @"[In Appendix B: Product Behavior] Implementation does execute Sub-requests with different or same Priority in any order with respect to each other. (<8> Section 2.2.2.1:  SharePoint Server 2010 and SharePoint Server 2013 execute Sub-requests with different or same Priority in any order with respect to each other.)");
            }
        }

        /// <summary>
        /// A method used to test the protocol server returns a coherency failure instead of Referenced Data Element Not Found when the D - Favor Coherency Failure Over Not Found attribute is specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC08_PutChanges_FavorCoherencyFailureOverNotFound()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            // Put changes to upload the content once.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges((ulong)SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)));
            CellStorageResponse putResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType putSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(putResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putSubResponse.ErrorCode, this.Site), "The operation PutChanges should succeed.");
            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(putSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            // Create a putChanges cellSubRequest with nonexistent ExpectedStorageIndexExtendedGUID
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.Unicode.GetBytes(Common.GenerateResourceName(this.Site, this.DefaultFileUrl)), out storageIndexExGuid);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChangeSecond.ExpectedStorageIndexExtendedGUID = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Favor the coherency failure other than not found element.
            putChangeSecond.FavorCoherencyFailureOverNotFound = 1;

            cellRequest.AddSubRequest(putChangeSecond, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // For Microsoft product, in this case the server still responses Referenceddataelementnotfound.
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 51701, SharedContext.Current.Site))
                {
                    Site.CaptureRequirementIfAreEqual<string>(
                        "referenceddataelementnotfound".ToLower(CultureInfo.CurrentCulture),
                        fsshttpbResponse.CellSubResponses[0].ResponseError.ErrorData.ErrorDetail.ToLower(CultureInfo.CurrentCulture),
                        "MS-FSSHTTPB",
                        51701,
                        @"[In Put Changes] Implementation does return a Referenced Data Element Not Found failure, when D - Favor Coherency Failure Over Not Found is set to 1 and a Referenced Data Element Not Found (section 2.2.3.2.1) failure occurred. (SharePoint 2010 and above follow this behavior.)");
                }
            }
            else
            {
                this.Site.Assert.AreEqual<string>(
                    "Referenceddataelementnotfound".ToLower(CultureInfo.CurrentCulture),
                    fsshttpbResponse.CellSubResponses[0].ResponseError.ErrorData.ErrorDetail.ToLower(CultureInfo.CurrentCulture),
                    "For Microsoft, the server still responses Referenceddataelementnotfound even the FavorCoherencyFailureOverNotFound is set to 1.");
            }
        }

        /// <summary>
        /// A method used to test the protocol server returns rough same put changes response when the DataPackage's reserved value is different in the request.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC09_PutChanges_DataPackageReservedIgnore()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);

            // Make the DataElementPackage reserved value to 0.
            cellRequest.DataElementPackage.Reserved = 0;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuidSecond;
            List<DataElement> dataElementsSecond = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuidSecond);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuidSecond);
            cellRequestSecond.AddSubRequest(putChangeSecond, dataElementsSecond);

            // Make the DataElementPackage reserved value to 1.
            cellRequestSecond.DataElementPackage.Reserved = 1;
            CellSubRequestType cellSubRequestSecond = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            CellStorageResponse responseSecond = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequestSecond });
            CellSubResponseType cellSubResponseSecond = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(responseSecond, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponseSecond.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponseSecond, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponseSecond, this.Site);

            // Compare this two responses roughly, and if these two response are identical in main part, then capture the requirement MS-FSSHTTPB_R371
            bool isVerifyR371 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(putChangeResponse, putChangeResponseSecond, this.Site);

            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifyR371);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R371
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR371,
                         "MS-FSSHTTPB",
                         371,
                         @"[In Data Element Package] Whenever the Reserved field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isVerifyR371,
                    @"[In Data Element Package] Whenever the Reserved field is set to 0 or 1, the protocol server must return the same response.");
            }
        }

        /// <summary>
        /// A method used to test the protocol server can accept partial creating file contents.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC10_PutChanges_Partial()
        {
            string fileUrl = Common.GetConfigurationPropertyValue("BigFile", this.Site);

            // Initialize the service
            this.InitializeContext(fileUrl, this.UserName01, this.Password01, this.Domain);

            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                        @"Test case cannot continue unless the query change operation succeeds.");

            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            // Get the file contents
            byte[] bytes = new IntermediateNodeObject.RootNodeObjectBuilder()
                                            .Build(fsshttpbResponse.DataElementPackage.DataElements, fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID)
                                            .GetContent()
                                            .ToArray();

            // Update file contents
            byte[] newBytes = SharedTestSuiteHelper.GenerateRandomFileContent(Site);
            System.Array.Copy(newBytes, 0, bytes, 0, newBytes.Length);

            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(bytes, out storageIndexExGuid);

            // Send the first partial put changes sub request null storageIndexExGuid.
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), null);

            cellRequest.AddSubRequest(putChange, dataElements.Take(4).ToList());
            putChange.Partial = 1;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            // Expect the operation succeeds
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);

            // Send the last partial put changes sub request
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.PartialLast = 1;
            cellRequest.AddSubRequest(putChange, dataElements.Skip(4).ToList());

            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            cellSubRequest.SubRequestData.CoalesceSpecified = true;
            cellSubRequest.SubRequestData.Coalesce = true;
            response = Adapter.CellStorageRequest(fileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
        }

        /// <summary>
        /// A method used to verify whether A–Reserved (1 bit) is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC11_PutChanges_AReserved()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest A-Reserved (1 bit) set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);
            cellRequest.Reserve1 = 0;
            cellRequest.IsRequestHashingOptionsUsed = true;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the Fsshttpb response
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create a putChanges cellSubRequest A-Reserved (1 bit) set to 1.
            cellRequest.Reserve1 = 1;
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R990251
            bool isVerifiedR990251 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR990251);
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR990251,
                         "MS-FSSHTTPB",
                         990251,
                         @"[In Request Message Syntax] Whenever the[A – Reserved (1 bit, optional) field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(isVerifiedR990251, @"[In Request Message Syntax] Whenever the[A – Reserved (1 bit) field is set to 0 or 1, the protocol server must return the same response.");
            }
        }

        /// <summary>
        /// A method used to verify whether B–Reserved (1 bit) is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC12_PutChanges_BReserved()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest B-Reserved (1 bit) set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);
            cellRequest.Reserve2 = 0;
            cellRequest.IsRequestHashingOptionsUsed = true;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create a putChanges cellSubRequest B-Reserved (1 bit) set to 1.
            cellRequest.Reserve2 = 1;
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R990271
            bool isVerifiedR990271 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR990271);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR990271,
                         "MS-FSSHTTPB",
                         990271,
                         @"[In Request Message Syntax] Whenever the[B – Reserved (1 bit, optional) field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(isVerifiedR990271, @"[In Request Message Syntax] Whenever the[B – Reserved (1 bit) field is set to 0 or 1, the protocol server must return the same response.");
            }
        }

        /// <summary>
        /// A method used to verify whether E–Reserved (1 bit) is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC13_PutChanges_EReserved()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest E-Reserved (1 bit) set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);
            cellRequest.Reserve3 = 0;
            cellRequest.IsRequestHashingOptionsUsed = true;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create a putChanges cellSubRequest E-Reserved (1 bit) set to 1.
            cellRequest.Reserve3 = 1;
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R990271
            bool isVerifiedR990341 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR990341);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR990341,
                         "MS-FSSHTTPB",
                         990341,
                         @"[In Request Message Syntax] Whenever the[E – Reserved (4 bit, optional) field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(isVerifiedR990341, @"[In Request Message Syntax] Whenever the[E – Reserved (4 bit) field is set to 0 or 1, the protocol server must return the same response.");
            }
        }

        /// <summary>
        /// A method used to verify whether E-Abort Remaining Put Changes on Failure (1 bit) is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC14_PutChanges_EAbort()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest E-Abort Remaining Put Changes On Failure (1 bit) set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.AbortRemainingPutChangesOnFailure = 0;
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create a putChanges cellSubRequest E-Abort Remaining Put Changes On Failure (1 bit) set to 1.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.AbortRemainingPutChangesOnFailure = 1;
            cellRequest.AddSubRequest(putChange, dataElements);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R990271
            bool isVerifiedR51801 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR51801);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR51801,
                         "MS-FSSHTTPB",
                         51801,
                         @"[In Put Changes] Whenever the E - Abort Remaining Put Changes on Failure field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(isVerifiedR51801, @"[In Put Changes] Whenever the E - Abort Remaining Put Changes on Failure field is set to 0 or 1, the protocol server must return the same response.");
            }
        }

        /// <summary>
        /// A method used to verify whether Reserved (11 bits) Flag is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC15_PutChanges_Reserved()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest Reserved set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);

            PutChangesCellSubRequest putChangeFirst = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChangeFirst.IsAdditionalFlagsUsed = true;
            putChangeFirst.Reserve = 0;
            cellRequest.AddSubRequest(putChangeFirst, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the first fsshttpb subResponse            
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create another putChanges subRequest with Reserved set to 1.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChangeSecond.IsAdditionalFlagsUsed = true;
            putChangeSecond.Reserve = 1;
            cellRequest.AddSubRequest(putChangeSecond, dataElements);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });

            // Extract the second fsshttpb subResponse            
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R990271
            bool isVerifiedR99054 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR99054);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR99054,
                         "MS-FSSHTTPB",
                         99054,
                         @"[Additional Flags] The server response is the same whether the Reserved (11 bits) Flag is set to 0 or 1.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isVerifiedR99054,
                    "[Additional Flags] The server response is the same whether the Reserved (11 bits) Flag is set to 0 or 1.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server does not return the Applied Storage Index Id when the ReturnAppliedStorageIndexIdEntries set to 0.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC16_PutChanges_ReturnAppliedStorageIndexIdEntries_Zero()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Make the ReturnAppliedStorageIndexIdEntries  value to 0.
            putChange.IsAdditionalFlagsUsed = true;
            putChange.ReturnAppliedStorageIndexIdEntries = 0;
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
               ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            bool notInlcudeStorageIndex = putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse != null
                                                 && new ExGuid().Equals(putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.AppliedStorageIndexID);

            Site.Log.Add(
                TestTools.LogEntryKind.Debug,
                "When the ReturnAppliedStorageIndexIdEntries flag is not set, the server will not return AppliedStorageIndexID, actually it {0} return",
                notInlcudeStorageIndex ? "does not" : "does");

            Site.Assert.IsTrue(
                notInlcudeStorageIndex,
                "When the ReturnAppliedStorageIndexIdEntries flag is not set, the server will not return AppliedStorageIndexID");
        }

        /// <summary>
        /// A method used to verify the protocol server return the Applied Storage Index Id successfully when the ReturnAppliedStorageIndexIdEntries set to 1.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC17_PutChanges_ReturnAppliedStorageIndexIdEntries_One()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Make the ReturnAppliedStorageIndexIdEntries  value to 1.
            putChange.IsAdditionalFlagsUsed = true;
            putChange.ReturnAppliedStorageIndexIdEntries = 1;
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
               ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            bool isVerifyR99044 = putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse != null
                          && !new ExGuid().Equals(putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.AppliedStorageIndexID);

            Site.Log.Add(
                TestTools.LogEntryKind.Debug,
                "When the ReturnAppliedStorageIndexIdEntries flag is set, the server will return AppliedStorageIndexID, actually it {0} return",
                isVerifyR99044 ? "does" : "does not");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99044
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR99044,
                         "MS-FSSHTTPB",
                         99044,
                         @"[Additional Flags] A – Return Applied Storage Index Id Entries (1 bit): A bit that specifies that the Storage Indexes that are applied to the storage as part of the Put Changes will be returned in a Storage Index specified in the Put Changes response by the Applied Storage Index Id (section 2.2.3.1.3).");

                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99108
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         99108,
                         @"[Appendix B: Product Behavior] Additional Flags is supported by SharePoint Server 2013 and SharePoint Server 2016.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isVerifyR99044,
                    @"[Additional Flags] A – Return Applied Storage Index Id Entries (1 bit): A bit that specifies that the Storage Indexes that are applied to the storage as part of the Put Changes will be returned in a Storage Index specified in the Put Changes response by the Applied Storage Index Id (section 2.2.3.1.3).");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server return in a Data Element Collection successfully when the  Return Data Elements Added set to 1.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC18_PutChanges_ReturnDataElementsAddedFlag_One()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Make the ReturnDataElementsAdded value to 1.
            putChange.IsAdditionalFlagsUsed = true;
            putChange.ReturnDataElementsAdded = 1;
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            bool isVerifyR99045002 = putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse != null
                    && putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.DataElementAdded != null
                    && putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.DataElementAdded.Count.DecodedValue != 0;

            Site.Log.Add(
                TestTools.LogEntryKind.Debug,
                "When the ReturnDataElementsAdded flag is set, the server will return the added data elements, actually it {0} return",
                isVerifyR99045002 ? "does" : "does not");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99045002
                Site.CaptureRequirementIfIsTrue(
                         isVerifyR99045002,
                         "MS-FSSHTTPB",
                         99045002,
                         @"[Additional Flags] B – Return Data Elements Added (1 bit): When the ReturnDataElementsAdded flag is set, the server will return the added data elements.");
            }
            else
            {
                Site.Assert.IsTrue(
                    isVerifyR99045002,
                    @"[Additional Flags] B – Return Data Elements Added (1 bit): When the ReturnDataElementsAdded flag is set, the server will return the added data elements.");
            }
        }

        /// <summary>
        /// A method used to verify the protocol server does not return in a Data Element Collection when the  Return Data Elements Added set to 0.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC19_PutChanges_ReturnDataElementsAddedFlag_Zero()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest without specified ExpectedStorageIndexExtendedGUID attribute and Imply Null Expected if No Mapping flag set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Make the ReturnDataElementsAdded value to 0.
            putChange.IsAdditionalFlagsUsed = true;
            putChange.ReturnDataElementsAdded = 0;
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);
            bool notIncludeAddedDataElements = putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse != null
                                                    && putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.DataElementAdded != null
                                                    && putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().PutChangesResponse.DataElementAdded.Count.DecodedValue == 0;

            Site.Log.Add(
                TestTools.LogEntryKind.Debug,
                "When the ReturnDataElementsAdded flag is set, the server will not return the added data elements, actually it {0} return",
                notIncludeAddedDataElements ? "does not" : "does");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99045001, this.Site))
                {
                    Site.CaptureRequirementIfIsTrue(
                        notIncludeAddedDataElements,
                        99045001,
                        @"[Additional Flags] B – Return Data Elements Added (1 bit): When the ReturnDataElementsAdded flag is not set, the server will not return the added data elements.");
                }
            }
            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99045001, this.Site))
                {
                    Site.Assert.IsTrue(
                        notIncludeAddedDataElements,
                        @"When the ReturnDataElementsAdded flag is not set, the server will not return the added data elements");
                }
            }
        }

        /// <summary>
        /// A method used to verify server check the index entry that is actually applied when Coherency Check Only Applied Index Entries set to 1.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC20_PutChanges_CheckForIdReuse()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            var preGroupData = fsshttpbResponse.DataElementPackage.DataElements.First(e => e.DataElementType == DataElementType.ObjectGroupDataElementData);

            // Change the new object group data element's element id as the previous step returned id, this will make the data element id for reusing.
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            var groupData = dataElements.First(e => e.DataElementType == DataElementType.ObjectGroupDataElementData);
            groupData.DataElementExtendedGUID = preGroupData.DataElementExtendedGUID;

            // Send the upload put changes request with CheckForIdReuse value as 1.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            cellRequest.AddSubRequest(putChange, dataElements);
            putChange.IsAdditionalFlagsUsed = true;
            putChange.CheckForIdReuse = 1;
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            Site.Assert.AreEqual<CellErrorCode>(
                    CellErrorCode.ExtendedGuidCollision,
                    fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                    "When the data element id is reused, the server will respond the error code ExtendedGuidCollision");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R99049
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTPB",
                         99049,
                         @"[Additional Flags] When D – Coherency Check Only Applied Index Entries (1 bit) is set, the server check the index entry that is actually applied and an index entry that is not applied is not checked.");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTPB",
                         99046,
                         @"[Additional Flags] C – Check for Id Reuse (1 bit): A bit that specifies that the server will attempt to check the Put Changes request for the re-use of previously used IDs. ");

                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.CellRequestFail,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTPB",
                         99047,
                         @"[Additional Flags] This [check the Put Changes Request for the re-use of previously used Ids] might occur when ID allocations are used and a client rollback occurs.");
            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.CellRequestFail,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                    @"[Additional Flags] When D - Coherency Check Only Applied Index Entries (1 bit) is set, the server check the index entry that is actually applied and an index entry that is not applied is not checked.");
            }
        }

        /// <summary>
        /// A method used to verify server will not bypass the necessary checks when the FullFileReplacePut flag is true.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC21_PutChanges_FullFileReplacePut()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse queryChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryChangeResponse, this.Site);

            // Put changes to upload the content once to change the server state. In this case, the server expected the storage index value is different with previous returned.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges((ulong)SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)));
            CellStorageResponse putResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType putSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(putResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putSubResponse.ErrorCode, this.Site), "The operation PutChanges should succeed.");
            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(putSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            // Create a putChanges cellSubRequest with specified ExpectedStorageIndexExtendedGUID value as the step 1 returned and with the full file replace put flag as 1.
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)), out storageIndexExGuid);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Specify ExpectedStorageIndexExtendedGUID and FullFileReplacePut flag.
            putChangeSecond.ExpectedStorageIndexExtendedGUID = queryChangeResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;
            putChangeSecond.IsAdditionalFlagsUsed = true;
            putChangeSecond.FullFileReplacePut = 1;
            dataElements.AddRange(queryChangeResponse.DataElementPackage.DataElements);
            cellRequestSecond.AddSubRequest(putChangeSecond, dataElements);

            // Put changes to the protocol server 
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Due to the Invalid Expected StorageIndexExtendedGUID, the put changes should fail.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 9905101, SharedContext.Current.Site))
                {
                    Site.CaptureRequirementIfAreEqual<CellErrorCode>(
                             CellErrorCode.Coherencyfailure,
                             fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                             "MS-FSSHTTPB",
                             9905101,
                             @"[In Appendix B: Product Behavior] Implementation does not bypass a full file save and related checks that would otherwise be unnecessary, when his flag [E – Full File Replace Put (1 bit)] is set. (Microsoft SharePoint Server 2010 and Microsoft SharePoint Workspace 2010 and above follow this behavior.)");
                }
            }
            else
            {
                Site.Assert.AreEqual<CellErrorCode>(
                         CellErrorCode.Coherencyfailure,
                         fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                         @"[In Appendix B: Product Behavior] Implementation does not bypass a full file save and related checks that would otherwise be unnecessary, when his flag [E – Full File Replace Put (1 bit)] is set. (Microsoft SharePoint Server 2010 and Microsoft SharePoint Workspace 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// A method used to verify server will not bypass the necessary checks when the ForceRevisionChainOptimization flag is false.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC22_PutChanges_ForceRevisionChainOptimization_Zero()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 4130, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Diagnostic Request Option Output field.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest with specified DiagnosticRequestOptionInput attribute and ForceRevisionChainOptimization set to zero
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.IsDiagnosticRequestOptionInputUsed = true;
            putChange.ForceRevisionChainOptimization = 0;
            cellRequest.AddSubRequest(putChange, dataElements);

            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Test case cannot continue unless the put changes operation on the file {0} succeed.",
                this.DefaultFileUrl);

            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            bool isForced = putChangeResponse.CellSubResponses[0].GetSubResponseData<PutChangesSubResponseData>().DiagnosticRequestOptionOutput.IsDiagnosticRequestOptionOutput;
            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R409502
                Site.CaptureRequirementIfIsFalse(
                         isForced,
                         "MS-FSSHTTPB",
                         409502,
                         @"[In Put Changes] F – Forced (1 bit): [False] specifies whether a forced Revision Chain optimization [does not] occurred.");
            }
            else
            {
                Site.Assert.IsFalse(
                    isForced,
                    @"[In Put Changes] F – Forced (1 bit): [False] specifies whether a forced Revision Chain optimization [does not] occurred.");
            }
        }

        /// <summary>
        /// A method used to verify if the key that is to be updated in the protocol server's Storage Index does not exist in the 
        /// expected Storage Index, the Imply Null Expected if No Mapping flag MUST be evaluated.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC23_PutChanges_ImplyFlagWithKeyNotExist()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            ExGuid storageIndex = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest with specify expected Storage Index and ImplyNullExpectedIfNoMapping set to 0.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            putChange.ImplyNullExpectedIfNoMapping = 0;
            dataElements.AddRange(fsshttpbResponse.DataElementPackage.DataElements);
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            // If the key that is to be updated in the protocol server’s Storage Index does not exist in the expected Storage Index, the Imply Null Expected if No Mapping flag MUST be evaluated.
            // If this flag is zero, the protocol server MUST apply the change without checking the current value.
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");

            queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(fsshttpbResponse, this.Site);
            ExGuid index2 = storageIndex;
            storageIndex = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest with specify expected Storage Index and ImplyNullExpectedIfNoMapping set to 1.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            putChange.ImplyNullExpectedIfNoMapping = 1;
            dataElements.AddRange(fsshttpbResponse.DataElementPackage.DataElements);
            cellRequest.AddSubRequest(putChange, dataElements);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            // If the key that is to be updated in the protocol server’s Storage Index does not exist in the expected Storage Index, the Imply Null Expected if No Mapping flag MUST be evaluated.
            // If this flag is one, the protocol server MUST only apply the change if no mapping exists.
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTPB requirement: MS-FSSHTTPB_R2168
                // This requirement can be captured directly after above steps.
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         2168,
                         @"[In Put Changes] Expected Storage Index Extended GUID (variable): If the key that is to be updated in the protocol server’s Storage Index does not exist in the expected Storage Index, the Imply Null Expected if No Mapping flag MUST be evaluated.");

                // If the PutChanges operation succeeds, then capture MS-FSSHTTPB_R940
                Site.CaptureRequirement(
                         "MS-FSSHTTPB",
                         940,
                         @"[In Put Changes structure] Expected Storage Index Extended GUID (variable): otherwise[If Imply Null Expected if No Mapping is not zero], if the flag[Imply Null Expected if No Mapping] specifies one, the protocol server MUST only apply the change if no mapping exists (the key that is to be updated in the protocol server’s Storage Index doesn't exist or it maps to nil).");
            }
        }

        /// <summary>
        /// A method used to verify the Put Changes request will fail coherency if any of the supplied Storage Indexes are unrooted.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC24_PutChanges_RequireStorageMappingsRooted()
        {
            if (!Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 99108, this.Site))
            {
                Site.Assume.Inconclusive("Implementation does not support the Additional Flags.");
            }

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse queryChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(queryChangeResponse, this.Site);

            // Put changes to upload the content once to change the server state. In this case, the server expected the storage index value is different with previous returned.
            CellSubRequestType putChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedPutChanges((ulong)SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)));
            CellStorageResponse putResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChange });
            CellSubResponseType putSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(putResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(putSubResponse.ErrorCode, this.Site), "The operation PutChanges should succeed.");
            FsshttpbResponse putChangeResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(putSubResponse, this.Site);
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(putChangeResponse, this.Site);

            // Create a putChanges cellSubRequest with specified ExpectedStorageIndexExtendedGUID value as the step 1 returned and with the  Require Storage Mappings Rooted flag as 1.
            FsshttpbCellRequest cellRequestSecond = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.UTF8.GetBytes(SharedTestSuiteHelper.GenerateRandomString(5)), out storageIndexExGuid);
            PutChangesCellSubRequest putChangeSecond = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);

            // Specify ExpectedStorageIndexExtendedGUID and RequireStorageMappingsRooted flag.
            putChangeSecond.ExpectedStorageIndexExtendedGUID = queryChangeResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;
            putChangeSecond.IsAdditionalFlagsUsed = true;
            putChangeSecond.RequireStorageMappingsRooted = 1;
            dataElements.AddRange(queryChangeResponse.DataElementPackage.DataElements);
            cellRequestSecond.AddSubRequest(putChangeSecond, dataElements);

            // Put changes to the protocol server 
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequestSecond.ToBase64());
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreNotEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "Due to the Invalid Expected StorageIndexExtendedGUID, the put changes should fail.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfAreEqual<CellErrorCode>(
                            CellErrorCode.Coherencyfailure,
                            fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                            "MS-FSSHTTPB",
                        4054,
                        @"[Additional Flags] F – Require Storage Mappings Rooted (1 bit): A bit that specifies that the Put Changes request will fail coherency if any of the supplied Storage Indexes are unrooted.");
            }
            else
            {
                Site.Assert.AreEqual<CellErrorCode>(
                         CellErrorCode.Coherencyfailure,
                         fsshttpbResponse.CellSubResponses[0].ResponseError.GetErrorData<CellError>().ErrorCode,
                            @"[Additional Flags] F – Require Storage Mappings Rooted (1 bit): A bit that specifies that the Put Changes request will fail coherency if any of the supplied Storage Indexes are unrooted.");
            }
        }

        /// <summary>
        /// A method used to verify whether E-Abort Remaining Put Changes on Failure (1 bit) is set to zero or 1, the protocol server returns the same response.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S13_TC25_PutChanges_Reserve1Byte()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Create a putChanges cellSubRequest Reserved field is set to zero.
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(SharedTestSuiteHelper.GenerateRandomFileContent(this.Site), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.Reserve1Byte = 0;
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            CellStorageResponse response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The first PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseFirst = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // Create a putChanges cellSubRequest Reserved field is set to 1.
            cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.Reserve1Byte = 1;
            cellRequest.AddSubRequest(putChange, dataElements);
            cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());
            response = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site),
                "The second PutChanges operation should succeed.");

            // Extract the fsshttpb response
            FsshttpbResponse fsshttpbResponseSecond = SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site);

            // If the main part of the two subResponse is same, then capture MS-FSSHTTPB requirement: MS-FSSHTTPB_R217102
            bool isVerifiedR217102 = SharedTestSuiteHelper.CompareSucceedFsshttpbPutChangesResponse(fsshttpbResponseFirst, fsshttpbResponseSecond, this.Site);
            this.Site.Log.Add(TestTools.LogEntryKind.Debug, "Expect the two put changes responses are same, actual {0}", isVerifiedR217102);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                Site.CaptureRequirementIfIsTrue(
                         isVerifiedR217102,
                         "MS-FSSHTTPB",
                         217102,
                         @"[In Put Changes] Whenever the Reserved field is set to 0 or 1, the protocol server must return the same response.");
            }
            else
            {
                Site.Assert.IsTrue(isVerifiedR217102, @"[In Put Changes] Whenever the Reserved field is set to 0 or 1, the protocol server must return the same response.");
            }
        }
        #endregion 
    }
}