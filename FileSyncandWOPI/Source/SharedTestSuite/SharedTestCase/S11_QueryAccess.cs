namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with QueryAccess operation.
    /// </summary>
    [TestClass]
    public abstract class S11_QueryAccess : SharedTestSuiteBase
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

        #region Test Case Initialization

        /// <summary>
        /// A method used to initialize the test class.
        /// </summary>
        [TestInitialize]
        public void S11_QueryAccessInitialization()
        {
            // Initialize the default file URL, for this scenario, the target file URL should not need unique for each test case, just using the preparing one.
            this.DefaultFileUrl = Common.GetConfigurationPropertyValue("NormalFile", this.Site);
           
        }

        #endregion

        #region Test Case
        /// <summary>
        /// A method used to verify QueryAccess when the user has read/write permission.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S11_TC01_QueryAccessReadWrite()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site), "The operation QueryChanges should succeed.");
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            ExGuid storageIndex = fsshttpbResponse.CellSubResponses[0].GetSubResponseData<QueryChangesSubResponseData>().StorageIndexExtendedGUID;

            // Create a putChanges cellSubRequest
            FsshttpbCellRequest cellRequest = SharedTestSuiteHelper.CreateFsshttpbCellRequest();
            ExGuid storageIndexExGuid;
            List<DataElement> dataElements = DataElementUtils.BuildDataElements(System.Text.Encoding.Unicode.GetBytes("bad"), out storageIndexExGuid);
            PutChangesCellSubRequest putChange = new PutChangesCellSubRequest(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID(), storageIndexExGuid);
            putChange.ExpectedStorageIndexExtendedGUID = storageIndex;
            dataElements.AddRange(fsshttpbResponse.DataElementPackage.DataElements);
            cellRequest.AddSubRequest(putChange, dataElements);
            CellSubRequestType putChangesSubRequest = SharedTestSuiteHelper.CreateCellSubRequest(SequenceNumberGenerator.GetCurrentToken(), cellRequest.ToBase64());

            // Put changes to the protocol server
            CellStorageResponse response = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { putChangesSubRequest });
            CellSubResponseType cellSubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(response, 0, 0, this.Site);

            // Expect the Put changes operation succeeds
            this.Site.Assert.AreEqual(ErrorCodeType.Success, SharedTestSuiteHelper.ConvertToErrorCodeType(cellSubResponse.ErrorCode, this.Site), "The PutChanges operation should succeed.");
            SharedTestSuiteHelper.ExpectMsfsshttpbSubResponseSucceed(SharedTestSuiteHelper.ExtractFsshttpbResponse(cellSubResponse, this.Site), this.Site);

            // Call QueryAccess.
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryAccess(0);
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            FsshttpbResponse queryAccessResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            // Get readAccessResponse and writeAccessResponse Data.
            HRESULTError dataRead = queryAccessResponse.CellSubResponses[0]
                                                 .GetSubResponseData<QueryAccessSubResponseData>()
                                                 .ReadAccessResponse.ReadResponseError
                                                 .GetErrorData<HRESULTError>();
            this.Site.Assert.AreEqual<int>(
                                    0,
                                    dataRead.ErrorCode,
                                    "Test case cannot continue unless the read HRESULTError code equals 0 when the user have read/write permission.");

            HRESULTError dataWrite = queryAccessResponse.CellSubResponses[0]
                                                 .GetSubResponseData<QueryAccessSubResponseData>()
                                                 .WriteAccessResponse.WriteResponseError
                                                 .GetErrorData<HRESULTError>();

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error type is HRESULTError, then capture R946.
                Site.CaptureRequirementIfIsNotNull(
                         dataWrite,
                         "MS-FSSHTTPB",
                         946,
                         @"[In Query Access] Response Error (variable): If the Put Changes operation will succeed, the response error will have an error type of HRESULT error.");

                // If the error type is HRESULTError and error code equals 0, then capture R2229.
                Site.CaptureRequirementIfAreEqual<int>(
                         0,
                         dataWrite.ErrorCode,
                         "MS-FSSHTTPB",
                         2229,
                         @"[In Query Access] Response Error (variable): [If the Put Changes operation will succeed, ]the HRESULT error code will be zero.");

                // If error code equals 0, then capture R2231.
                Site.CaptureRequirementIfAreEqual<int>(
                         0,
                         dataWrite.ErrorCode,
                         "MS-FSSHTTPB",
                         2231,
                         @"[In HRESULT Error] Error Code (4 bytes): Zero means that no error occurred.");
            }
            else
            {
                Site.Assert.IsNotNull(
                    dataWrite,
                    @"[In Query Access] Response Error (variable): If the Put Changes operation will succeed, the response error will have an error type of HRESULT error.");

                Site.Assert.AreEqual<int>(
                    0,
                    dataWrite.ErrorCode,
                    @"[In Query Access] Response Error (variable): [If the Put Changes operation will succeed, ]the HRESULT error code will be zero.");
            }
        }

        /// <summary>
        /// A method used to verify QueryAccess when the user has read permission.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S11_TC02_QueryAccessRead()
        {
            string readOnlyUser = Common.GetConfigurationPropertyValue("ReadOnlyUser", this.Site);
            string readOnlyUserPassword = Common.GetConfigurationPropertyValue("ReadOnlyUserPwd", this.Site);

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, readOnlyUser, readOnlyUserPassword, this.Domain);

            // Query changes from the protocol server
            CellSubRequestType queryChange = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryChanges(SequenceNumberGenerator.GetCurrentFSSHTTPBSubRequestID());
            CellStorageResponse queryResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { queryChange });
            CellSubResponseType querySubResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(queryResponse, 0, 0, this.Site);
            FsshttpbResponse fsshttpbResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(querySubResponse, this.Site);
            this.Site.Assert.AreEqual(
                        ErrorCodeType.Success, 
                        SharedTestSuiteHelper.ConvertToErrorCodeType(querySubResponse.ErrorCode, this.Site),
                        "The operation QueryChanges should succeed when the user {0} only has ViewOnly permission.",
                        readOnlyUser);
            
            this.Site.Assert.IsFalse(
                            fsshttpbResponse.CellSubResponses[0].Status, 
                            "The operation QueryChanges should succeed when the user {0} only has ViewOnly permission.",
                            readOnlyUser);

            // Call QueryAccess 
            CellSubRequestType cellSubRequest = SharedTestSuiteHelper.CreateCellSubRequestEmbeddedQueryAccess(0);
            CellStorageResponse cellStorageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { cellSubRequest });
            CellSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<CellSubResponseType>(cellStorageResponse, 0, 0, this.Site);
            FsshttpbResponse queryAccessResponse = SharedTestSuiteHelper.ExtractFsshttpbResponse(subResponse, this.Site);

            // Get readAccessResponse and writeAccessResponse Data.
            HRESULTError dataRead = queryAccessResponse.CellSubResponses[0]
                                                 .GetSubResponseData<QueryAccessSubResponseData>()
                                                 .ReadAccessResponse.ReadResponseError
                                                 .GetErrorData<HRESULTError>();

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // If the error type is HRESULTError, then capture R944.
                Site.CaptureRequirementIfIsNotNull(
                         dataRead,
                         "MS-FSSHTTPB",
                         944,
                         @"[In Query Access] Response Error (variable): If read operations will succeed, the Response Error will have an error type of HRESULT error.");

                // If HResult is 0,MS-FSSHTTPB_R2228 should be covered.
                Site.CaptureRequirementIfAreEqual<int>(
                         0,
                         dataRead.ErrorCode,
                         "MS-FSSHTTPB",
                         2228,
                         @"[In Query Access] Response Error (variable): [If read operations will succeed, ]the HRESULT error code will be zero.");
            }
            else
            {
                Site.Assert.IsNotNull(
                    dataRead,
                    @"[In Query Access] Response Error (variable): If read operations will succeed, the response error will have an error type of HRESULT error.");

                Site.Assert.AreEqual<int>(
                    0,
                    dataRead.ErrorCode,
                    @"[In Query Access] Response Error (variable): [If read operations will succeed, ]the HRESULT error code will be zero.");
            }

            HRESULTError dataWrite = queryAccessResponse.CellSubResponses[0]
                                                 .GetSubResponseData<QueryAccessSubResponseData>()
                                                 .WriteAccessResponse.WriteResponseError
                                                 .GetErrorData<HRESULTError>();

            this.Site.Assert.AreNotEqual<int>(
                                   0,
                                   dataWrite.ErrorCode,
                                   "Test case cannot continue unless the write HRESULTError code not equals 0 when the user have read permission.");
        }
        #endregion
    }
}