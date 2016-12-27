namespace Microsoft.Protocols.TestSuites.SharedTestSuite
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.SharedAdapter;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A class which contains test cases used to capture the requirements related with FileOperation operation.
    /// </summary>
    [TestClass]
    public abstract class S17_FileOperation : SharedTestSuiteBase
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
        public void S17_FileOperationInitialization()
        {
            this.DefaultFileUrl = this.PrepareFile();
        }

        #endregion

        #region Test Cases for "FileOperation" sub-request.

        /// <summary>
        /// A method used to verify that FileOperation sub-request can be executed successfully when all input parameters are correct.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC01_FileOperation_Success()
        {
            // Initialize the context using user01 and defaultFileUrl.
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Get the exclusive lock with all valid parameters, expect the server responses the error code "Success".
            ExclusiveLockSubRequestType subRequest = SharedTestSuiteHelper.CreateExclusiveLockSubRequest(ExclusiveLockRequestTypes.GetLock);
            CellStorageResponse cellStoreageResponse = this.Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { subRequest });
            ExclusiveLockSubResponseType exclusiveResponse = SharedTestSuiteHelper.ExtractSubResponse<ExclusiveLockSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                        ErrorCodeType.Success,
                        SharedTestSuiteHelper.ConvertToErrorCodeType(exclusiveResponse.ErrorCode, this.Site),
                        "Test case cannot continue unless the Get Lock of ExclusiveLock sub request succeeds.");

            // Record the current file status.
            this.StatusManager.RecordExclusiveLock(this.DefaultFileUrl, subRequest.SubRequestData.ExclusiveLockID);

            string fileName = this.DefaultFileUrl.Substring(this.DefaultFileUrl.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);
            string newName = Common.GenerateResourceName(this.Site, "fileName") + ".txt";

            FileOperationSubRequestType fileOperationSubRequest = SharedTestSuiteHelper.CreateFileOperationSubRequest(FileOperationRequestTypes.Rename, newName, SharedTestSuiteHelper.DefaultExclusiveLockID, this.Site);
            
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { fileOperationSubRequest });
            
            FileOperationSubResponseType fileOperationSubResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.IsNotNull(fileOperationSubResponse, "The object 'versioningSubResponse' should not be null.");
            this.Site.Assert.IsNotNull(fileOperationSubResponse.ErrorCode, "The object 'versioningSubResponse.ErrorCode' should not be null.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11146
                Site.CaptureRequirementIfAreEqual<string>(
                    "Success",
                    fileOperationSubResponse.ErrorCode,
                    "MS-FSSHTTP",
                    11109,
                    @"[In FileOperationSubRequestDataType] This parameter [ExclusiveLockID] is used to validate that the file operation can be performed even though the file is under exclusive lock.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11120
                // This requirement can be captured directly after capturing MS-FSSHTTP_R11109
                Site.CaptureRequirement(
                    "MS-FSSHTTP",
                    11120,
                    @"[In FileOperationSubResponseType] In the case of success, it contains information requested as part of a file operation subrequest.");
            }
            else
            {
                Assert.AreEqual<string>(
                    "Success",
                    fileOperationSubResponse.ErrorCode,
                    "MS-FSSHTTP",
                    @"[In FileOperationSubRequestDataType] This parameter [ExclusiveLockID] is used to validate that the file operation can be performed even though the file is under exclusive lock.");
            }
        }

        /// <summary>
        /// A method used to verify that FileOperation sub-request failed with empty url.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC02_FileOperation_EmptyUrl()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            string fileName = this.DefaultFileUrl.Substring(this.DefaultFileUrl.LastIndexOf("/", StringComparison.OrdinalIgnoreCase) + 1);
            string newName = Common.GenerateResourceName(this.Site, "fileName") + ".txt";

            FileOperationSubRequestType fileoperationSubRequest = SharedTestSuiteHelper.CreateFileOperationSubRequest(FileOperationRequestTypes.Rename, newName, null, this.Site);
            
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(string.Empty, new SubRequestType[] { fileoperationSubRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11121
                Site.CaptureRequirementIfAreNotEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.Success,
                    cellStoreageResponse.ResponseVersion.ErrorCode,
                    "MS-FSSHTTP",
                    11121,
                    @"[In FileOperationSubResponseType] In the case of failure, the ErrorCode attribute that is part of a SubResponse element specifies the error code result for this subrequest.");
            }
            else
            {
                Site.Assert.AreNotEqual<GenericErrorCodeTypes>(
                    GenericErrorCodeTypes.Success,
                    cellStoreageResponse.ResponseVersion.ErrorCode,
                    "Error should occur if call fileoperation request with empty url.");
            }
        }

        /// <summary>
        /// A method used to verify that FileOperation sub-request failed with FileOperationRequestType is not specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC03_FileOperation_ErrorCode()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            FileOperationSubRequestType fileOperationSubRequest = new FileOperationSubRequestType();

            fileOperationSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            fileOperationSubRequest.SubRequestData = new FileOperationSubRequestDataType();
            fileOperationSubRequest.SubRequestData.FileOperation = FileOperationRequestTypes.Rename;
            fileOperationSubRequest.SubRequestData.FileOperationRequestTypeSpecified = false;
            fileOperationSubRequest.SubRequestData.ExclusiveLockID = null;

            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { fileOperationSubRequest });

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 11267, this.Site))
                {
                    FileOperationSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
                    ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidArgument,
                        errorCode,
                        "MS-FSSHTTP",
                        11267,
                        @"[In Appendix B: Product Behavior] If the specified attributes[FileOperationRequestType attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the file opeartion subresponse. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
                }

                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 11268, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                             GenericErrorCodeTypes.HighLevelExceptionThrown,
                             cellStoreageResponse.ResponseVersion.ErrorCode,
                             "MS-FSSHTTP",
                             11268,
                             @"[In Appendix B: Product Behavior] The implementation does return a ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the file operation subresponse.<27> Section 2.3.1.33:  In SharePoint Server 2013, if the FileOperationRequestType attributes is not provided, a ""HighLevelExceptionThrown"" error code MUST be returned as part of the SubResponseData element associated with the file operation subresponse.");
                }
            }

            else
            {
                FileOperationSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
                ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidArgument,
                    errorCode,
                    "MS-FSSHTTP",
                    11267,
                    @"[In Appendix B: Product Behavior] If the specified attributes[FileOperationRequestType attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the file opeartion subresponse. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 11268, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<GenericErrorCodeTypes>(
                            GenericErrorCodeTypes.HighLevelExceptionThrown,
                            cellStoreageResponse.ResponseVersion.ErrorCode,
                            "MS-FSSHTTP",
                            11268,
                            @"[In Appendix B: Product Behavior] The implementation does return a ""HighLevelExceptionThrown"" error code as part of the SubResponseData element associated with the file operation subresponse.<27> Section 2.3.1.33:  In SharePoint Server 2013, if the FileOperationRequestType attributes is not provided, a ""HighLevelExceptionThrown"" error code MUST be returned as part of the SubResponseData element associated with the file operation subresponse.");
            }
        }
        #endregion
    }
}