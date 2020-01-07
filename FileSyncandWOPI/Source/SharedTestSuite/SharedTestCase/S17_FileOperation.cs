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
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 111111, this.Site), "This test case only runs when FileOperation subrequest is supported.");
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
        /// A method used to verify that FileOperation sub-request failed with FileOperation is not specified.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC02_FileOperation_ErrorCode()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            FileOperationSubRequestType fileOperationSubRequest = new FileOperationSubRequestType();

            fileOperationSubRequest.SubRequestToken = SequenceNumberGenerator.GetCurrentToken().ToString();
            fileOperationSubRequest.SubRequestData = new FileOperationSubRequestDataType();
            fileOperationSubRequest.SubRequestData.FileOperation = FileOperationRequestTypes.Rename;
            fileOperationSubRequest.SubRequestData.FileOperationSpecified = false;
            fileOperationSubRequest.SubRequestData.ExclusiveLockID = null;

            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { fileOperationSubRequest });

            FileOperationSubResponseType subResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            ErrorCodeType errorCode = SharedTestSuiteHelper.ConvertToErrorCodeType(subResponse.ErrorCode, this.Site);

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 11267, this.Site))
                {

                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                    Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                        ErrorCodeType.InvalidArgument,
                        errorCode,
                        "MS-FSSHTTP",
                        11267,
                        @"[In Appendix B: Product Behavior] If the specified attributes[FileOperation attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the file opeartion subresponse. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }

            else
            {
                if (Common.IsRequirementEnabled("MS-FSSHTTP-FSSHTTPB", 11267, this.Site))
                {
                    // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11267
                    Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.InvalidArgument,
                    errorCode,
                    @"[In Appendix B: Product Behavior] If the specified attributes[FileOperation attribute] are not provided, the implementation does return an ""InvalidArgument"" error code as part of the SubResponseData element associated with the file opeartion subresponse. (Microsoft Office 2010 suites/Microsoft SharePoint Foundation 2010/Microsoft SharePoint Server 2010/Microsoft SharePoint Workspace 2010/Microsoft Office 2016/Microsoft SharePoint Server 2016/Microsoft Office 2019/Microsoft SharePoint Server 2019 follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// A method used to verify that ResourceID is not changed even if the URL of the file changes.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC03_ResourceIDNotChanged()
        {
            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetVersions"sub-request with correct input parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { getVersionsSubRequest },
                "1", 2, 2, null, null, null, null, null, null, true);
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                "GetVersions should be succeed.");

            VersionType version = cellStoreageResponse.ResponseVersion as VersionType;
            Site.Assume.AreEqual<ushort>(3, version.MinorVersion, "This test case runs only when MinorVersion is 3 which indicates the protocol server is capable of performing ResourceID specific behavior.");

            string resourceID = cellStoreageResponse.ResponseCollection.Response[0].ResourceID;

            // Rename the file
            string newName = Common.GenerateResourceName(this.Site, "fileName") + ".txt";
            FileOperationSubRequestType fileOperationSubRequest = SharedTestSuiteHelper.CreateFileOperationSubRequest(FileOperationRequestTypes.Rename, newName, null, this.Site);
            cellStoreageResponse = Adapter.CellStorageRequest(this.DefaultFileUrl, new SubRequestType[] { fileOperationSubRequest });
            FileOperationSubResponseType fileOperationSubResponse = SharedTestSuiteHelper.ExtractSubResponse<FileOperationSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                "Rename file should be succeed.");

            this.DefaultFileUrl = this.DefaultFileUrl.Substring(0, this.DefaultFileUrl.LastIndexOf('/') + 1) + newName;
            cellStoreageResponse = Adapter.CellStorageRequest(
                this.DefaultFileUrl,
                new SubRequestType[] { getVersionsSubRequest },
                "1", 2, 2, null, null, null, null, null, null, true);
            getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);
            this.Site.Assert.AreEqual<ErrorCodeType>(
                ErrorCodeType.Success,
                SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                "GetVersions should be succeed.");

            string resourceID2 = cellStoreageResponse.ResponseCollection.Response[0].ResourceID;

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11026
                Site.CaptureRequirementIfAreEqual<string>(
                    resourceID,
                    resourceID2,
                    "MS-FSSHTTP",
                    11026,
                    @"[In Response] [ResourceID] A ResourceID MUST NOT change over the lifetime of a file, even if the URL of the file changes.");
            }
            else
            {
                Site.Assert.AreEqual<string>(
                    resourceID, resourceID2, "ResourceID MUST NOT change over the lifetime of a file, even if the URL of the file changes.");
            }
        }

        /// <summary>
        /// A method used to verify that an error when neither of the ResourceID and Url attributes identify valid files.
        /// </summary>
        [TestCategory("SHAREDTESTCASE"), TestMethod()]
        public void TestCase_S17_TC04_ResourceIdDoesNotExist()
        {
            string invalidUrl = this.DefaultFileUrl + "Invalid";

            // Initialize the service
            this.InitializeContext(this.DefaultFileUrl, this.UserName01, this.Password01, this.Domain);

            // Invoke "GetVersions"sub-request with correct invalid parameters.
            GetVersionsSubRequestType getVersionsSubRequest = SharedTestSuiteHelper.CreateGetVersionsSubRequest(SequenceNumberGenerator.GetCurrentToken());
            CellStorageResponse cellStoreageResponse = Adapter.CellStorageRequest(
               invalidUrl,
                new SubRequestType[] { getVersionsSubRequest },
                "1", 2, 2, null, null, null, null, null, "test", true);
            GetVersionsSubResponseType getVersionsSubResponse = SharedTestSuiteHelper.ExtractSubResponse<GetVersionsSubResponseType>(cellStoreageResponse, 0, 0, this.Site);

            VersionType version = cellStoreageResponse.ResponseVersion as VersionType;
            Site.Assume.AreEqual<ushort>(3, version.MinorVersion, "This test case runs only when MinorVersion is 3 which indicates the protocol server is capable of performing ResourceID specific behavior.");

            if (SharedContext.Current.IsMsFsshttpRequirementsCaptured)
            {
                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R2171
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ResourceIdDoesNotExist,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         2171,
                         @"[In GenericErrorCodeTypes] InvalidArgument indicates an error when any of the cell storage service subrequests for the targeted URL for the file contains input parameters that are not valid.");

                // Verify MS-FSSHTTP requirement: MS-FSSHTTP_R11023
                Site.CaptureRequirementIfAreEqual<ErrorCodeType>(
                         ErrorCodeType.ResourceIdDoesNotExist,
                         SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                         "MS-FSSHTTP",
                         11023,
                         @"[In Request] [UseResourceID]In the case where the protocol server is using the ResourceID attribute but the ResourceID attribute does not identify a valid file the protocol server SHOULD set an error code in the ErrorCode attribute of the corresponding Response attribute.");

            }
            else
            {
                Site.Assert.AreEqual<ErrorCodeType>(
                    ErrorCodeType.ResourceIdDoesNotExist,
                    SharedTestSuiteHelper.ConvertToErrorCodeType(getVersionsSubResponse.ErrorCode, this.Site),
                    @"[In GenericErrorCodeTypes] InvalidArgument indicates an error when any of the cell storage service subrequests for the targeted URL for the file contains input parameters that are not valid.");
            }
        }
        #endregion
    }
}