namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to get and delete versions for a specified file with valid input parameters.
    /// </summary>
    [TestClass]
    public class S01_DeleteVersion : TestClassBase
    {
        #region Variables
        /// <summary>
        /// The instance of the SUT control adapter.
        /// </summary>
        private IMS_VERSSSUTControlAdapter sutControlAdapterInstance;

        /// <summary>
        /// The instance of the protocol adapter.
        /// </summary>
        private IMS_VERSSAdapter protocolAdapterInstance;

        /// <summary>
        /// The instance of ILISTSWSSUTControlAdapter interface.
        /// </summary>
        private IMS_LISTSWSSUTControlAdapter listsSutControlAdaterInstance;

        /// <summary>
        /// The name of list in the site.
        /// </summary>
        private string documentLibrary;

        /// <summary>
        /// The name of file in the list.
        /// </summary>
        private string fileName;

        /// <summary>
        /// The relative name of the file.
        /// </summary>
        private string fileRelativeUrl;

        /// <summary>
        /// The absolute name of the file.
        /// </summary>
        private string fileAbsoluteUrl;
       
        /// <summary>
        /// The absolute URL for the site collection.
        /// </summary>
        private string requestUrl;

        /// <summary>
        /// The instance of the TestSuiteHelper class.
        /// </summary>
        private TestSuiteHelper testSuiteHelper;
        #endregion

        #region Test Suite Initialization
        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="testContext">Context information associated with MS-VERSS.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clean up the test class after all test cases finished running.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region The test cases in S01_DeleteVersion scenario.
        /// <summary>
        /// A test case used to test that the client can get expected DeleteVersionSoapOut and GetVersionsSoapOut messages by calling DeleteVersion and GetVersions operations with the relative URL of a file.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S01_TC01_DeleteVersionUsingRelativeUrl()
        {
            this.DeleteVersionVerification(this.fileRelativeUrl);
        }
        
        /// <summary>
        /// A test case used to test that the client can get expected DeleteVersionSoapOut and GetVersionsSoapOut messages by calling DeleteVersion and GetVersions operations with the absolute URL of a file.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S01_TC02_DeleteVersionUsingAbsoluteUrl()
        {
            this.DeleteVersionVerification(this.fileAbsoluteUrl);
        }

        /// <summary>
        /// A test case used to test that the client can get expected DeleteAllVersionsSoapOut and GetVersionsSoapOut messages by calling DeleteAllVersions and GetVersions operations with the relative URL of a file.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S01_TC03_DeleteAllVersionsUsingRelativeUrl()
        {
            this.DeleteAllVersionsVerification(this.fileRelativeUrl);
        }

        /// <summary>
        /// A test case used to test that the client can get expected DeleteAllVersionsSoapOut and GetVersionsSoapOut messages by calling DeleteAllVersions and GetVersions operations with the absolute URL of a file.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S01_TC04_DeleteAllVersionsUsingAbsoluteUrl()
        {
            this.DeleteAllVersionsVerification(this.fileAbsoluteUrl);
        }
        #endregion

        #region Test Case Initialization

        /// <summary>
        /// Initialize test case and test environment.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.sutControlAdapterInstance = this.Site.GetAdapter<IMS_VERSSSUTControlAdapter>();
            this.protocolAdapterInstance = this.Site.GetAdapter<IMS_VERSSAdapter>();
            Common.CheckCommonProperties(this.Site, true);
            this.listsSutControlAdaterInstance = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();

            #region Initialize the variables
            string datetimestamp = Common.FormatCurrentDateTime();

            this.requestUrl = Common.GetConfigurationPropertyValue("RequestUrl", this.Site);
            this.documentLibrary = Common.GetConfigurationPropertyValue("DocumentLibraryName", this.Site) +
                "_" + datetimestamp;
            string fileNameValue = Common.GetConfigurationPropertyValue("FileName", this.Site);
            this.fileName = System.IO.Path.GetFileNameWithoutExtension(fileNameValue) +
                "_" + datetimestamp +
                System.IO.Path.GetExtension(fileNameValue);

            this.fileRelativeUrl = this.documentLibrary + "/" + this.fileName;
            Uri requestUri = new Uri(this.requestUrl);
            this.fileAbsoluteUrl = AdapterHelper.ConstructDocFileFullUrl(requestUri, this.documentLibrary, this.fileName).AbsoluteUri;
            #endregion

            #region Initialize the server
            bool isAddList = this.listsSutControlAdaterInstance.AddList(this.documentLibrary);
            Site.Assume.IsTrue(
                isAddList,
                "AddList operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isAddList);

            this.testSuiteHelper = new TestSuiteHelper(
                this.Site,
                this.documentLibrary,
                this.fileName,
                this.listsSutControlAdaterInstance,
                this.protocolAdapterInstance,
                this.sutControlAdapterInstance);
            #endregion
        }

        /// <summary>
        /// Clean up test environment.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.testSuiteHelper.CleanupTestEnvironment();
        }
        #endregion

        /// <summary>
        /// A common method used to verify that the client can get expected DeleteVersionSoapOut and 
        /// GetVersionsSoapOut messages by calling DeleteVersion and GetVersions operations with the URL of a file.
        /// </summary>
        /// <param name="url">The URL of a file.</param>
        private void DeleteVersionVerification(string url)
        {
            // Enable the versioning of the list.
            bool setVersioning = this.sutControlAdapterInstance.SetVersioning(this.documentLibrary, true, true);
            Site.Assert.IsTrue(
                setVersioning,
                "SetVersioning operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                setVersioning);

            // Upload the file into specific list.
            bool isAddFileSuccessful = this.sutControlAdapterInstance.AddFile(this.documentLibrary, this.fileName, TestSuiteHelper.UploadFileName);
            Site.Assert.IsTrue(
                isAddFileSuccessful,
                "AddFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isAddFileSuccessful);

            // Check out and check in file one time to create a new version of the file. 
            this.testSuiteHelper.AddOneFileVersion(this.fileName);

            // Call GetVersions with the relative filename to get details about all versions of the file.
            GetVersionsResponseGetVersionsResult getVersionsResponse =
                this.protocolAdapterInstance.GetVersions(url);

            // Verify the GetVersions response results.
            this.testSuiteHelper.VerifyResultsInformation(getVersionsResponse.results, OperationName.GetVersions, true);

            // Get the current version information by using the results element in the response of GetVersions.
            string currentVersion = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);

            // Enable the Recycle Bin.
            bool isRecycleBinEnable = this.sutControlAdapterInstance.SetRecycleBinEnable(true);
            Site.Assert.IsTrue(
                isRecycleBinEnable,
                "SetRecycleBinEnable operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isRecycleBinEnable);

            int waitTime = Common.GetConfigurationPropertyValue<int>("WaitTime", this.Site);
            int retryCount = Common.GetConfigurationPropertyValue<int>("RetryCount", this.Site);

            while (retryCount > 0)
            {
                retryCount--;
                System.Threading.Thread.Sleep(waitTime);

                isRecycleBinEnable = this.sutControlAdapterInstance.GetRecycleBin();
                if (isRecycleBinEnable)
                {
                    break;
                }
            }
            Site.Assert.IsTrue(isRecycleBinEnable, "Recycle bin should be enable.");

            // Get the version that needs to be deleted.
            string deleteFileVersion = AdapterHelper.GetPreviousVersion(getVersionsResponse.results.result);

            // Call DeleteVersion to delete a specific version of the file by using the relative filename.
            DeleteVersionResponseDeleteVersionResult deleteVersionResponse =
                this.protocolAdapterInstance.DeleteVersion(url, deleteFileVersion);

            // Verify DeleteVersion response results.
            this.testSuiteHelper.VerifyResultsInformation(deleteVersionResponse.results, OperationName.DeleteVersion, true);

            // Check whether the current version exists in the DeleteVersion response.
            bool isCurrentVersionExist =
                AdapterHelper.IsVersionExist(deleteVersionResponse.results.result, currentVersion);

            Site.Assert.IsTrue(
                isCurrentVersionExist,
                "The DeleteVersion operation should not delete the current version {0}",
                currentVersion);

            // Check whether the deleted version exists in the Recycle Bin.
            bool isDeleteFileVersionExistInRecycleBin =
                this.sutControlAdapterInstance.IsFileExistInRecycleBin(this.fileName, deleteFileVersion);

            // Verify MS-VERSS requirement: MS-VERSS_R173
            Site.CaptureRequirementIfIsTrue(
                isDeleteFileVersionExistInRecycleBin,
                173,
                @"[In DeleteVersion operation] If the Recycle Bin is enabled, the version is placed in the Recycle Bin, instead.");
        }

        /// <summary>
        /// A common method used to verify that the client can get expected DeleteAllVersionsSoapOut and 
        /// GetVersionsSoapOut messages by calling DeleteAllVersions and GetVersions operations with the URL of a file.
        /// </summary>
        /// <param name="url">The URL of a file.</param>
        private void DeleteAllVersionsVerification(string url)
        {
            // Enable the versioning of the list.
            bool setVersioning = this.sutControlAdapterInstance.SetVersioning(this.documentLibrary, true, true);
            Site.Assert.IsTrue(
                setVersioning,
                "SetVersioning operation returns {0}, TRUE means the operation was executed successfully, " +
                "FALSE means the operation failed",
                setVersioning);

            // Upload the file into specific list.
            bool isAddFileSuccessful = this.sutControlAdapterInstance.AddFile(this.documentLibrary, this.fileName, TestSuiteHelper.UploadFileName);
            Site.Assert.IsTrue(
                isAddFileSuccessful,
                "AddFile operation returns {0}, TRUE means the operation was executed successfully, " +
                " FALSE means the operation failed",
                isAddFileSuccessful);

            // Check out and check in file one time to create a new version of the file. 
            this.testSuiteHelper.AddOneFileVersion(this.fileName);

            // Call SUT Control Adapter method SetFilePublish to publish the current version of the file.
            bool isFilePublished = this.sutControlAdapterInstance.SetFilePublish(this.documentLibrary, this.fileName, true);
            Site.Assert.IsTrue(
                isFilePublished,
                "SetFilePublish operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isFilePublished);

            // Call GetVersions with the relative filename to get details about all versions of the file.
            GetVersionsResponseGetVersionsResult getVersionsResponse = this.protocolAdapterInstance.GetVersions(url);

            // Verify the GetVersions response results.
            this.testSuiteHelper.VerifyResultsInformation(getVersionsResponse.results, OperationName.GetVersions, true);

            // Get previous version before DeleteAllVersions operation by using the results element in the response of GetVersions.
            string previousVersion = AdapterHelper.GetPreviousVersion(getVersionsResponse.results.result);

            // Get the published version information by using the results element in the response of GetVersions.
            string publishedVersion = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);

            // Check out and check in file one time to create a new version of the file. 
            this.testSuiteHelper.AddOneFileVersion(this.fileName);

            // Call GetVersions with the relative filename to get details about all versions of the file.
            getVersionsResponse = this.protocolAdapterInstance.GetVersions(url);

            // Get the current version information by using the results element in the response of GetVersions.
            string currentVersion = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);

            // Enable the Recycle Bin.
            bool isRecycleBinEnable = this.sutControlAdapterInstance.SetRecycleBinEnable(true);
            Site.Assert.IsTrue(
                isRecycleBinEnable,
                "SetRecycleBinEnable operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isRecycleBinEnable);

            // Call DeleteAllVersions operation with the relative filename to delete all the previous versions except
            // the published version and the current version.
            DeleteAllVersionsResponseDeleteAllVersionsResult deleteAllVersionsResponse =
                this.protocolAdapterInstance.DeleteAllVersions(url);

            // Verify DeleteAllVersions response results.
            this.testSuiteHelper.VerifyResultsInformation(deleteAllVersionsResponse.results, OperationName.DeleteAllVersions, true);

            // Verify whether the published version and the current version exist in the results element in the response 
            // of DeleteAllVersions.
            bool isCurrentVersionExist =
                AdapterHelper.IsVersionExist(deleteAllVersionsResponse.results.result, currentVersion);
            bool isPublishedVersionExist = AdapterHelper.IsVersionExist(
                deleteAllVersionsResponse.results.result,
                publishedVersion);

            Site.Assert.IsTrue(
                isCurrentVersionExist,
                "The DeleteAllVersions operation should not delete the current version {0}",
                currentVersion);

            Site.Assert.IsTrue(
                isPublishedVersionExist,
                "The DeleteAllVersions operation should not delete the published version {0}",
                publishedVersion);

            // Verify whether the previous version exists in the results element in the response of DeleteAllVersions.
            bool isPreviousVersionExist = AdapterHelper.IsVersionExist(deleteAllVersionsResponse.results.result, previousVersion);
            Site.Assert.IsFalse(
                isPreviousVersionExist,
                "The DeleteAllVersions operation should delete the previous version {0}",
                previousVersion);

            // Since all the previous versions of the specified file do not exist, except for the published version and the current version, capture requirement MS-VERSS_R79.
            Site.CaptureRequirement(
                79,
                @"[In DeleteAllVersions operation] The DeleteAllVersions operation deletes all the previous versions of the specified file, except for the published version and the current version. ");

            bool isDeleteFileVersionExistInRecycleBin = false;
            bool isDeleted = false;

            foreach (VersionData versionData in getVersionsResponse.results.result)
            {
                isDeleted = !AdapterHelper.IsVersionExist(deleteAllVersionsResponse.results.result, versionData.version);
                if (isDeleted)
                {
                    // Verify whether the deleted versions exist in the Recycle Bin.
                    isDeleteFileVersionExistInRecycleBin = this.sutControlAdapterInstance.IsFileExistInRecycleBin(
                        this.fileName,
                        versionData.version);

                    if (!isDeleteFileVersionExistInRecycleBin)
                    {
                        break;
                    }
                }
            }

            // Verify MS-VERSS requirement: MS-VERSS_R169
            Site.CaptureRequirementIfIsTrue(
                isDeleteFileVersionExistInRecycleBin,
                169,
                @"[In DeleteAllVersions operation] If the Recycle Bin is enabled, the versions are placed in the Recycle Bin, instead.");
        }
    }
}