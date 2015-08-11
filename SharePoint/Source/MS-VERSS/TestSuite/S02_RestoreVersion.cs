namespace Microsoft.Protocols.TestSuites.MS_VERSS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to get and restore versions for a specified file with valid input parameters. 
    /// </summary>
    [TestClass]
    public class S02_RestoreVersion : TestClassBase
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
        private Uri fileAbsoluteUrl;

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

        #region The test cases in S02_RestoreVersion scenario.

        /// <summary>
        ///  A test case used to test that the client uses the RestoreVersion operation to successfully restore the file specified by a relative URL to a specific version when check out is enforced.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S02_TC01_RestoreVersionUsingRelativeUrl()
        {
            this.RestoreVersionVerification(this.fileRelativeUrl);
        }

        /// <summary>
        /// A test case used to test that the client uses the RestoreVersion operation to successfully restore the file specified by an absolute URL to a specific version when check out is enforced.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S02_TC02_RestoreVersionUsingAbsoluteUrl()
        {
            this.RestoreVersionVerification(this.fileAbsoluteUrl.AbsoluteUri);
        }

        /// <summary>
        /// A test case used to test that the client uses the RestoreVersion operation to successfully restore
        /// the file specified by an absolute URL to a specific version when check out is not enforced.
        /// </summary>
        [TestCategory("MSVERSS"), TestMethod()]
        public void MSVERSS_S02_TC03_RestoreVersionWithoutEnforceCheckout()
        {
            // Enable the versioning of the list.
            bool setVersioning = this.sutControlAdapterInstance.SetVersioning(this.documentLibrary, true, true);
            Site.Assert.IsTrue(
                setVersioning, 
                "SetVersioning operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                setVersioning);

            // Protocol server does not enforce that only checked out files can be modified.
            bool isSetServerEnforceCheckOut = this.sutControlAdapterInstance.SetEnforceCheckout(this.documentLibrary, false);
            Site.Assert.IsTrue(
                isSetServerEnforceCheckOut, 
                "SetServerEnforceCheckOut operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isSetServerEnforceCheckOut);

            // Upload the first file into specific list.
            bool isAddFileSuccessful = this.sutControlAdapterInstance.AddFile(this.documentLibrary, this.fileName, TestSuiteHelper.UploadFileName);
            Site.Assert.IsTrue(
                isAddFileSuccessful,
                "AddFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isAddFileSuccessful);

            // Check out and check in first file to add the file versions.
            this.testSuiteHelper.AddOneFileVersion(this.fileName);

            // Call GetVersions operation using absolute fileName.
            GetVersionsResponseGetVersionsResult getVersionsResponse = this.protocolAdapterInstance.GetVersions(this.fileAbsoluteUrl.AbsoluteUri);

            // Verify the GetVersions results.
            this.testSuiteHelper.VerifyResultsInformation(getVersionsResponse.results, OperationName.GetVersions, true);

            // Get the current version.
            string currentVersionBeforeRestore = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);

            // Get the previous version which is specific to be restored.
            string restoreVersion = AdapterHelper.GetPreviousVersion(getVersionsResponse.results.result);

            // Call RestoreVersion operation using absolute fileName and restore a specified file to a specific version.
            RestoreVersionResponseRestoreVersionResult restoreVersionNotCheckOutResponse = 
                this.protocolAdapterInstance.RestoreVersion(this.fileAbsoluteUrl.AbsoluteUri, restoreVersion);

            // Verify the RestoreVersion results.
            this.testSuiteHelper.VerifyResultsInformation(restoreVersionNotCheckOutResponse.results, OperationName.RestoreVersion, true);

            // Get the current version in RestoreVersion response.
            string currentVersionAfterNotCheckOutRestore = AdapterHelper.GetCurrentVersion(
                restoreVersionNotCheckOutResponse.results.result);

            // Verify whether the current version after RestoreVersion is increased when the file is not checked out.
            bool isCurrentVersionIncreased = AdapterHelper.IsCurrentVersionIncreased(
                currentVersionBeforeRestore,
                currentVersionAfterNotCheckOutRestore);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug, 
                "Verify MS-VERSS_R185, if the file is not checked out, the current version before RestoreVersion is {0}," +
                " and after RestoreVersion is {1}",
                currentVersionBeforeRestore,
                currentVersionAfterNotCheckOutRestore);

            // Verify MS-VERSS requirement: MS-VERSS_R185
            Site.CaptureRequirementIfIsTrue(
                isCurrentVersionIncreased,
                185,
                @"[In RestoreVersion] If the file is not checked out before the restoration, the current version number of the file MUST still be increased, as with any other change.");
 
            // Check out the file.
            bool isCheckOutFile = this.listsSutControlAdaterInstance.CheckoutFile(this.fileAbsoluteUrl);
            Site.Assert.IsTrue(
                isCheckOutFile,
                "CheckOutFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckOutFile);

            // Call GetVersions operation using absolute fileName.
            getVersionsResponse = this.protocolAdapterInstance.GetVersions(this.fileAbsoluteUrl.AbsoluteUri);

            // Get current version before RestoreVersion operation in GetVersions response.
            currentVersionBeforeRestore = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);
            
            // Get the previous version which is specific to be restored.
            restoreVersion = AdapterHelper.GetPreviousVersion(getVersionsResponse.results.result);

            // Call RestoreVersion operation using absolute fileName and restore a specified file to a specific version.
            RestoreVersionResponseRestoreVersionResult restoreVersionIsCheckOutResponse =
                this.protocolAdapterInstance.RestoreVersion(this.fileAbsoluteUrl.AbsoluteUri, restoreVersion);

            // Get the current version in RestoreVersion response.
            string currentVersionAfterIsCheckOutRestore = AdapterHelper.GetCurrentVersion(
                restoreVersionIsCheckOutResponse.results.result);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-VERSS_R186, if the file is checked out, the current version before RestoreVersion is {0}," +
                " and after RestoreVersion is {1}",
                currentVersionBeforeRestore, 
                currentVersionAfterIsCheckOutRestore);

            // Verify MS-VERSS requirement: MS-VERSS_R186
            Site.CaptureRequirementIfAreEqual<string>(
                currentVersionBeforeRestore,
                currentVersionAfterIsCheckOutRestore,
                186,
                @"[In RestoreVersion] If the file is checked out, the current version number of the file after restoration MUST remain the same as before restoration.");

            // Check in file.
            bool isCheckInFile = this.listsSutControlAdaterInstance.CheckInFile(
                this.fileAbsoluteUrl,
                TestSuiteHelper.FileComments,
                ((int)VersionType.MinorCheckIn).ToString());
            Site.Assert.IsTrue(
                isCheckInFile,
                "CheckInFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckInFile);
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
            this.fileAbsoluteUrl = AdapterHelper.ConstructDocFileFullUrl(requestUri, this.documentLibrary, this.fileName);
            #endregion

            #region Initialize the server
            bool isAddList = this.listsSutControlAdaterInstance.AddList(this.documentLibrary);
            Site.Assume.IsTrue(
                isAddList,
                "AddList operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isAddList);
            #endregion

            this.testSuiteHelper = new TestSuiteHelper(
                this.Site,
                this.documentLibrary,
                this.fileName,
                this.listsSutControlAdaterInstance,
                this.protocolAdapterInstance,
                this.sutControlAdapterInstance);
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
        /// A common method used to verify that the client uses the RestoreVersion operation to successfully restore
        /// the file specified by a URL to a specific version when check out is enforced.
        /// </summary>
        /// <param name="url">The URL of a file.</param>
        private void RestoreVersionVerification(string url)
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

            // Enforce the file to be checked out.
            bool setServerEnforceCheckOut = this.sutControlAdapterInstance.SetEnforceCheckout(this.documentLibrary, true);
            Site.Assert.IsTrue(
                setServerEnforceCheckOut,
                "SetServerEnforceCheckOut operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                setServerEnforceCheckOut);

            // Call GetVersions operation using relative fileName.
            GetVersionsResponseGetVersionsResult getVersionsResponse = this.protocolAdapterInstance.GetVersions(
                url);

            // Verify the GetVersions results.
            this.testSuiteHelper.VerifyResultsInformation(getVersionsResponse.results, OperationName.GetVersions, true);

            // Check out the specified file.
            bool isCheckOutFile = this.listsSutControlAdaterInstance.CheckoutFile(this.fileAbsoluteUrl);
            Site.Assert.IsTrue(
                isCheckOutFile,
                "CheckOutFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckOutFile);

            // Get the current version of the file.
            string currentVersionBeforeRestore = AdapterHelper.GetCurrentVersion(getVersionsResponse.results.result);

            // Get the previous version which is specific to be restored.
            string restoreVersion = AdapterHelper.GetPreviousVersion(getVersionsResponse.results.result);

            // Call RestoreVersion operation using relative fileName and restore a specified file to a specific version.
            RestoreVersionResponseRestoreVersionResult restoreVersionReponse = this.protocolAdapterInstance.RestoreVersion(
                url, restoreVersion);

            // Check in the specified file.
            bool isCheckInFile = this.listsSutControlAdaterInstance.CheckInFile(
                this.fileAbsoluteUrl,
                TestSuiteHelper.FileComments,
                ((int)VersionType.MinorCheckIn).ToString());
            Site.Assert.IsTrue(
                isCheckInFile,
                "CheckInFile operation returns {0}, TRUE means the operation was executed successfully," +
                " FALSE means the operation failed",
                isCheckInFile);

            // Verify the RestoreVersion results.
            this.testSuiteHelper.VerifyResultsInformation(restoreVersionReponse.results, OperationName.RestoreVersion, true);

            // Get the current version in RestoreVersion response.
            string currentVersionAfterRestore = AdapterHelper.GetCurrentVersion(restoreVersionReponse.results.result);

            // Verify whether the current version was increased by RestoreVersion.
            bool isCurrentVersionIncreased = AdapterHelper.IsCurrentVersionIncreased(
                currentVersionBeforeRestore,
                currentVersionAfterRestore);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-VERSS_R182, the current version before RestoreVersion is {0}, and after RestoreVersion is {1}",
                currentVersionBeforeRestore,
                currentVersionAfterRestore);

            // Verify MS-VERSS requirement: MS-VERSS_R182
            Site.CaptureRequirementIfIsTrue(
                isCurrentVersionIncreased,
                182,
                @"[In RestoreVersion] After the restoration, the current version number of the file MUST still be increased, as with any other change.");
        }
    }
}