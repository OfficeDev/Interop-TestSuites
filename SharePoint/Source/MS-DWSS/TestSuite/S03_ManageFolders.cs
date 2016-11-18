namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Provides test methods for validating the operation: CreateFolder and DeleteFolder. 
    /// </summary>
    [TestClass]
    public class S03_ManageFolders : TestClassBase
    {
        #region Variables
        
        /// <summary>
        /// Adapter Instance.
        /// </summary>
        private IMS_DWSSAdapter dwsAdapter;

        #endregion Variables

        #region Test Suite Initialization
        
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }
        
        /// <summary>
        /// Use ClassCleanup to run code after all tests in a class have run.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion Test Suite Initialization

        #region Test Cases

        #region Test DeleteFolder Operation
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CreateFolder operation to create a subfolder in the document library of the current Document Workspace site successfully.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC01_CreateFolder_CreateFolderSuccessfully()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the newly created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Create a sub folder in the web site.
            string folderUrl = Common.GetConfigurationPropertyValue("ValidFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a <Result/>, not an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R155");
            
            // Verify MS-DWSS requirement: MS-DWSS_R155
            this.Site.CaptureRequirementIfIsNull(
                error,
                155,
                @"[In CreateFolder] If none of the prior conditions [FolderNotFound, AlreadyExists, NoAccess, Failed or ServerFailure] apply, the protocol server MUST create the folder specified in the CreateFolder element.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with FolderNotFound code should be returned by CreateFolder operation if the server cannot locate the folder.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC02_CreateFolder_FolderNotFound()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);

            Error error;
            
            // Request the web site with the nonexistent folder URL; the server will return a FolderNotFound Error element.
            string folderUrl = Common.GetConfigurationPropertyValue("ValidDocumentLibraryName", this.Site);
            
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNotNull(error, "The server should return an error!");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R149");
            
            // Verify MS-DWSS requirement: MS-DWSS_R149
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.FolderNotFound,
                error.Value,
                149,
                @"[In CreateFolder] If the parent folder for the specified URL does not exist, the protocol server MUST return an Error element with a FolderNotFound error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R29");
            
            // Verify MS-DWSS requirement: MS-DWSS_R29
            this.Site.CaptureRequirementIfAreEqual<string>(
                "10",
                error.ID,
                29,
                @"[In Error] The value 10 [ID] matches the Error type FolderNotFound.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R320");
            
            // Verify MS-DWSS requirement: MS-DWSS_R320
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                320,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains FolderNotFound error code.");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with AlreadyExists code should be returned by CreateFolder operation if the specified URL already exists.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC03_CreateFolder_AlreadyExists()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the newly created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Create a folder in a specified folder URL.
            string folderUrl = Common.GetConfigurationPropertyValue("ValidFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a <Result/>, not an error.");
            
            // Create a folder using the same folder URL, then the server will return a AlreadyExists Error element.
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected an error.");
            
            // If error is not null, it indicates that the server did returns an Error element as is specified.
            this.Site.CaptureRequirement(
                159,
                @"[In CreateFolderResponse] Error: An Error element as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R150");
            
            // Verify MS-DWSS requirement: MS-DWSS_R150
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.AlreadyExists,
                error.Value,
                150,
                @"[In CreateFolder] If the specified URL already exists, the protocol server MUST return an Error element with an AlreadyExists error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R32");
            
            // Verify MS-DWSS requirement: MS-DWSS_R32
            this.Site.CaptureRequirementIfAreEqual<string>(
                "13",
                error.ID,
                32,
                @"[In Error] The value 13 [ID] matches the Error type AlreadyExists.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R323");
            
            // Verify MS-DWSS requirement: MS-DWSS_R323
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                323,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains AlreadyExists error code.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with NoAccess code should be returned by CreateFolder operation when the user does not have sufficient access permissions to create the folder.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC04_CreateFolder_NoAccess()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Set Dws service credential to Reader credential which does not have permission to create a folder.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            Error error;
            string folderUrl = Common.GetConfigurationPropertyValue("NewFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R151");
            
            // Verify MS-DWSS requirement: MS-DWSS_R151
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.NoAccess,
                error.Value,
                151,
                @"[In CreateFolder] If the user does not have sufficient access permissions to create the folder, the protocol server MUST return an Error element with a NoAccess error code.");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with ServerFailure code should be returned by CreateFolder operation.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC05_CreateFolder_ServerFailure()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Set Dws service credential to None credential.
            string userName = Common.GetConfigurationPropertyValue("NoneRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("NoneRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            Error error;
            string folderUrl = Common.GetConfigurationPropertyValue("ValidFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R154");
            
            // Verify MS-DWSS requirement: MS-DWSS_R154
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                154,
                @"[In CreateFolder] The protocol server MUST return an Error element with a ServerFailure error code if an unspecified error prevents creating the specified folder.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CreateFolder operation to validate that if URL contains multiple folders and the first folder in the URL doesn't exist in the site, the server will ignore the first folder in the URL and replace it with "Shared Documents".
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC06_CreateFolder_InvalidParentFolderUrl()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the newly created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Create a sub folder in the web site.
            string folderUrl = "InvalidParentFolder/MSDWSS_TestFolder";
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a <Result/>, not an error.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by DeleteFolder operation while deleting a subfolder in the document library of the current Document Workspace site successfully.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC07_DeleteFolder_DeleteFolderSuccessfully()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the newly created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);

            string folderUrl = Common.GetConfigurationPropertyValue("ValidFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            // If the server creates the folder successfully, then validates that the server can delete the specified folder using the DeleteFolder operation.
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a <Result/>, not an error.");
            
            string invalidFolderUrl = "MSDWSS_DocumentLibrary/MSDWSS_NotExistFolder";
            this.dwsAdapter.DeleteFolder(invalidFolderUrl, out error);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R173");
            
            // Verify MS-DWSS requirement: MS-DWSS_R173
            this.Site.CaptureRequirementIfIsNull(
                error,
                173,
                @"[In DeleteFolder] If the specified URL does not exist, the protocol server MUST return a Result element as specified in DeleteFolderResponse (section 3.1.4.5.2.2).");

            this.dwsAdapter.DeleteFolder(folderUrl, out error);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R177");
            
            // Verify MS-DWSS requirement: MS-DWSS_R177
            this.Site.CaptureRequirementIfIsNull(
                error,
                177,
                @"[In DeleteFolder] If none of the prior conditions apply, the protocol server MUST delete the folder specified in the CreateFolder element and return a Result element as specified in DeleteFolderResponse (section 3.1.4.5.2.2).");
            
            // Delete the created folder without url.
            this.dwsAdapter.DeleteFolder(null, out error);
            this.Site.Assert.IsNotNull(error, "The server should return an error.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by DeleteFolder operation while the parent folder for the specified URL does not exist.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S03_TC08_DeleteFolder_FolderNotFound()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);

            Error error;
            
            // Request the web site with the invalid folder URL; the server will return a FolderNotFound Error element.
            string folderUrl = Common.GetConfigurationPropertyValue("ValidDocumentLibraryName", this.Site);
            this.dwsAdapter.DeleteFolder(folderUrl, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // If the error is not null, it indicates that the server returns an Error element as is specified.
            this.Site.CaptureRequirement(
                1523,
                @"[In DeleteFolderResponse] Error: An Error element as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R172");
            
            // Verify MS-DWSS requirement: MS-DWSS_R172
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.FolderNotFound,
                error.Value,
                172,
                @"[In DeleteFolder] If the parent of the specified URL does not exist, the protocol server MUST return an Error element with a FolderNotFound error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R29");
            
            // Verify MS-DWSS requirement: MS-DWSS_R29
            this.Site.CaptureRequirementIfAreEqual<string>(
                "10",
                error.ID,
                29,
                @"[In Error] The value 10 [ID] matches the Error type FolderNotFound.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R320");
            
            // Verify MS-DWSS requirement: MS-DWSS_R320
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                320,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains FolderNotFound error code.");
        }

        #endregion

        #endregion Test Cases

        #region Test Case Initialization
        
        /// <summary>
        /// Initialize Test case and test environment
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.dwsAdapter = Site.GetAdapter<IMS_DWSSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            // Set default Dws service credential to admin credential.
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
        }
        
        /// <summary>
        /// Clean up test environment.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.dwsAdapter.Reset();
        }

        #endregion Test Case Initialization
    }
}