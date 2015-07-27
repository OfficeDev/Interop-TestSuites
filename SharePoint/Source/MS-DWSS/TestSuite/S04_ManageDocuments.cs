//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_DWSS
{
    using System;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Provides test methods for validating the operation: FindDwsDoc. 
    /// </summary>
    [TestClass]
    public class S04_ManageDocuments : TestClassBase
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

        #region Test FindDwsDoc Operation
        
        /// <summary>
        /// This test case is intended to validate that the returning absolute URL of a document listed in the documents parameter of a previous call to the CreateDws method successfully.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S04_TC01_FindDwsDoc_ValidId()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            string dwsTitle = Common.GetConfigurationPropertyValue("ValidTitle", Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, dwsTitle, documents, out error);
            
            // Redirect the web service to the newly created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Find the document with valid document id.
            string findDwsDocResult = this.dwsAdapter.FindDwsDoc(documents.ID, out error);
            this.Site.Assert.IsNull(error, "The response should not be an error!");

            Uri uriAddress = new Uri(findDwsDocResult);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R536");
            
            // Verify MS-DWSS requirement: MS-DWSS_R536
            this.Site.CaptureRequirementIfIsTrue(
                uriAddress.IsAbsoluteUri,
                536,
                @"[In FindDwsDocResponse] Result: A Result element for which the content MUST be an absolute URL that refers to the requested document.");
            
            if (Common.IsRequirementEnabled(687, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R687");
                
                // Verify MS-DWSS requirement: MS-DWSS_R687
                this.Site.CaptureRequirementIfIsTrue(
                    uriAddress.IsAbsoluteUri,
                    687,
                    @"[In Appendix B: Product Behavior] Implementation does reply with a Result element as specified in FindDwsDocResponse containing an absolute URL to the specified document. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
            
            // Find the document without document id.
            this.dwsAdapter.FindDwsDoc(null, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error!");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The response should not be an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with ItemNotFound code should be returned if the server cannot locate a document from the specified document ID.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S04_TC02_FindDwsDoc_ItemNotFound()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string docId = "InvalidDocumentID";
            
            // Request the web site with a nonexistent document ID; the server will return an ItemNotFound Error element.
            this.dwsAdapter.FindDwsDoc(docId, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // If the error is not null, then the server did returned an Error element as is specified.
            this.Site.CaptureRequirement(
                535,
                @"[In FindDwsDocResponse] Error: An Error element as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R183");
            
            // Verify MS-DWSS requirement: MS-DWSS_R183
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ItemNotFound,
                error.Value,
                183,
                @"[In FindDwsDoc] If the protocol server cannot locate a document with the specified identifier, it MUST return an Error element with a code of ItemNotFound.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R315");
            
            // Verify MS-DWSS requirement: MS-DWSS_R315
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                315,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains ItemNotFound error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R24");
            
            // Verify MS-DWSS requirement: MS-DWSS_R24
            this.Site.CaptureRequirementIfAreEqual<string>(
                "5",
                error.ID,
                24,
                @"[In Error] The value 5 [ID] matches the Error type ItemNotFound.");
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