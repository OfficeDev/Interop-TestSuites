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
    using System.Globalization;
    using System.Net;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Provides test methods for validating the operation: CanCreateDwsUrl, CreateDws, DeleteDws and RenameDws. 
    /// </summary>
    [TestClass]
    public class S01_ManageSites : TestClassBase
    {
        #region Variables
        
        /// <summary>
        /// Adapter Instance.
        /// </summary>
        private IMS_DWSSAdapter dwsAdapter;

        #endregion Variables

        #region Test Suite Initialization
        
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
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

        #region Test CanCreateDwsUrl Operation
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CanCreateDwsUrl operation when using the valid URL.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC01_CanCreateDwsUrl_ValidUrl()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string dwsUrl = Common.GetConfigurationPropertyValue("SiteCollectionName", this.Site) + "_" + Common.FormatCurrentDateTime();

            string createDwsUrl = this.dwsAdapter.CanCreateDwsUrl(dwsUrl, out error);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R421");
            
            // Verify MS-DWSS requirement: MS-DWSS_R421
            this.Site.CaptureRequirementIfAreEqual<string>(
                dwsUrl,
                createDwsUrl,
                421,
                @"[In Message Processing Events and Sequencing Rules] CanCreateDwsUrl: It also returns a URL that is unique for the current site (2).");
            
            // Create a Dws with this url, this operation should be success.
            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(createDwsUrl, null, string.Empty, null, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be an CreateDwsResult, not an error!");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R474");
            
            // Verify MS-DWSS requirement: MS-DWSS_R474
            string expectedDwsUrl = this.dwsAdapter.ServiceUrl.ToLower().Replace(Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site).ToLower(), "/" + createDwsUrl);
            this.Site.CaptureRequirementIfAreEqual<string>(
                expectedDwsUrl.ToLower(),
                createDwsRespResults.Url.ToLower(),
                474,
                @"[In CreateDwsResponse] Url: URL for the new workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R684");
            
            // Verify MS-DWSS requirement: MS-DWSS_R684
            // It indicate the server create the Dws successfully and returns a Results element if the createDwsRespResults is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                createDwsRespResults,
                684,
                @"[In CreateDws] The protocol server MUST create the specified Document Workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R685");
            
            // Verify MS-DWSS requirement: MS-DWSS_R685
            // It indicate the server create the Dws successfully and returns a Results element if the createDwsRespResults is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                createDwsRespResults,
                685,
                @"[In CreateDws] The protocol server MUST create a CreateDwsResponse response message with a Result element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R681");
            
            // Verify MS-DWSS requirement: MS-DWSS_R681
            // It indicate the server create the Dws successfully and returns a Results element if the createDwsRespResults is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                createDwsRespResults,
                681,
                @"[In CreateDws] The protocol server MUST return a CreateDwsResponse response message with a Result element.");
            
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CanCreateDwsUrl operation when using the empty URL.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC02_CanCreateDwsUrl_EmptyUrl()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(695, this.Site), "Test is executed only when R695Enabled is set to true.");

            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);

            Error error;

            // Request the web site with an empty Name; the server returns a newly created web site name, such as GUID.
            string createDwsUrl = this.dwsAdapter.CanCreateDwsUrl(string.Empty, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a site name, not an error!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R695");

            // Verify MS-DWSS requirement: MS-DWSS_R695
            this.Site.CaptureRequirementIfAreNotEqual<string>(
                string.Empty,
                createDwsUrl,
                695,
                @"[In Appendix B: Product Behavior] Implementation does return a created name if the url input parameter is empty, such as a GUID. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node that contains HTTP Error returned by CanCreateDwsUrl operation when the user is not authorized.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC03_CanCreateDwsUrl_HTTPError()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Set Dws service credential to reader credential.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);

            Error error;
            
            try
            {
                this.dwsAdapter.CanCreateDwsUrl(string.Empty, out error);
                this.Site.Assert.Fail("The expected HTTP status code 401 is not returned for CanCreateDwsUrl when the user is not authorized to create the specified Document Workspace.");
            }
            catch (WebException ex)
            {
                if (ex.Response == null)
                {
                    throw;
                }

                this.Site.Assert.IsInstanceOfType(ex.Response, typeof(HttpWebResponse), "The protocol server should respond with an HTTP response, the actual response is '{0}'.", ex.Message);
                
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R101");
                
                // Verify MS-DWSS requirement: MS-DWSS_R101
                this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    ((HttpWebResponse)ex.Response).StatusCode,
                    101,
                    @"[In CanCreateDwsUrl] The protocol server MUST respond with an HTTP 401 error if the user is not authorized to create the specified Document Workspace.");
            }
        }

        #endregion

        #region Test CreateDws Operation
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CreateDws operation using the empty name and empty title.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC04_CreateDws_EmptyNameAndEmptyTitle()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            
            // construct an invalid user
            users.Name = "invalidUser";
            users.Email = "invalidUserEmail";
            
            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, null, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a CreateDwsResult, not an error!");
            
            // Get last element
            string newCreatedDwsName = createDwsRespResults.Url.Substring(createDwsRespResults.Url.LastIndexOf('/') + 1);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R467, the new unique GUID server generated is {0}.", newCreatedDwsName);

            Guid dwsName;
            
            // Verify MS-DWSS requirement: MS-DWSS_R467
            this.Site.CaptureRequirementIfIsTrue(
                Guid.TryParse(newCreatedDwsName, out dwsName),
                467,
                @"[In CreateDws] If the name and title parameters are empty, the protocol server MUST generate a new unique GUID to use as the name of the new workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R475");
            
            // Verify MS-DWSS requirement: MS-DWSS_R475
            Uri listUri;
            bool isVerifiedR475 = Uri.TryCreate(createDwsRespResults.DoclibUrl, UriKind.Relative, out listUri);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR475,
                475,
                @"[In CreateDwsResponse] DoclibUrl: Site-relative URL for the shared documents list (1) associated with the workspace.");
            
            // Get the title of the parent site.
            Results getDwsDataRespResults = this.dwsAdapter.GetDwsData(string.Empty, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R476");
            
            // Verify MS-DWSS requirement: MS-DWSS_R476
            this.Site.CaptureRequirementIfAreEqual<string>(
                getDwsDataRespResults.Title,
                createDwsRespResults.ParentWeb,
                476,
                @"[In CreateDwsResponse] ParentWeb: Title of the site (2) on which the workspace is created.");
            
            this.Site.Assert.IsNotNull(createDwsRespResults.FailedUsers, "The server should return the FailedUsers element.");
            this.Site.Assert.IsTrue(createDwsRespResults.FailedUsers.Length > 0, "The failed users is expected to be more than one.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R478");
            
            // Verify MS-DWSS requirement: MS-DWSS_R478
            this.Site.CaptureRequirementIfAreEqual<string>(
                users.Email,
                createDwsRespResults.FailedUsers[0].User.Email,
                478,
                @"[In CreateDwsResponse] FailedUsers: A list of users from the CreateDws Users field that could not be added to the list of authorized users in the new workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R479");
            
            // Verify MS-DWSS requirement: MS-DWSS_R479
            Uri addUsersUrl;
            bool isVerifiedR479 = Uri.TryCreate(createDwsRespResults.AddUsersUrl, UriKind.Absolute, out addUsersUrl);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR479,
                479,
                @"[In CreateDwsResponse] AddUsersUrl: An absolute URL to a Web page that provides the ability to add users to the workspace.");
            
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by CreateDws operation using empty name and valid title.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC05_CreateDws_EmptyNameAndValidTitle()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string dwsTitle = Common.GetConfigurationPropertyValue("ValidTitle", Site) + "_" + Common.FormatCurrentDateTime();
            
            // Request the web site with empty Name, non-empty and valid Title, the server creates a new web site that is based on Title.
            CreateDwsResultResults createDws1RespResults = this.dwsAdapter.CreateDws(string.Empty, null, dwsTitle, null, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a CreateDwsResult, not an error!");
            
            // Create another DWS with the same title.
            CreateDwsResultResults createDws2RespResults = this.dwsAdapter.CreateDws(string.Empty, null, dwsTitle, null, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a CreateDwsResult, not an error!");
            
            // Get last element
            string newCreatedDwsName = createDws1RespResults.Url.Substring(createDws1RespResults.Url.LastIndexOf('/') + 1);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R459");
            
            // Verify MS-DWSS requirement: MS-DWSS_R459
            this.Site.CaptureRequirementIfAreEqual<string>(
                dwsTitle,
                newCreatedDwsName,
                459,
                @"[In CreateDws] If [name is] empty, the protocol server MUST use the title parameter as the name of the new workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R465");
            
            // Verify MS-DWSS requirement: MS-DWSS_R465
            this.Site.CaptureRequirementIfAreEqual<string>(
                dwsTitle,
                newCreatedDwsName,
                465,
                @"[In CreateDws] If the name parameter is empty, the protocol server MUST use the title parameter as the name of the new workspace when there is no site associated with that name.");
            
            // Get last element
            newCreatedDwsName = createDws2RespResults.Url.Substring(createDws2RespResults.Url.LastIndexOf('/') + 1);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1466");
            
            // Verify MS-DWSS requirement: MS-DWSS_R1466
            bool isVerifiedR1466 = !string.Equals(dwsTitle, newCreatedDwsName, StringComparison.CurrentCultureIgnoreCase) &&
                                   newCreatedDwsName.ToLower(CultureInfo.CurrentCulture).Contains(dwsTitle.ToLower(CultureInfo.CurrentCulture));
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1466,
                1466,
                @"[In CreateDws] If the name parameter is empty, the protocol server MUST generate a different name from the title as the name of the new workspace [if there already is a site associate with that name.]");
            
            // Delete the new created two DWS
            this.dwsAdapter.ServiceUrl = createDws1RespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
            
            this.dwsAdapter.ServiceUrl = createDws2RespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node that contains HTTP Error returned by CreateDws operation when the user is not authorized to create the specified Document Workspace.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC06_CreateDws_HTTPError()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Set Dws service credential to Reader credential.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            Error error;
            string dwsName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site) + "_" + Common.FormatCurrentDateTime();
            string dwsTitle = Common.GetConfigurationPropertyValue("ValidTitle", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            try
            {
                this.dwsAdapter.CreateDws(dwsName, null, dwsTitle, null, out error);
            
                this.Site.Assert.Fail("The expected HTTP status code 401 is not returned for CreateDws when the authenticated user does not have enough permissions to create the Document WorkSpace.");
            }
            catch (WebException webException)
            {
                if (webException.Response == null)
                {
                    throw;
                }

                // The web exception response must not be null.
                this.Site.Assert.IsNotNull(webException.Response, "The server should return http status code 401, the response must not be null web exception. The actual exception is:{0}", webException.Message);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R120");
                
                // Verify MS-DWSS requirement: MS-DWSS_R120
                this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    ((HttpWebResponse)webException.Response).StatusCode,
                    120,
                    @"[In CreateDws] The protocol server MUST reply with an HTTP 401 error if the authenticated user is not authorized to create the Document Workspace.");
                
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R698");
                
                // Verify MS-DWSS requirement: MS-DWSS_R698
                this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    ((HttpWebResponse)webException.Response).StatusCode,
                    698,
                    @"[In Transport] Protocol server faults MUST be returned by using HTTP status codes as specified in [RFC2616] Status Code Definitions section 10 or SOAP faults as specified in [SOAP1.1] SOAP Fault section 4.4 or in [SOAP1.2/1] SOAP Fault section 5.4.");
            }
        }
        
        /// <summary>
        /// This test case is intended to validate the result node that contains ServerFailure Error returned by CreateDws operation when using same input parameters twice to create the Document Workspace.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC07_CreateDws_ServerFailure()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            string dwsName = Common.GetConfigurationPropertyValue("SutComputerName", this.Site) + "_" + Common.FormatCurrentDateTime();
            string dwsTitle = Common.GetConfigurationPropertyValue("ValidTitle", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(dwsName, users, dwsTitle, documents, out error);
            this.Site.Assert.IsNotNull(createDwsRespResults, "The server should return a CreateDws response!");
            
            // Set the name to the name returned by CreateDws operation.
            this.dwsAdapter.CreateDws(dwsName, users, dwsTitle, documents, out error);
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R121");
            
            // Verify MS-DWSS requirement: MS-DWSS_R121
            this.Site.CaptureRequirementIfIsNotNull(
                error,
                121,
                @"[In CreateDws] The protocol server MUST reply with an Error element in the CreateDwsResponse response message if it fails to create the specified Document Workspace.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R682");
            
            // Verify MS-DWSS requirement: MS-DWSS_R682
            this.Site.CaptureRequirementIfIsNotNull(
                error,
                682,
                @"[In CreateDws] The protocol server MUST return a CreateDwsResponse response message with an Error element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R458");
            
            // Verify MS-DWSS requirement: MS-DWSS_R458
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                458,
                @"[In CreateDws] If this is non-empty and another site (2) with the same name already exists on the site (2) on which the workspace is being created, the protocol server MUST return a ServerFailure error code (see section 2.2.3.2).");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R485");
            
            // Verify MS-DWSS requirement: MS-DWSS_R485
            this.Site.CaptureRequirementIfAreEqual<string>(
                "1",
                error.ID,
                485,
                @"[In CreateDwsResponse] ServerFailure, Identifier is ""1"", means: Protocol server encountered an error during the attempt to create the workspace.");
            
            if (Common.IsRequirementEnabled(1683, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1683");
                
                // Verify MS-DWSS requirement: MS-DWSS_R1683
                this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                    ErrorTypes.ServerFailure,
                    error.Value,
                    1683,
                    @"[In Appendix B: Product Behavior] Implementation does return error ServerFailure when a site (2) with the specified name already exists on this site (2). (<4> When a site (2) with the specified name already exists on this site (2), Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 will return error ServerFailure instead of AlreadyExists.)");
            }
            
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }

        #endregion

        #region Test RenameDws Operation
        
        /// <summary>
        /// This test case is intended to validate the result node returned by RenameDws operation when using a valid title for the Document Workspace site.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC08_RenameDws_ValidTitle()
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
            
            string dwsNewTitle = Common.GetConfigurationPropertyValue("ValidTitle", Site) + "_" + Common.FormatCurrentDateTime();
            this.dwsAdapter.RenameDws(dwsNewTitle, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
            
            Results getDwsDataRespResults = this.dwsAdapter.GetDwsData(documents.Name, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R554");
            
            // Verify MS-DWSS requirement: MS-DWSS_R554
            this.Site.CaptureRequirementIfAreEqual<string>(
                dwsNewTitle,
                getDwsDataRespResults.Title,
                554,
                @"[In GetDwsDataResponse] Title: The title of the workspace.");
            
            // Rename Document Workspace without title.
            this.dwsAdapter.RenameDws(null, out error);
            this.Site.Assert.IsNotNull(error, "The server should return an error!");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by RenameDws operation when the user does not have sufficient access to rename the title for the Document Workspace site.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC09_RenameDws_NoAccess()
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

            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Set Dws service credential to Reader credential.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            string dwsNewTitle = Common.GetConfigurationPropertyValue("ValidTitle", Site) + "_" + Common.FormatCurrentDateTime();
            this.dwsAdapter.RenameDws(dwsNewTitle, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be a NoAccess error.");
            
            // The precondition has verified this requirement already.
            this.Site.CaptureRequirement(
                299,
                @"[In RenameDwsResponse] Error: This element is returned when an error occurs in processing.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R294");
            
            // Verify MS-DWSS requirement: MS-DWSS_R294
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.NoAccess,
                error.Value,
                294,
                @"[In RenameDws] If the user submitting the request is not authorized to change the title, the protocol server MUST return an Error element with a NoAccess code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R22");
            
            // Verify MS-DWSS requirement: MS-DWSS_R22
            this.Site.CaptureRequirementIfAreEqual<string>(
                "3",
                error.ID,
                22,
                @"[In Error] The value 3 [ID] matches the Error type NoAccess.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R295");
            
            // Verify MS-DWSS requirement: MS-DWSS_R295
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                295,
                @"[In RenameDws] The Error element MUST NOT contain an AccessUrl attribute.");
            
            // Set default Dws service credential to admin credential.
            userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            password = Common.GetConfigurationPropertyValue("Password", this.Site);
            domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by DeleteDws operation when successfully deleting the current Document Workspace site and its contents.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC10_DeleteDws_DeleteCurrentSiteSuccessfully()
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
            
            // Delete the created web site.
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The response should not contains an error.");
            
            // If the response isn't an error, it indicate that the server returns a Result element
            this.Site.CaptureRequirement(
                166,
                @"[In DeleteDws] If none of the prior conditions apply, the protocol server MUST delete the specified Document Workspace and return a Result element.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by DeleteDws operation when the target Document Workspace does not exist.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC11_DeleteDws_NotExist()
        {
            this.Site.Assume.IsTrue(
                Common.IsRequirementEnabled(1164, this.Site) || Common.IsRequirementEnabled(2164, this.Site) || Common.IsRequirementEnabled(3164, this.Site),
                "Test is executed only when R1164Enabled, R2164Enabled or R3164Enabled is set to true.");

            // Change the web service url to a non-existent site.
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("NonexistentDWSSWebsiteUrl", this.Site);

            Error error;

            try
            {
                this.dwsAdapter.DeleteDws(out error);

                // If no exception was caught, then the case failed.
                this.Site.Assert.Fail("The expected exception is not returned for DeleteDws when the specified Document Workspace does not exist.");
            }
            catch (WebException webException)
            {
                if (webException.Response == null)
                {
                    throw;
                }

                // The web exception response must not be null.
                this.Site.Assert.IsNotNull(webException.Response, "The server should return http status code 404, the response must not be null web exception. The actual exception is:{0}", webException.Message);

                Site.Log.Add(LogEntryKind.Comment, "The specified Document Workspace does not exist, catch the exception: {0}", webException.Message);

                HttpStatusCode statusCode = ((HttpWebResponse)webException.Response).StatusCode;

                if (Common.IsRequirementEnabled(1164, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1164");

                    // Verify MS-DWSS requirement: MS-DWSS_R1164
                    this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                        HttpStatusCode.NotFound,
                        statusCode,
                        1164,
                        @"[In Appendix B: Product Behavior] [If the specified Document Workspace does not exist,] Implementation does return HTTP status code 404 with response body which contains text ""404 FILE NOT FOUND"". (Microsoft SharePoint Foundation 2013 Preview follows this behavior.)");
                }

                if (Common.IsRequirementEnabled(2164, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R2164");

                    // Verify MS-DWSS requirement: MS-DWSS_R2164
                    this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                        HttpStatusCode.NotFound,
                        statusCode,
                        2164,
                        @"[In Appendix B: Product Behavior] [If the specified Document Workspace does not exist,] Implementation does return HTTP status code 404 with an empty response body. (<6>For WSS3, the text is empty)");
                }

                if (Common.IsRequirementEnabled(3164, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R3164");

                    // Verify MS-DWSS requirement: MS-DWSS_R3164
                    this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                        HttpStatusCode.NotFound,
                        statusCode,
                        3164,
                        @"[In Appendix B: Product Behavior] [If the specified Document Workspace does not exist,] Implementation does return HTTP status code 404 with response body which contains text ""404 NOT FOUND"". (<6>For wss4, the text is ""404 NOT FOUND"". )");
                }
            }
        }
        
        /// <summary>
        /// This test case is intended to validate the result node that contains WebContainsSubwebs error returned by DeleteDws operation.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC12_DeleteDws_WebContainsSubwebs()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDws1RespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the new created site.
            this.dwsAdapter.ServiceUrl = createDws1RespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // create a SubSite
            CreateDwsResultResults createDws2RespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsTrue(error.Value == ErrorTypes.WebContainsSubwebs, "The response is expected to be a WebContainsSubwebs error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R30");
            
            // Verify MS-DWSS requirement: MS-DWSS_R30
            this.Site.CaptureRequirementIfAreEqual<string>(
                "11",
                error.ID,
                30,
                @"[In Error] The value 11 [ID] matches the Error type WebContainsSubwebs.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R321");
            
            // Verify MS-DWSS requirement: MS-DWSS_R321
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                321,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains WebContainsSubwebs error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R163");
            
            // Verify MS-DWSS requirement: MS-DWSS_R163
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.WebContainsSubwebs,
                error.Value,
                163,
                @"[In DeleteDws] If the specified Document Workspace has sub sites, the protocol server MUST return an Error element with the WebContainsSubwebs error code.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1671");
            
            // Verify MS-DWSS requirement: MS-DWSS_R1671
            // If error is not null, it indicate that the server did return an error element with a correct schema.
            this.Site.CaptureRequirementIfIsNotNull(
                error,
                1671,
                @"[In DeleteDwsResponse] Error: An Error element as specified in section 2.2.3.2.");
            
            // Redirect the web service to the sub site.
            this.dwsAdapter.ServiceUrl = createDws2RespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
            
            // Redirect the web service to the parent site.
            this.dwsAdapter.ServiceUrl = createDws1RespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by DeleteDws operation when the user does not have sufficient access to delete the target Document Workspace site.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S01_TC13_DeleteDws_NoAccess()
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

            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Set Dws service credential to Reader credential which does not have sufficient rights.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R162");
            
            // Verify MS-DWSS requirement: MS-DWSS_R162
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.NoAccess,
                error.Value,
                162,
                @"[In DeleteDws] The protocol server MUST return an Error element with a NoAccess code if the authenticated user is not authorized to delete the Document Workspace.");
            
            // Set Dws service credential to a user who has sufficient permission to call DeleteDws operation.
            userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            password = Common.GetConfigurationPropertyValue("Password", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            // Redirect the web service to the site collection without sub sites.
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("SiteCollectionWithoutSubSite", this.Site);
            
            // Try to delete the site collection.
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R499");
            
            // Verify MS-DWSS requirement: MS-DWSS_R499
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                499,
                @"[In DeleteDws] If the specified Document Workspace is the root site of the site collection, the protocol server MUST return an Error element with the ServerFailure error code.");
            
            // Redirect the web service to the created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
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