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
    /// Provides test methods for validating the operation: RemoveDwsUser. 
    /// </summary>
    [TestClass]
    public class S05_ManageSiteUsers : TestClassBase
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

        #region Test RemoveDwsUser Operation
        
        /// <summary>
        /// This test case is intended to validate that the RemoveDwsUser operation removes the specified user from the list of users for the current Document Workspace site successfully.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S05_TC01_RemoveDwsUser_RemoveUserSuccessfully()
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
            
            // Get first member of the current site.
            Results getDwsDataRespResults = this.dwsAdapter.GetDwsData(documents.Name, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a GetDwsDataResult not an error");
            this.Site.Assert.IsNotNull(getDwsDataRespResults.Members.Items, "The server should return a member element.");
            this.Site.Assert.IsTrue(getDwsDataRespResults.Members.Items.Length > 0, "The site members should be more than one.");
            
            Member firstMember = getDwsDataRespResults.Members.Items[0] as Member;
            this.Site.Assert.IsNotNull(firstMember, "The user should exist on server.");
            
            // Remove the first member.
            this.dwsAdapter.RemoveDwsUser(int.Parse(firstMember.ID), out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a Result element not an error");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The response should not be an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the related requirements when the server returns a ServerFailure Error element during processing RemoveDwsUser operation.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S05_TC02_RemoveDwsUser_ServerFailure()
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
            
            int invalidUserId = -1;
            this.dwsAdapter.RemoveDwsUser(invalidUserId, out error);
            this.Site.Assert.IsNotNull(error, "The expected response is an error.");
            
            // If the error is not null, it indicates that the server returned an Error element as is specified.
            this.Site.CaptureRequirement(
                289,
                @"[In RemoveDwsUserResponse] Error: An Error element as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R286");
            
            // Verify MS-DWSS requirement: MS-DWSS_R286
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                286,
                @"[In RemoveDwsUser] If an error of any type occurs during the processing, the protocol server MUST return an Error element as specified in section 2.2.3.2 with an error code of ServerFailure.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R20");
            
            // Verify MS-DWSS requirement: MS-DWSS_R20
            this.Site.CaptureRequirementIfAreEqual<string>(
                "1",
                error.ID,
                20,
                @"[In Error] The value 1 [ID] matches the Error type ServerFailure.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R312");
            
            // Verify MS-DWSS requirement: MS-DWSS_R312
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                312,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains ServerFailure error code.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The response should not be an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with NoAccess code should be returned by RemoveDwsUser operation if the user does not have sufficient access permissions to remove the user.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S05_TC03_RemoveDwsUser_NoAccess()
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
            
            Results getDwsDataRespResults = this.dwsAdapter.GetDwsData(documents.Name, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a GetDwsDataResult not an error");
            this.Site.Assert.IsNotNull(getDwsDataRespResults.Members.Items, "The members not expected to be null");
            
            // Get first member
            Member firstMember = getDwsDataRespResults.Members.Items[0] as Member;
            this.Site.Assert.IsNotNull(firstMember, "The user should exist on server.");
            
            // Set Dws service credential to Reader credential.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            this.dwsAdapter.RemoveDwsUser(int.Parse(firstMember.ID), out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be a NoAccess error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R22");
            
            // Verify MS-DWSS requirement: MS-DWSS_R22
            bool isVerifiedR22 = string.Equals("3", error.ID, StringComparison.CurrentCultureIgnoreCase) &&
                                 error.Value == ErrorTypes.NoAccess;
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR22,
                22,
                @"[In Error] The value 3 [ID] matches the Error type NoAccess.");
            
            // Set default Dws service credential to admin credential.
            userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            password = Common.GetConfigurationPropertyValue("Password", this.Site);
            domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The response should not be an error!");
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