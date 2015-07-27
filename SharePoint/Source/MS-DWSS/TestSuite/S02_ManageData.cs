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
    using System.Collections.Generic;
    using System.Net;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Provide test methods for validating the operation: GetDwsData, GetDwsMetaData and UpdateDwsData. 
    /// </summary>
    [TestClass]
    public class S02_ManageData : TestClassBase
    {
        #region Variables
        
        /// <summary>
        /// Adapter Instance.
        /// </summary>
        private IMS_DWSSAdapter dwsAdapter;
        
        /// <summary>
        /// The instance of the SUT control adapter instance. 
        /// </summary>
        private IMS_DWSSSUTControlAdapter sutControlAdapterInstance;

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

        #region Test GetDwsData Operation
        
        /// <summary>
        /// This test case is intended to validate that if the list in the Document Workspace has not changed since lastUpdate, GetDwsData must return a NoChanges child element of the List element as specified in GetDwsDataResponse. 
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC01_GetDwsData_NoChanges()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);

            // Set Dws service credential to Reader credential.
            string userName = Common.GetConfigurationPropertyValue("ReaderRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("ReaderRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);

            Error error;
            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            
            Results getDwsDataResult1 = this.dwsAdapter.GetDwsData(docUrl, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The server should return a GetDwsDataResult, not an error!");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R200");
            
            // Verify MS-DWSS requirement: MS-DWSS_R200
            this.Site.CaptureRequirementIfIsNotNull(
                getDwsDataResult1,
                200,
                @"[In GetDwsData] If no Error elements are returned as previously described, the protocol server MUST return a Result element with the appropriate information for the Document Workspace and document.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R729");
            
            // Verify MS-DWSS requirement: MS-DWSS_R729
            this.Site.CaptureRequirementIfIsNotNull(
                getDwsDataResult1,
                729,
                @"[In GetDwsData] If the protocol client provides an empty string, the protocol server MUST provide [all] data for the specified context.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R555");
            
            // Verify MS-DWSS requirement: MS-DWSS_R555
            long lastUpdate;
            bool isVerifiedR555 = long.TryParse(getDwsDataResult1.LastUpdate, out lastUpdate);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR555,
                555,
                @"[In GetDwsDataResponse] LastUpdate: An integer indicating the last time that the workspace was updated.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R42");
            
            // Verify MS-DWSS requirement: MS-DWSS_R42
            this.Site.Assert.IsNotNull(getDwsDataResult1.List1.Items, "The server should return the Document list element.");
            this.Site.Assert.IsTrue(getDwsDataResult1.List1.Items.Length > 0, "The server should not return an empty list element.");
            this.Site.Assert.IsFalse(getDwsDataResult1.List1.Items[0] is Error, "The first element in Documents list is expected to be an ID element, it is an error actually");
            
            ID listId = getDwsDataResult1.List1.Items[0] as ID;
            Guid listGuid;
            bool isVerifiedR42 = Guid.TryParse(listId.Value, out listGuid);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR42,
                42,
                @"[In List] ID: Contains a GUID that corresponds to a document on the server.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R556");
            
            // Verify MS-DWSS requirement: MS-DWSS_R556
            // The user's login name is prefixed with token like "i:0#.w|" if the server use claim based authentication.
            string actualName = getDwsDataResult1.User.LoginName.Substring(getDwsDataResult1.User.LoginName.LastIndexOf('|') + 1);
            string expectName = domain.Split('.')[0] + "\\" + userName;
            bool isVerifiedR556 = getDwsDataResult1.User != null && string.Equals(expectName, actualName, StringComparison.CurrentCultureIgnoreCase);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR556,
                556,
                @"[In GetDwsDataResponse] User: The member (1) information for the user who requested the GetDwsData operation.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R211");
            
            // Verify MS-DWSS requirement: MS-DWSS_R211
            // If the User is not null , it indicates that the IsDomainGroup and Email is present.
            this.Site.CaptureRequirementIfIsNotNull(
                getDwsDataResult1.User,
                211,
                @"[In GetDwsDataResponse] The content of this element MUST be a Member element as specified in section 2.2.3.5, with the exception that both IsDomainGroup and Email MUST be present.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R679");
            
            // Verify MS-DWSS requirement: MS-DWSS_R679
            // If the User is not null , it indicates that the IsSiteAdmin element is present.
            this.Site.CaptureRequirementIfIsNotNull(
                getDwsDataResult1.User,
                679,
                @"[In GetDwsDataResponse] In addition, the IsSiteAdmin element MUST be present.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R680");
            
            // Verify MS-DWSS requirement: MS-DWSS_R680
            bool isSiteAdmin;
            bool isVerifiedR680 = bool.TryParse(getDwsDataResult1.User.IsSiteAdmin.ToString(), out isSiteAdmin);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR680,
                680,
                @"[In GetDwsDataResponse] In addition, the IsSiteAdmin element MUST contain a Boolean value.");

            this.Site.Assert.IsTrue(getDwsDataResult1.Members.Items.Length > 0, "The member is expected to be more than one.");
            
            bool isIdUnique = true;
            foreach (Member member in getDwsDataResult1.Members.Items)
            {
                Dictionary<int, string> idloginNameList = new Dictionary<int, string>();
                if (idloginNameList.ContainsKey(int.Parse(member.ID)))
                {
                    isIdUnique = false;
                }
                else
                {
                    idloginNameList.Add(int.Parse(member.ID), member.LoginName);
                }
                
                // If the current member is a domain group
                if (string.Equals("MSDWSS_CustomGroup", member.Name, StringComparison.OrdinalIgnoreCase))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R371");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R371
                    this.Site.CaptureRequirementIfAreEqual<MemberIsDomainGroup>(
                        MemberIsDomainGroup.True,
                        member.IsDomainGroup,
                        371,
                        @"[In Member] When its value [IsDomainGroup] is ""True"", it [This member] is a group.");
                    
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R67");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R67
                    this.Site.CaptureRequirementIfAreEqual(
                        string.Empty,
                        member.Email,
                        67,
                        @"[In Member] If IsDomainGroup is set to True, this field [E-mail] MUST be empty.");
                    
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R66");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R66
                    this.Site.CaptureRequirementIfAreEqual(
                        string.Empty,
                        member.LoginName,
                        66,
                        @"[In Member] If IsDomainGroup is set to True, this field [LoginName] MUST be empty.");
                    
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R700");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R700
                    this.Site.CaptureRequirementIfIsTrue(
                        isIdUnique,
                        700,
                        @"[In Member] ID: A positive integer that MUST uniquely identify a group (2).");
                }
                else if (member.IsDomainGroup == MemberIsDomainGroup.False)
                {
                    // If the current member is a user
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R372");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R372
                    this.Site.CaptureRequirementIfIsTrue(
                        !string.IsNullOrEmpty(member.LoginName),
                        372,
                        @"[In Member] When its value [IsDomainGroup] is ""False"", it [This member] is a user.");
                    
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R65");
                    
                    // Verify MS-DWSS requirement: MS-DWSS_R65
                    this.Site.CaptureRequirementIfIsTrue(
                        isIdUnique,
                        65,
                        @"[In Member] ID: A positive integer that MUST uniquely identify a user.");
                }
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R214");
            
            // Verify MS-DWSS requirement: MS-DWSS_R214
            bool isVerifiedR214 = getDwsDataResult1.List != null && getDwsDataResult1.List1 != null &&
                                  getDwsDataResult1.List2 != null && getDwsDataResult1.Assignees != null;
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR214,
                214,
                @"[In GetDwsDataResponse] The following elements MUST be present for the GetDwsData operation: [Assignees, List, List, and List].");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R217");
            
            // Verify MS-DWSS requirement: MS-DWSS_R217
            this.Site.CaptureRequirementIfAreEqual<ListType>(
                ListType.Tasks,
                getDwsDataResult1.List.Name,
                217,
                @"[In GetDwsDataResponse] The Name attribute MUST be set to ""Tasks"".");
            
            // If the list is not null, it indicates that the List element type is correct.
            this.Site.CaptureRequirement(
                218,
                @"[In GetDwsDataResponse] The type of the List element MUST be List.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R219");
            
            // Verify MS-DWSS requirement: MS-DWSS_R219
            this.Site.CaptureRequirementIfAreEqual<ListType>(
                ListType.Documents,
                getDwsDataResult1.List1.Name,
                219,
                @"[In GetDwsDataResponse] The Name attribute MUST be set to ""Documents"".");
            
            // If the list is not null, it indicates that the List element type is correct.
            this.Site.CaptureRequirement(
                560,
                @"[In GetDwsDataResponse] The type of the List element MUST be List.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R227");
            
            // Verify MS-DWSS requirement: MS-DWSS_R227
            this.Site.CaptureRequirementIfAreEqual<ListType>(
                ListType.Links,
                getDwsDataResult1.List2.Name,
                227,
                @"[In GetDwsDataResponse] The Name attribute MUST be set to ""Links"".");
            
            // If the list is not null, it indicates that the List element type is correct.
            this.Site.CaptureRequirement(
                228,
                @"[In GetDwsDataResponse] The type of the List element MUST be List.");
            
            // Get Document Workspace data, GetDwsDataLastUpdateType set to LastUpdate, and the LastUpdate value is the value returned in the result of the previous call to GetDwsData method.
            Results getDwsDataResult2 = this.dwsAdapter.GetDwsData(docUrl, getDwsDataResult1.LastUpdate, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a GetDwsDataResult, not an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R40");
            
            // Verify MS-DWSS requirement: MS-DWSS_R40
            // If the list1 is neither an error nor an ID, it is a NoChanges element, this has been verified in schema.
            this.Site.Assert.IsNotNull(getDwsDataResult2.List1, "The server should return a Documents list.");
            this.Site.Assert.IsTrue(getDwsDataResult2.List1.Items.Length > 0, "The Documents list should not be empty.");

            bool isVerifiedR40 = !(getDwsDataResult2.List1.Items[0] is Error) && !(getDwsDataResult2.List1.Items[0] is ID);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR40,
                40,
                @"[In List] NoChanges: This element MUST be returned when the protocol client has provided a LastUpdate parameter and the specified document has not changed since the value in LastUpdate.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R41");
            
            // Verify MS-DWSS requirement: MS-DWSS_R41
            bool isVerifiedR41 = string.IsNullOrEmpty(getDwsDataResult2.List1.Items[0].ToString());
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR41,
                41,
                @"[In List] This element [NoChanges] MUST be empty.");
            
            // The server does return an NoChange element, verified in R40 and R41.
            this.Site.CaptureRequirement(
                205,
                @"[In GetDwsData] If the list in the Document Workspace has not changed since lastUpdate, GetDwsData MUST return a NoChanges child element of the List element as specified in GetDwsDataResponse.");
            
            // Get Document Workspace data without document url and LastUpdate.
            this.dwsAdapter.GetDwsData(null, null, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with ListNotFound code should be returned if the server cannot locate the document. 
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC02_GetDwsData_ListNotFound()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string docUrl = "InvalidDocumentUrl";
            
            Results getDwsDataResult = this.dwsAdapter.GetDwsData(docUrl, string.Empty, out error);
            this.Site.Assert.IsNull(error, "The server should return a GetDwsDataResult, not an error!");
            
            this.Site.Assert.IsNotNull(getDwsDataResult.List1.Items, "The server should return an Documents list element");
            this.Site.Assert.IsTrue(getDwsDataResult.List1.Items.Length > 0, "The Document list element should not be empty.");
            this.Site.Assert.IsTrue(getDwsDataResult.List1.Items[0] is Error, "The element in documents list should be an error, actual:{0}.", getDwsDataResult.List1.Items[0].ToString());

            error = getDwsDataResult.List1.Items[0] as Error;
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R26");
            
            // Verify MS-DWSS requirement: MS-DWSS_R26
            this.Site.CaptureRequirementIfAreEqual<string>(
                "7",
                error.ID,
                26,
                @"[In Error] The value 7 [ID] matches the Error type ListNotFound.");
            
            if (Common.IsRequirementEnabled(688, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R688");
                
                // Verify MS-DWSS requirement: MS-DWSS_R688
                this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                    ErrorTypes.ListNotFound,
                    error.Value,
                    688,
                    @"[In Appendix B: Product Behavior] Implementation does return an Error element with the ListNotFound code. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R317");
            
            // Verify MS-DWSS requirement: MS-DWSS_R317
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                317,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains ListNotFound error code.");
            
            // If R26, R688 is verified, it indicates that the server does return an ListNotFound error.
            this.Site.CaptureRequirement(
                223,
                @"[In GetDwsDataResponse] If the URL specified in the GetDwsData request cannot be found, the protocol server MUST return an Error element with a code of ListNotFound.");
            
            // If R26, R688 is verified, it indicates that the server does return an error when list not found.
            this.Site.CaptureRequirement(
                705,
                @"[In MemberData] Error MUST be used when there is a ListNotFound issue.");
        }
        
        /// <summary>
        /// This test case is intended to validate that an Error element with NoAccess code should be returned if the server detects an access restriction during processing.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC03_GetDwsData_NoAccess()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Set Dws service credential to none credential.
            string userName = Common.GetConfigurationPropertyValue("NoneRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("NoneRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);

            Error error;
            
            this.dwsAdapter.GetDwsData(string.Empty, string.Empty, out error);
            this.Site.Assert.IsNotNull(error, "The server is expect to return an Error");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R703");
            
            // Verify MS-DWSS requirement: MS-DWSS_R703
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.NoAccess,
                error.Value,
                703,
                @"[In MemberData] Error MUST be used when there is an access issue.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R198");
            
            // Verify MS-DWSS requirement: MS-DWSS_R198
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(error.AccessUrl),
                198,
                @"[In GetDwsData] If the protocol server detects an access restriction during processing, it MUST return an Error with the NoAccess code and a URL for an authentication page.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R22");
            
            // Verify MS-DWSS requirement: MS-DWSS_R22
            this.Site.CaptureRequirementIfAreEqual<string>(
                "3",
                error.ID,
                22,
                @"[In Error] The value 3 [ID] matches the Error type NoAccess.");
        }

        #endregion

        #region Test GetDwsMetaData Operation
        
        /// <summary>
        /// This test case is intended to validate that an Error element with DocumentNotFound code should be returned if the server cannot find the document. 
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC04_GetDwsMetaData_DocumentNotFound()
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
            
            // Redirect the web service to the new created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);

            string folderUrl = Common.GetConfigurationPropertyValue("ValidFolderUrl", this.Site) + "_" + Common.FormatCurrentDateTime();
            
            // Create a folder in the web site.
            this.dwsAdapter.CreateFolder(folderUrl, out error);
            this.Site.Assert.IsNull(error, "The response is expected to be a <Result/>, not an error.");

            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            
            // Get the metaData of the DWS; however, no document exists under the subfolder in the web site.
            this.dwsAdapter.GetDwsMetaData(docUrl, documents.ID, false, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be an error");
            
            // If the error is not null, it indicates that the server returns an Error element as is specified.
            this.Site.CaptureRequirement(
                596,
                @"[In GetDwsMetaDataResponse] If the element is an Error, it[GetDwsMetaDataResult] MUST contain an error code as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R28");
            
            // Verify MS-DWSS requirement: MS-DWSS_R28
            this.Site.CaptureRequirementIfAreEqual<string>(
                "9",
                error.ID,
                28,
                @"[In Error] The value 9 [ID] matches the Error type DocumentNotFound.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R704");
            
            // Verify MS-DWSS requirement: MS-DWSS_R704
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.DocumentNotFound,
                error.Value,
                704,
                @"[In MemberData] Error MUST be used when there is a DocumentNotFound issue.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R319");
            
            // Verify MS-DWSS requirement: MS-DWSS_R319
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(error.AccessUrl),
                319,
                @"[In Error] This attribute [AccessUrl] MUST NOT be present when the Error element contains DocumentNotFound error code.");
            
            // If the above condition are met, the server does returns a DocumentNotFound error.
            this.Site.CaptureRequirement(
                597,
                @"[In GetDwsMetaDataResponse] It[GetDwsMetaDataResult] MUST contain one of the error codes from the following table: [NoAccess, DocumentNotFound, ServerFailure].");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the parameter minimal is true. 
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC05_GetDwsMetaData_MinimalIsTrue()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            string docId = Guid.NewGuid().ToString();
            
            GetDwsMetaDataResultTypeResults getDwsMetaDataResult = this.dwsAdapter.GetDwsMetaData(docUrl, docId, true, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7071");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7071
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.Schema,
                7071,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server won't return the Schema (Tasks) element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7081");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7081
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.Schema,
                7081,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server won't return the Schema (Documents) element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7091");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7091
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.Schema,
                7091,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server won't return the Schema (Links) element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R710");
            
            // Verify MS-DWSS requirement: MS-DWSS_R710
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.ListInfo,
                710,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server MUST NOT return them:
                    From the GetDwsMetaData Results element:
                       ListInfo (Tasks)");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R711");
            
            // Verify MS-DWSS requirement: MS-DWSS_R711
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.ListInfo,
                711,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server MUST NOT return them:
                    From the GetDwsMetaData Results element:
                        ListInfo (Documents)");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R712");
            
            // Verify MS-DWSS requirement: MS-DWSS_R712
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult.ListInfo,
                712,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server MUST NOT return them:
                    From the GetDwsMetaData Results element:
                        ListInfo (Links)");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R706");
            
            // Verify MS-DWSS requirement: MS-DWSS_R706
            this.Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(getDwsMetaDataResult.SubscribeUrl),
                706,
                @"[In GetDwsMetaData] If [Minimal is] true, the protocol server MUST NOT return them:
                    From the GetDwsMetaData Results element:
                     SubscribeUrl");
            
            // If MS-DWSS_R706 ~ MS-DWSS_R712 verified successfully, MS-DWSS_R5821 verified.
            this.Site.CaptureRequirement(
                215,
                @"[In GetDwsDataResponse] The elements [Assignees, List, List, and List] MUST NOT be present if this data is being returned by GetDwsMetaData and the minimal parameter is set to TRUE.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the parameter minimal is false.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC06_GetDwsMetaData_MinimalIsFalse()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;
            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            string docId = Guid.NewGuid().ToString();
            
            GetDwsMetaDataResultTypeResults getDwsMetaDataResult1 = this.dwsAdapter.GetDwsMetaData(docUrl, docId, false, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R602");
            
            // Verify MS-DWSS requirement: MS-DWSS_R602
            Uri settingUrl;
            bool isVerifiedR602 = Uri.TryCreate(getDwsMetaDataResult1.SettingUrl, UriKind.RelativeOrAbsolute, out settingUrl);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR602,
                602,
                @"[In GetDwsMetaDataResponse] SettingUrl: URI of a page that enables workspace settings to be modified.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R603");
            
            // Verify MS-DWSS requirement: MS-DWSS_R603
            Uri permsUrl;
            bool isVerifiedR603 = Uri.TryCreate(getDwsMetaDataResult1.PermsUrl, UriKind.RelativeOrAbsolute, out permsUrl);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR603,
                603,
                @"[In GetDwsMetaDataResponse] PermsUrl: URI of a page that enables the workspace permissions settings to be modified.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R604");
            
            // Verify MS-DWSS requirement: MS-DWSS_R604
            Uri userInfoUrl;
            bool isVerifiedR604 = Uri.TryCreate(getDwsMetaDataResult1.UserInfoUrl, UriKind.RelativeOrAbsolute, out userInfoUrl);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR604,
                604,
                @"[In GetDwsMetaDataResponse] UserInfoUrl: URI of a page that enables the list (1) of users to be modified.");
            
            this.Site.Assert.IsNotNull(getDwsMetaDataResult1.ListInfo, "The server should return ListInfo element.");
            this.Site.Assert.IsTrue(getDwsMetaDataResult1.ListInfo.Length == 3, "The server should return 3 ListInfo elements");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R587");
            
            // Verify MS-DWSS requirement: MS-DWSS_R587
            this.Site.CaptureRequirementIfAreEqual<string>(
                "Tasks",
                getDwsMetaDataResult1.ListInfo[0].Name,
                587,
                @"[In GetDwsMetaData] If [Minimal is] false, the protocol server MUST return the following elements:
                    From the GetDwsMetaData Results element:
                        ListInfo (Tasks)");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R588");
            
            // Verify MS-DWSS requirement: MS-DWSS_R588
            this.Site.CaptureRequirementIfAreEqual<string>(
                "Documents",
                getDwsMetaDataResult1.ListInfo[1].Name,
                588,
                @"[In GetDwsMetaData] If [Minimal is] false, the protocol server MUST return the following elements:
                    From the GetDwsMetaData Results element:
                        ListInfo (Documents)");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R589");
            
            // Verify MS-DWSS requirement: MS-DWSS_R589
            this.Site.CaptureRequirementIfAreEqual<string>(
                "Links",
                getDwsMetaDataResult1.ListInfo[2].Name,
                589,
                @"[In GetDwsMetaData] If [Minimal is] false, the protocol server MUST return the following elements:
                    From the GetDwsMetaData Results element:
                        ListInfo (Links)");
            
            // If R587, R588, R589 are verified, it indicates that the Name attribute does contains the name of the list.
            this.Site.CaptureRequirement(
                634,
                @"[In ListInfo] Name: Contains the name of the list.");

            this.Site.Assert.IsFalse(getDwsMetaDataResult1.ListInfo[1].Items[0] is Error && getDwsMetaDataResult1.ListInfo[1].Items.Length > 1, "The Document ListInfo should not be an Error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R628");
            
            // Verify MS-DWSS requirement: MS-DWSS_R628
            ListInfoListPermissions listPermissions = getDwsMetaDataResult1.ListInfo[1].Items[1] as ListInfoListPermissions;
            bool isVerifiedR628 = listPermissions.DeleteListItems != null &&
                                  listPermissions.EditListItems != null &&
                                  listPermissions.InsertListItems != null &&
                                  listPermissions.ManageLists != null;
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR628,
                628,
                @"[In ListInfo] The following elements MUST be present if, and only if, the user has that permission: InsertListIems, EditListItems, DeleteListItems, and ManageLists.");
            
            // If R628 is verified, it indicates that the ListPermissions element does displays the current user permissions.
            this.Site.CaptureRequirement(
                627,
                @"[In ListInfo] ListPermissions: This element displays the current user permissions that are associated with the list (1).");
            
            // If R628 is verified, it indicates that the user does have permission to add new items.
            this.Site.CaptureRequirement(
                629,
                @"[In ListInfo] InsertListItems: Specifies that the current user can add new items to the list.");
            
            // If R628 is verified, it indicates that the user does have permission to edit list items.
            this.Site.CaptureRequirement(
                630,
                @"[In ListInfo] EditListItems: Specifies that the current user can edit list items.");
            
            // If R628 is verified, it indicates that the user does have permission to delete list item.
            this.Site.CaptureRequirement(
                631,
                @"[In ListInfo] DeleteListItems: Specifies that the current user can delete list items.");
            
            // If R628 is verified, it indicates that the user does have permission to manage the list.
            this.Site.CaptureRequirement(
                632,
                @"[In ListInfo] ManageLists: Specifies that the current user can manage the list.");
            
            if (Common.IsRequirementEnabled(689, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R689");
                
                // Verify MS-DWSS requirement: MS-DWSS_R689
                this.Site.CaptureRequirementIfAreEqual(
                    string.Empty,
                    getDwsMetaDataResult1.MtgInstance,
                    689,
                    @"[In Appendix B: Product Behavior] Implementation does return a GetDwsMetaData response in which MtgInstance element value is an empty string. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R583");
            
            // Verify MS-DWSS requirement: MS-DWSS_R583
            this.Site.CaptureRequirementIfIsFalse(
                string.IsNullOrEmpty(getDwsMetaDataResult1.SubscribeUrl),
                583,
                @"[In GetDwsMetaData] If [Minimal is] false, the protocol server MUST return the following elements:
                From the GetDwsMetaData Results element:
                    SubscribeUrl");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R611");
            
            // Verify MS-DWSS requirement: MS-DWSS_R611
            this.Site.Assert.IsNotNull(getDwsMetaDataResult1.HasUniquePerm, "The server should return the HasUniquePerm element");
            XmlNode[] hasUniquePerm = (XmlNode[])getDwsMetaDataResult1.HasUniquePerm;
            this.Site.Assert.IsTrue(hasUniquePerm.Length > 0, "The HasUniquePerm value is expected to be False or True, it should not be empty.");
                
            this.Site.CaptureRequirementIfAreEqual<string>(
                "True",
                hasUniquePerm[0].Value,
                611,
                @"[In GetDwsMetaDataResponse] HasUniquePerm: Set to True if, and only if, the workspace has custom role assignments.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R270");
            
            // Verify MS-DWSS requirement: MS-DWSS_R270
            this.Site.Assert.IsNotNull(getDwsMetaDataResult1.WorkspaceType, "The server should return the WorkspaceType element");
            XmlNode[] workspaceType = (XmlNode[])getDwsMetaDataResult1.WorkspaceType;
            this.Site.Assert.IsTrue(workspaceType.Length > 0, "The WorkspaceType should have value.");
                                  
            bool isVerifiedR270 = string.Equals("DWS", workspaceType[0].Value, StringComparison.CurrentCulture) ||
                                  string.Equals("MWS", workspaceType[0].Value, StringComparison.CurrentCulture) ||
                                  string.IsNullOrEmpty(workspaceType[0].Value);
                
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR270,
                270,
                @"[In GetDwsMetaDataResponse] WorkspaceType: This value MUST be ""DWS"", ""MWS"", or an empty string.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R617");
            
            // Verify MS-DWSS requirement: MS-DWSS_R617
            this.Site.Assert.IsNotNull(getDwsMetaDataResult1.IsADMode, "The server should return an IsADMode element");
            XmlNode[] isADMode = (XmlNode[])getDwsMetaDataResult1.IsADMode;
            this.Site.Assert.IsTrue(isADMode.Length > 0, "The IsADMode value should be True or False, it should not be empty.");
                
            this.Site.CaptureRequirementIfAreEqual<string>(
                "False",
                isADMode[0].Value,
                617,
                @"[In GetDwsMetaDataResponse] IsADMode: Set to FALSE if, and only if, the workspace is not set to Active Directory mode, that is, a mode in which new site (2) members are not created in Active Directory Domain Services (AD DS).");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R281");
            
            // Verify MS-DWSS requirement: MS-DWSS_R281
            this.Site.Assert.IsNotNull(getDwsMetaDataResult1.Minimal, "The server should return an Minimal element.");
            XmlNode[] minimal = (XmlNode[])getDwsMetaDataResult1.Minimal;
            this.Site.Assert.IsTrue(minimal.Length > 0, "The Minimal value should be True or False, it should not be empty.");
                
            this.Site.CaptureRequirementIfAreEqual<string>(
                "False",
                minimal[0].Value,
                281,
                @"[In GetDwsMetaDataResponse] This value [Minimal] MUST match the value in the request.");
            
            // If R281 is verified, it indicates that the Minimal match the minimal flag from the GetDwsMetaData request.
            this.Site.CaptureRequirement(
                280,
                @"[In GetDwsMetaDataResponse] Minimal: This element contains the minimal flag from the GetDwsMetaData request.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R282");
            
            // Verify MS-DWSS requirement: MS-DWSS_R282
            // If item is not null, it indicates that the GetDwsDataResult element does contain Result element.
            this.Site.CaptureRequirementIfIsNotNull(
                getDwsMetaDataResult1.Results,
                282,
                @"[In GetDwsMetaDataResponse] GetDwsDataResult: This element is identical to the GetDwsDataResult element specified in section 3.1.4.7.2.2, with the exception that it MUST contain Results element.");
            
            // Create another workspace.
            UsersItem users = new UsersItem();
            DocumentsItem documents = new DocumentsItem();
            
            users.Name = Common.GetConfigurationPropertyValue("UserName", this.Site);
            users.Email = Common.GetConfigurationPropertyValue("RegisteredUsersEmail", this.Site);
            
            documents.ID = Guid.NewGuid().ToString();
            documents.Name = Common.GetConfigurationPropertyValue("DocumentsName", this.Site) + "_" + Common.FormatCurrentDateTime();

            CreateDwsResultResults createDwsRespResults = this.dwsAdapter.CreateDws(string.Empty, users, string.Empty, documents, out error);
            
            // Redirect the web service to the new created site.
            this.dwsAdapter.ServiceUrl = createDwsRespResults.Url + Common.GetConfigurationPropertyValue("TestDWSSSuffix", this.Site);
            
            // Retrieve the current SUT version
            string sutVersion = Common.GetConfigurationPropertyValue("SutVersion", Site);
            string baseUrl = createDwsRespResults.Url;
            
            // Add the "Tasks", "Documents" and "Links" list in the new created workspace.
            bool isAddTasksList = this.sutControlAdapterInstance.AddList("Tasks", 107, baseUrl);
            this.Site.Assert.IsTrue(isAddTasksList, "Failed to add the Tasks list.");
            
            bool isAddLinksList = this.sutControlAdapterInstance.AddList("Links", 103, baseUrl);
            this.Site.Assert.IsTrue(isAddLinksList, "Failed to add the Links list.");
            
            // The documents list name in MOSS15 and WSS15 is "Documents".
            if (string.Equals(sutVersion, "SharePointFoundation2013") ||
                string.Equals(sutVersion, "SharePointServer2013"))
            {
                bool isAddDocList = this.sutControlAdapterInstance.AddList("Documents", 101, baseUrl);
                this.Site.Assert.IsTrue(isAddDocList, "Failed to add the Documents list.");
            }
            else
            {
                bool isAddDocList = this.sutControlAdapterInstance.AddList("Shared Documents", 101, baseUrl);
                this.Site.Assert.IsTrue(isAddDocList, "Failed to add the Shared Documents list.");
            }
            
            GetDwsMetaDataResultTypeResults getDwsMetaDataResult2 = this.dwsAdapter.GetDwsMetaData(docUrl, docId, false, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");

            this.Site.Assert.IsNotNull(getDwsMetaDataResult2.Schema, "The server should return Schema elements.");
            
            bool isVerifiedR584 = false;
            bool isVerifiedR585 = false;
            bool isVerifiedR586 = false;
            
            foreach (Schema schema in getDwsMetaDataResult2.Schema)
            {
                // If the current schema name is "Tasks", R584 is verified.
                if (string.Equals("Tasks", schema.Name, StringComparison.CurrentCulture))
                {
                    isVerifiedR584 = true;
                    continue;
                }
                
                // If the current schema name is "Documents", R585 is verified.
                if (string.Equals("Documents", schema.Name, StringComparison.CurrentCulture))
                {
                    isVerifiedR585 = true;
                    continue;
                }
                
                // If the current schema name is "Links", R586 is verified.
                if (string.Equals("Links", schema.Name, StringComparison.CurrentCulture))
                {
                    isVerifiedR586 = true;
                    continue;
                }
            }
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R584");
            
            // Verify MS-DWSS requirement: MS-DWSS_R584
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR584,
                584,
                @"[In GetDwsMetaData] If [Minimal is] false and the workspace that document belongs to contains a Tasks list, the protocol server will return the Schema (Tasks) element[, and otherwise not].");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R585");
            
            // Verify MS-DWSS requirement: MS-DWSS_R585
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR585,
                585,
                @"[In GetDwsMetaData] If [Minimal is] false and the workspace that document belongs to contains a DocumentLibrary list, the protocol server will return the Schema (Documents) element[, and otherwise not].");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R586");
            
            // Verify MS-DWSS requirement: MS-DWSS_R586
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR586,
                586,
                @"[In GetDwsMetaData] If [Minimal is] false and the workspace that document belongs to contains a Links list, the protocol server will return the Schema (Links) element[, and otherwise not].");
            
            // Delete the "Tasks", "Documents" and "Links" list if there's any in the new created workspace.
            bool isDeleteTasksList = this.sutControlAdapterInstance.DeleteList("Tasks", baseUrl);
            this.Site.Assert.IsTrue(isDeleteTasksList, "Failed to delete the Tasks List.");
            
            bool isDeleteLinksList = this.sutControlAdapterInstance.DeleteList("Links", baseUrl);
            this.Site.Assert.IsTrue(isDeleteLinksList, "Failed to delete the Links List.");
            
            // The documents list name in MOSS15 and WSS15 is "Documents".
            if (string.Equals(sutVersion, "SharePointFoundation2013") ||
                string.Equals(sutVersion, "SharePointServer2013"))
            {
                bool isDeleteDocList = this.sutControlAdapterInstance.DeleteList("Documents", baseUrl);
                this.Site.Assert.IsTrue(isDeleteDocList, "Failed to delete the Documents List.");
            }
            else
            {
                bool isDeleteDocList = this.sutControlAdapterInstance.DeleteList("Shared Documents", baseUrl);
                this.Site.Assert.IsTrue(isDeleteDocList, "Failed to delete the Documents List.");
            }

            GetDwsMetaDataResultTypeResults getDwsMetaDataResult3 = this.dwsAdapter.GetDwsMetaData(docUrl, docId, false, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7072");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7072
            // If the Schema is null, it indicates that the server didn't return the Schema element.
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult3.Schema,
                7072,
                @"[In GetDwsMetaData] If [Minimal is] false and the the workspace that document belongs to doesn't contain a Tasks list, the protocol server won't return the Schema (Tasks) element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7082");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7082
            // If the Schema is null, it indicates that the server didn't return the Schema element.
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult3.Schema,
                7082,
                @"[In GetDwsMetaData] If [Minimal is] false and the workspace that document belongs to doesn't contain a DocumentLibrary list, the protocol server won't return the Schema (Documents) element.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R7092");
            
            // Verify MS-DWSS requirement: MS-DWSS_R7092
            // If the Schema is null, it indicates that the server didn't return the Schema element.
            this.Site.CaptureRequirementIfIsNull(
                getDwsMetaDataResult3.Schema,
                7092,
                @"[In GetDwsMetaData] If [Minimal is] false and the workspace that document belongs to doesn't contain a Links list, the protocol server won't return the Schema (Links) element.");
            
            this.dwsAdapter.DeleteDws(out error);
            this.Site.Assert.IsNull(error, "The server should not return an error!");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the inherited site is created.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC07_GetDwsMetaData_InheritPermission()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("InheritPermissionSite", this.Site);
            
            Error error;
            GetDwsMetaDataResultTypeResults getDwsMetaDataResult = this.dwsAdapter.GetDwsMetaData(string.Empty, string.Empty, true, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R612");
            
            // Verify MS-DWSS requirement: MS-DWSS_R612
            this.Site.Assert.IsNotNull(getDwsMetaDataResult.HasUniquePerm, "The server should return the HasUniquePerm element.");
            XmlNode[] hasUniquePerm = (XmlNode[])getDwsMetaDataResult.HasUniquePerm;
            this.Site.Assert.IsTrue(hasUniquePerm.Length > 0, "The HasUniquePerm value should be True or False, it should not be empty.");
                
            this.Site.CaptureRequirementIfAreEqual<string>(
                "False",
                hasUniquePerm[0].Value,
                612,
                @"[In GetDwsMetaDataResponse] HasUniquePerm: Set to False, if role assignments are inherited from the site in which the workspace is created.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the invalid document URL is request.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC08_GetDwsMetaData_InvalidDocumentUrl()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            // Get the metaData of Document Workspace with invalid document URL.
            Error error;
            string docUrl = "InvalidDocumentUrl";
            string docId = Guid.NewGuid().ToString();
            
            this.dwsAdapter.GetDwsMetaData(docUrl, docId, true, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the site is the site collection, whose workspace type returned is not DWS or MWS.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC09_GetDwsMetaData_WorkspaceTypeEmpty()
        {
            // Set the service url to the site collection which is not a DWS or MWS.
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("SiteCollection", this.Site);
            
            Error error;
            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            string docId = Guid.NewGuid().ToString();
            
            // Get the metaData of Document Workspace with isMinimal set to true.
            GetDwsMetaDataResultTypeResults getDwsMetaDataResult = this.dwsAdapter.GetDwsMetaData(docUrl, docId, true, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R614");
            
            // Verify MS-DWSS requirement: MS-DWSS_R614
            // If the WorkspaceType is not XmlNode array type, it indicates that the server returns an empty string.
            this.Site.CaptureRequirementIfIsFalse(
                getDwsMetaDataResult.WorkspaceType is XmlNode[],
                614,
                @"[In GetDwsMetaDataResponse] If the site(2) is not one of those types [""DWS"", ""MWS"", or an empty string.], an empty string MUST be returned.");
        }
        
        /// <summary>
        /// This test case is intended to validate the result node returned by GetDwsMetaData operation when the parameter id is empty.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC10_GetDwsMetaData_IDIsEmpty()
        {
            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);
            
            Error error;

            if (Common.IsRequirementEnabled(1278, this.Site))
            {
                // Get the metaData of Document Workspace with isMinimal to be false and set id to empty.
                string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
                GetDwsMetaDataResultTypeResults getDwsMetaDataResult1 = this.dwsAdapter.GetDwsMetaData(docUrl, string.Empty, true, out error);
                this.Site.Assert.IsNull(error, "The server should not return an error.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1278");
                
                // Verify MS-DWSS requirement: MS-DWSS_R1278
                this.Site.Assert.IsNotNull(getDwsMetaDataResult1.DocUrl, "The server should return an docUrl element.");
                XmlNode[] docUrlNode = (XmlNode[])getDwsMetaDataResult1.DocUrl;
                this.Site.Assert.IsTrue(docUrlNode.Length > 0, "The docUrl value should be the value of the document parameter, and it should not be empty.");
                    
                this.Site.CaptureRequirementIfIsFalse(
                    string.IsNullOrEmpty(docUrlNode[0].Value),
                    1278,
                    @"[In Appendix B: Product Behavior] Implementation does set the DocUrl to the value of document from the GetDwsMetaData request if the value of document is specified. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }
            
            if (Common.IsRequirementEnabled(1279, this.Site))
            {
                // Get the meta data of Document Workspace without document and set id to empty, set isMinimal to false.
                GetDwsMetaDataResultTypeResults getDwsMetaDataResult2 = this.dwsAdapter.GetDwsMetaData(null, string.Empty, false, out error);
                this.Site.Assert.IsNull(error, "The server should not return an error.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1279");
                
                // Verify MS-DWSS requirement: MS-DWSS_R1279
                // If the DocUrl is not an XmlNode array, it indicates that the server returns an empty DocUrl.
                this.Site.CaptureRequirementIfIsFalse(
                    getDwsMetaDataResult2.DocUrl is XmlNode[],
                    1279,
                    @"[In Appendix B: Product Behavior] Implementation does set the DocUrl to be empty if the value of document is not specified in the GetDwsMetaData request. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }

            // Get the meta data of Document Workspace without document url and id, set isMinimal to false.
            this.dwsAdapter.GetDwsMetaData(null, null, false, out error);
            this.Site.Assert.IsNull(error, "The server should not return an error.");
        }

        /// <summary>
        /// This test case is intended to validate the result node that contains HTTP Error returned by GetDwsMetaData operation when the authenticated user is not permitted to access this information.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC11_GetDwsMetaData_HTTPError()
        {
            this.Site.Assume.IsTrue(
                Common.IsRequirementEnabled(1680, this.Site) || Common.IsRequirementEnabled(1682, this.Site),
                "Test is executed only when R1680Enabled or R1682Enabled is set to true.");

            this.dwsAdapter.ServiceUrl = Common.GetConfigurationPropertyValue("TestDWSSWebSite", this.Site);

            // Set Dws service credential to NoneRoleUser credential.
            string userName = Common.GetConfigurationPropertyValue("NoneRoleUser", this.Site);
            string password = Common.GetConfigurationPropertyValue("NoneRoleUserPassword", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.dwsAdapter.Credentials = new NetworkCredential(userName, password, domain);

            Error error;
            string docUrl = Common.GetConfigurationPropertyValue("ValidDocumentUrl", this.Site);
            string docId = Guid.NewGuid().ToString();

            try
            {
                this.dwsAdapter.GetDwsMetaData(docUrl, docId, false, out error);

                this.Site.Assert.Fail("The expected HTTP status code 401 is not returned for GetDwsMetaData when the authenticated user is not permitted to access this information.");
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
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R1680");

                // Verify MS-DWSS requirement: MS-DWSS_R1680
                this.Site.CaptureRequirementIfAreEqual<HttpStatusCode>(
                    HttpStatusCode.Unauthorized,
                    ((HttpWebResponse)webException.Response).StatusCode,
                    1680,
                    @"[In Appendix B: Product Behavior] Implementation does return HTTP status code 401 with response body which contains text ""401 Unauthorized"" instead of the ""NoAccess"" error. (<8>Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 will never return ""NoAccess"" error code and will return HTTP status code 401 with response body which contains text ""401 Unauthorized"".)");

                // If R1680 is verified, then the server did return HTTP status code 401 when the user is unauthorized.
                this.Site.CaptureRequirement(
                    1682, 
                    @"[In Appendix B: Product Behavior] Implementation does return HTTP status code 401 with response body which contains text ""401 Unauthorized"" instead of the ""NoAccess"" error. (<10>Windows SharePoint Services 3.0, SharePoint Foundation 2010 and SharePoint Foundation 2013 will never return ""NoAccess"" error code and will return HTTP status code 401 with response body which contains text ""401 Unauthorized"".)");
            }
        }

        #endregion

        #region Test UpdateDwsData Operation
        
        /// <summary>
        /// This test case is intended to validate that an Error element with ServerFailure code should be returned when the UpdateDwsData operation is failed.
        /// </summary>
        [TestCategory("MSDWSS"), TestMethod()]
        public void MSDWSS_S02_TC12_UpdateDwsData_ServerFailure()
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
            
            // Update Dws data with updates and meetingInstance set to empty string.
            this.dwsAdapter.UpdateDwsData(string.Empty, string.Empty, out error);
            this.Site.Assert.IsNotNull(error, "The response is expected to be a ServerFailure error.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R303");
            
            // Verify MS-DWSS requirement: MS-DWSS_R303
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                303,
                @"[In UpdateDwsData] If there is a failure during processing of this operation, the protocol server MUST return a ServerFailure error, as specified in section 2.2.3.2.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-DWSS_R673");
            
            // Verify MS-DWSS requirement: MS-DWSS_R673
            this.Site.CaptureRequirementIfAreEqual<ErrorTypes>(
                ErrorTypes.ServerFailure,
                error.Value,
                673,
                @"[In UpdateDwsDataResponse] If there is a failure during processing of the operation, , it MUST contain a ServerFailure error, as specified in section 2.2.3.2.");
            
            // Update Document Workspace data without updates and meetingInstance.
            this.dwsAdapter.UpdateDwsData(null, null, out error);
            this.Site.Assert.IsNotNull(error, "The server should return an error.");
            
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
            this.sutControlAdapterInstance = Site.GetAdapter<IMS_DWSSSUTControlAdapter>();
            
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