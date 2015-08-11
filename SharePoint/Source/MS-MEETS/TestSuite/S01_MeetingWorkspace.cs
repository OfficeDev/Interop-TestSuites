namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Scenario 1 Test cases. Test meeting workspace related operations and requirements,
    /// include adding meeting workspace, setting workspace’s title, getting workspaces information and deleting the workspace.
    /// </summary>
    [TestClass]
    public class S01_MeetingWorkspace : TestClassBase
    {
        #region Variables
        /// <summary>
        /// An instance of IMEETSAdapter.
        /// </summary>
        private IMS_MEETSAdapter meetsAdapter;

        /// <summary>
        /// An instance of IMS_MEETSSUTControlAdapter.
        /// </summary>
        private IMS_MEETSSUTControlAdapter sutControlAdapter;
        #endregion

        #region Test suite initialization and cleanup
        /// <summary>
        /// Initialize the test suite.
        /// </summary>
        /// <param name="context">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            // Setup test site.
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Reset the test environment.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            // Cleanup test site, must be called to ensure closing of logs.
            TestClassBase.Cleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// This test case is used to test typical workspace scenario.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC01_WorkspaceOperations()
        {
            // Check whether creating workspace is supported and query available languages.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate | MeetingInfoTypes.QueryLanguages, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");

            // If AllowCreate is returned and its value is "true" or "false", MS-MEETS_R181, MS-MEETS_R194 can be verified.
            bool isR181R194Verified = getWorkspaceInfoResult.Result.MeetingsInformation.AllowCreate.ToLower(CultureInfo.CurrentCulture) == "true" || getWorkspaceInfoResult.Result.MeetingsInformation.AllowCreate.ToLower(CultureInfo.CurrentCulture) == "false";
            Site.CaptureRequirementIfIsTrue(
                isR181R194Verified,
                181,
                @"[In GetMeetingsInformation]If the value 0x1 is set, the request is to query whether the user has permission to create meeting workspaces on this Web site (2) [the operation is sent to].");

            Site.CaptureRequirementIfIsTrue(
                isR181R194Verified,
                194,
                @"[In GetMeetingsInformationResponse]This element [AllowCreate]is present in the response when bit flag 0x1 is specified in requestFlags.");

            // If ListTemplateLanguages is not null, that is to say, this element is present in the response, then MS-MEETS_R199 can be verified.
            Site.CaptureRequirementIfIsNotNull(
               getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages,
               199,
               @"[In GetMeetingsInformationResponse]This element [ListTemplateLanguages]is present in the response when bit flag 0x2 is specified inrequestFlags.");

            // If the operation is executed successfully and ListTemplateLanguages is not null, the returned ListTemplateLanguages represent the languages supported, then MS-MEETS_R182 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages,
                182,
                @"[In GetMeetingsInformation]If the value 0x2 is set, the request is to query for the site template languages supported.");
                  
            string lcidString = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages[0];
            uint lcid = 0;
            Site.Assert.IsTrue(uint.TryParse(lcidString, out lcid), "LCID must be integer");

            // Get available workspace templates.
            getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryTemplates, lcid);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");

            // If the operation is executed successfully and ListTemplates is not null, the returned ListTemplates represent the templates supported, then MS-MEETS_R183 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates,
                183,
                @"[In GetMeetingsInformation]If the value 0x4 is set, the request is to query for the list of site templates supported.");

            // If ListTemplates is not null, that is to say, this element is present in the response, then MS-MEETS_R203 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates,
                203,
                @"[In GetMeetingsInformationResponse]This element [ListTemplates]is present in the response when bit flag 0x4 is specified in requestFlags.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "If the length of ListTemplates is greater than 0, MS-MEETS_R202, MS-MEETS_R204 can be verified.");

            // If the length of ListTemplates is greater than 0, it is a list for site template and every single template is available template, MS-MEETS_R202, MS-MEETS_R204 can be verified.
            bool isVerifiedR202R204 = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates.Length > 0;
            
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR202R204,
                202,
                @"[In GetMeetingsInformationResponse]ListTemplates: The list of site templates supported.");

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR202R204,
                204,
                @"[In GetMeetingsInformationResponse]Template: The available site template.");

            // Create a new workspace using the template from the step above.
            Template template = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates[0];
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, template.Name, lcid, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Get workspaces on the site. Make sure the workspace has been created successfully.
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getWorkspaceResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assert.AreEqual<int>(1, getWorkspaceResult.Result.MeetingWorkspaces.Length, "There should be only 1 workspace.");
            Site.Assert.AreEqual<string>(workspaceTitle, getWorkspaceResult.Result.MeetingWorkspaces[0].Title, "Workspace title should be the same as the request value.");
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            
            // Get workspace status. Make sure there is no meetings.
            getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.AreEqual<string>("0", getWorkspaceInfoResult.Result.MeetingsInformation.WorkspaceStatus.MeetingCount, "There should be no meetings.");

            // If AllowAuthenticatedUsers is returned and its value is "true" or "false", MS-MEETS_R217 can be verified.
            bool isVerifiedR217 = getWorkspaceInfoResult.Result.MeetingsInformation.WorkspaceStatus.AllowAuthenticatedUsers.ToLower(CultureInfo.CurrentCulture) == "true" || getWorkspaceInfoResult.Result.MeetingsInformation.WorkspaceStatus.AllowAuthenticatedUsers.ToLower(CultureInfo.CurrentCulture) == "false";
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR217,
                217,
                @"[In GetMeetingsInformationResponse]AllowAuthenticatedUsers: Specifies whether access to a meeting workspace subsite has been granted to authenticated users.");

            // AllowAuthenticatedUsers is the status of a workspace, if MS-MEETS_R217 is verified, MS-MEETS_R211 can also be verified. 
            Site.CaptureRequirement(
                211,
                @"[In GetMeetingsInformationResponse]WorkspaceStatus: The status of a workspace.");
            
            // Update the workspace title.
            string newWorkspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<Null> setWorkspaceTitleResult = this.meetsAdapter.SetWorkspaceTitle(newWorkspaceTitle);
            Site.Assert.IsNull(setWorkspaceTitleResult.Exception, "SetWorkspaceTitle should succeed");

            // Get workspaces on the site. Make sure the workspace title has been updated.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            getWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getWorkspaceResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assert.AreEqual<string>(newWorkspaceTitle, getWorkspaceResult.Result.MeetingWorkspaces[0].Title, "Workspace title should set to the new value");

            // Delete the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteWorkspaceResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteWorkspaceResult.Exception, "DeleteWorkspace should succeed");

            // Get no workspaces to indicate the workspace has been deleted.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            getWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getWorkspaceResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assert.AreEqual<int>(0, getWorkspaceResult.Result.MeetingWorkspaces.Length, "The workspace should be deleted.");
        }

        /// <summary>
        /// This test case verifies that the CreateWorkspace operation cannot create a meeting workspace as a sub site of another meeting workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC02_CreateWorkspaceOnWorkspaceError()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);

            // Send CreateWorkspace to a Web site that is a meeting workspace.
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createNewWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);

            // If createAnotherWorkspaceResult.Exception is not null, it means create workspace under existed workspace is failed, MS-MEETS_R3005 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                createNewWorkspaceResult.Exception,
                3005,
                @"[In CreateWorkspaceResponse]This operation [CreateWorkspace]cannot create a meeting workspace as a subsite of another meeting workspace.");
        
            // If the response contains the SOAP fault code "0x00000001", MS-MEETS_R113 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000001",
                createNewWorkspaceResult.GetErrorCode(),
                113,
                @"[In CreateWorkspaceResponse]If this operation[CreateWorkspace] is sent to a Web site (2) that is a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000001"".");

            // The SOAP fault code is the error information which Detail element contains. MS-MEETS_R9 can be verified directly.
            Site.CaptureRequirement(
                9,
                @"[In Detail]The Detail element contains error information for a SOAP fault.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the error codes about workspace when the service Url is not a meeting workspace Url.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC03_WorkspaceInvalidUrlError()
        {
            // Send SetWorkspaceTitle to a web site that is not a meeting workspace.
            string newWorkspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<Null> setWorkspaceTitleResult = this.meetsAdapter.SetWorkspaceTitle(newWorkspaceTitle);

            // If the response contains the SOAP fault code "0x00000006", MS-MEETS_R320 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000006",
                setWorkspaceTitleResult.GetErrorCode(),
                320,
                @"[In SetWorkspaceTitleResponse]If this operation[SetWorkspaceTitle] is sent to a Web site (2) that is not a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000006"".");

            // Send DeleteWorkspace to a web site that is not a meeting workspace.
            SoapResult<Null> deleteWorkspaceResult = this.meetsAdapter.DeleteWorkspace();

            // If the exception is not null, it means DeleteWorkspace operation fail, MS-MEETS_R155 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                deleteWorkspaceResult.Exception,
                155,
                @"[In DeleteWorkspace]This operation [DeleteWorkspace] fails if sent to a Web site (2) that is not a meeting workspace.");

            // If the response contains the SOAP fault code "0x00000004", MS-MEETS_R166 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000004",
                deleteWorkspaceResult.GetErrorCode(),
                166,
                @"[In DeleteWorkspaceResponse]If this operation[DeleteWorkspace] is sent to a Web site (2) that is not a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000004"".");
        }

        /// <summary>
        /// This test case verifies that, when the bit flag 0x1 is specified in requestFlags, server returns a SOAP fault if sent the GetMeetingsInformation operation to a web site that is a meeting workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC04_AllowCreateOnWorkspaceError()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Send GetMeetingsInformation with requestFlags bit set to 0x1 to a web site that is a meeting workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getMeetingsInformastionResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate, null);

            // If the response contains the SOAP fault code "0x00000001", MS-MEETS_R196 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000001",
                getMeetingsInformastionResult.GetErrorCode(),
                196,
                @"[In GetMeetingsInformationResponse]If the operation [GetMeetingsInformation]is sent to a Web site (2) that is a meeting workspace, the response MUST be a SOAP fault with SOAP fault code ""0x00000001"".");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case verifies that if the title in CreateWorkspace is longer than 255 characters, the title will be truncated.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC05_CreateWorkspaceWithLongTitle()
        {
            // Create a workspace with long title.
            string longTitle = TestSuiteBase.GenerateRandomString(257);
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(longTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspacesResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getMeetingWorkspacesResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assume.AreEqual<int>(1, getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length, "There should be only 1 workspaces.");

            // Create workspace with the title exceed 255 characters, if the returned workspace title less than or equal 255 characters, MS-MEETS_R104 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                longTitle.Substring(0, 255),
                getMeetingWorkspacesResult.Result.MeetingWorkspaces[0].Title,
                104,
                @"[In CreateWorkspace][if title is larger than 255 characters]Remaining characters [of title]are truncated.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the behavior about SetWorkspaceTitle with title absent.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC06_SetWorkspaceTitleNotSpecified()
        {
            // Create a workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Set workspace title to null.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> setWorkSpaceTitleResponse = this.meetsAdapter.SetWorkspaceTitle(null);
            Site.Assert.IsNull(setWorkSpaceTitleResponse.Exception, "SetWorkspaceTitle should succeed");

            // Get workspaces on the site.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspacesResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getMeetingWorkspacesResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assume.AreEqual<int>(1, getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length, "There should be only 1 workspaces.");

            // Set empty title, if the returned new title is empty, MS-MEETS_R318 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                string.Empty,
                getMeetingWorkspacesResult.Result.MeetingWorkspaces[0].Title,
                318,
                @"[In SetWorkspaceTitle]If this parameter [title]is absent, the new title is the empty string.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the URL of new created workspace.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC07_VerifyUrlOfNewCreatedWorkspace()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Get the Url of new created workspace.
            string urlCreate = createWorkspaceResult.Result.CreateWorkspace.Url;
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            string urlGet = getMeetingWorkspaceResult.Result.MeetingWorkspaces[0].Url;

            // If the GetMeetingWorkspace's Url is equals to the CreateWorkspace's Url, MS-MEETS_R112 and MS-MEETS_R236 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                urlCreate,
                urlGet,
                112,
                @"[In CreateWorkspaceResponse]Url: The URL of the newly created meeting workspace.");
        
            Site.CaptureRequirementIfAreEqual<string>(
                urlCreate,
                urlGet,
                236,
                @"[In GetMeetingWorkspacesResponse]Url: The URL of the meeting workspace.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the elements in GetMeetingWorkspacesResponse.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC08_GetMeetingWorkspacesResponse()
        {
            // Create a workspace on test site.
            string workspaceTitleCreate = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitleCreate, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace succeed");

            // Get workspace title. 
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getWorkspaceResult.Exception, "Get meeting workspace succeed");
            string workspaceTitleGet = getWorkspaceResult.Result.MeetingWorkspaces[0].Title;

            // If the GetMeetingWorkspace's title is equals to the CreateWorkspace's title, MS-MEETS_R237 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                workspaceTitleCreate,
                workspaceTitleGet,
                237,
                @"[In GetMeetingWorkspacesResponse]Title: The title of the meeting workspace.");

            // Create another workspace on test site.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createMeetingWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace succeed");

            // Get workspace information.
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspaceResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getMeetingWorkspaceResult.Exception, "Get meeting workspace succeed");

            // If the returned MeetingWorkspaces is a 2 length list, then MS-MEETS_R234 can be verified.
            Site.CaptureRequirementIfAreEqual<int>(
                2,
                getMeetingWorkspaceResult.Result.MeetingWorkspaces.Length,
                234,
                @"[In GetMeetingWorkspacesResponse]MeetingWorkspaces: A list of meeting workspaces.");

            // 2 individual workspaces are returned, then MS-MEETS_R235 can also be verified.
            Site.CaptureRequirement(
                235,
                @"[In GetMeetingWorkspacesResponse]Workspace: An individual meeting workspace.");

            // Delete the created workspaces.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
            
            this.meetsAdapter.Url = createMeetingWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteMeetingResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteMeetingResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the WorkspaceStatus when requestFlags set to 0x8.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC09_VerifyWorkspaceStatusWhenRequestFlagIs0x8()
        {
            // Create a workspace on test site.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Query other status values of the workspace created above.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceStatusResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(getWorkspaceStatusResult.Exception, "Get meeting information should succeed. ");

            // If UniquePermissions is returned and its value is "true" or "false", MS-MEETS_R213 can be verified.
            bool isVerifiedR213 = getWorkspaceStatusResult.Result.MeetingsInformation.WorkspaceStatus.UniquePermissions.ToLower(CultureInfo.CurrentCulture) == "true" || getWorkspaceStatusResult.Result.MeetingsInformation.WorkspaceStatus.UniquePermissions.ToLower(CultureInfo.CurrentCulture) == "false";
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR213,
                213,
                @"[In GetMeetingsInformationResponse]UniquePermissions: Specifies whether the meeting workspace subsite contains unique role assignments.");

            // If AllowAuthenticatedUsers is returned and its value is "true" or "false", MS-MEETS_R217 can be verified.
            bool isVerifiedR217 = getWorkspaceStatusResult.Result.MeetingsInformation.WorkspaceStatus.AllowAuthenticatedUsers.ToLower(CultureInfo.CurrentCulture) == "true" || getWorkspaceStatusResult.Result.MeetingsInformation.WorkspaceStatus.AllowAuthenticatedUsers.ToLower(CultureInfo.CurrentCulture) == "false";
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR217,
                217,
                @"[In GetMeetingsInformationResponse]AllowAuthenticatedUsers: Specifies whether access to a meeting workspace subsite has been granted to authenticated users.");

            // If the workspace status is returned successfully and the UniquePermissions and AllowAuthenticatedUsers are all workspace status, then MS-MEETS_R184 can be verified also.
            Site.CaptureRequirement(
                184,
                @"[In GetMeetingsInformation]If the value 0x8 is set, the request is to query other status values of a workspace.");

            // If the workspace status is returned successfully and the UniquePermissions and AllowAuthenticatedUsers are all workspace status, then MS-MEETS_R211 can be verified also.
            Site.CaptureRequirement(
                211,
                @"[In GetMeetingsInformationResponse]WorkspaceStatus: The status of a workspace.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify the behavior about SetWorkspaceTitle with long title.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC10_SetWorkspaceTitleWithLongTitle()
        {
            // Create a workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Set workspace title with 255 characters string.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string validTitle = TestSuiteBase.GenerateRandomString(255);
            SoapResult<Null> setWorkSpaceTitleResponse = this.meetsAdapter.SetWorkspaceTitle(validTitle);
            Site.Assert.IsNull(setWorkSpaceTitleResponse.Exception, "SetWorkspaceTitle should succeed");

            // Get workspaces on the site.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspaces = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getMeetingWorkspaces.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assume.AreEqual<int>(1, getMeetingWorkspaces.Result.MeetingWorkspaces.Length, "There should be only 1 workspaces.");

            // Set workspace title with more than 255 characters string.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string invalidTitle = TestSuiteBase.GenerateRandomString(256);
            this.meetsAdapter.SetWorkspaceTitle(invalidTitle);
           
            // Get workspaces on the site.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> getMeetingWorkspacesResult = this.meetsAdapter.GetMeetingWorkspaces(null);
            Site.Assert.IsNull(getMeetingWorkspacesResult.Exception, "GetMeetingWorkspaces should succeed");
            Site.Assume.AreEqual<int>(1, getMeetingWorkspacesResult.Result.MeetingWorkspaces.Length, "There should be only 1 workspaces.");

            // If the given length of meeting workspace title exceeds 255, the returned title length equals to 255, all characters after 255 are truncated, MS-MEETS_R316 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                invalidTitle.Substring(0, 255),
                getMeetingWorkspacesResult.Result.MeetingWorkspaces[0].Title,
                316,
                @"[In SetWorkspaceTitle]This string [title]has a maximum length of 255 characters. ");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to verify get meeting information with valid and invalid URL.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC11_GetMeetingInformationByAllowCreate()
        {
            // Create a workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Send GetMeetingsInformation with requestFlags bit set to 0x1 to a web site that it is the workspace's parent site.
            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.IsNotNull(getWorkspaceInfoResult.Result.MeetingsInformation.AllowCreate.ToLower(CultureInfo.CurrentCulture), "AllowCreate should be not null");

            // Send GetMeetingsInformation with requestFlags bit set to 0x1 to a web site that is a meeting workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoInvalidUrl = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate, null);
            Site.Assert.IsNotNull(getWorkspaceInfoInvalidUrl.Exception, "GetMeetingsInformation should not succeed");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "If GetMeetingsInformation from workspace's parent site is successful and GetMeetingsInformation from meeting workspace is failed, MS-MEETS_R316 can be verified");

            // If GetMeetingsInformation from workspace's parent site is successful and GetMeetingsInformation from meeting workspace is not successful, it means the operation sent to a parent Web site, MS-MEETS_R195 can be verified.
            bool isVerifyR195 = getWorkspaceInfoResult.Exception == null && getWorkspaceInfoInvalidUrl.Exception != null;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR195,
                195,
                @"[In GetMeetingsInformationResponse]When bit flag 0x1 is specified, the operation [GetMeetingsInformation] MUST be sent to a parent Web site (2), as opposed to a meeting workspace subsite itself.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test GetMeetingsInformation operation for single meeting instance.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC12_GetMeetingInformationForSingleMeetingInstance()
        {
            // Create a new workspace.
            string workspaceTitle = TestSuiteBase.GetUniqueWorkspaceTitle();
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(workspaceTitle, null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Add two meetings in the workspace.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            string organizerEmail = Common.GetConfigurationPropertyValue("OrganizerEmail", this.Site);
            SoapResult<AddMeetingResponseAddMeetingResult> addMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addMeetingResult.Exception, "AddMeeting should succeed");

            SoapResult<AddMeetingResponseAddMeetingResult> addAnotherMeetingResult = this.meetsAdapter.AddMeeting(organizerEmail, Guid.NewGuid().ToString(), null, DateTime.Now, TestSuiteBase.GetUniqueMeetingTitle(), TestSuiteBase.GetUniqueMeetingLocation(), DateTime.Now, DateTime.Now.AddHours(1), null);
            Site.Assert.IsNull(addAnotherMeetingResult.Exception, "AddMeeting should succeed");

            // Send GetMeetingsInformation to get the number of single meeting instances.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> singleInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryOthers, null);
            Site.Assert.IsNull(singleInfoResult.Exception, "GetMeetingsInformation should succeed");

            // Add two single meeting instances on workspace, if the returned meeting count equals to 2, MS-MEETS_R214 can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "2",
                singleInfoResult.Result.MeetingsInformation.WorkspaceStatus.MeetingCount,
                214,
                @"[In GetMeetingsInformationResponse]MeetingCount: The number of single meeting instances on the subsite.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test create workspace with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC13_CreateWorkspaceWithAllParametersSpecified()
        {
            // Check whether creating workspace is supported and query available languages.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate | MeetingInfoTypes.QueryLanguages, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assume.AreEqual<string>("true", getWorkspaceInfoResult.Result.MeetingsInformation.AllowCreate.ToLower(CultureInfo.CurrentCulture), "Site must support creating workspace");
            Site.Assume.AreNotEqual<int>(0, getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages.Length, "ListTemplateLanguages must contain one or more items.");
            string lcidString = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages[0];
            uint lcid = 0;
            Site.Assert.IsTrue(uint.TryParse(lcidString, out lcid), "LCID must be integer");

            // Get available workspace templates.
            getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryTemplates, lcid);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.AreNotEqual<int>(0, getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates.Length, "ListTemplates must contain one or more items.");
            Template template = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates[0];

            // The templates returned in above are all server supported template, see the first template, template.Name is the short name of the template, MS-MEETS_R205 can be verified directly when it returns.
            Site.CaptureRequirementIfIsNotNull(
                template.Name,
                205,
                @"[In GetMeetingsInformationResponse]Name: The short name of the site template.");

            // The templates returned in above are all server supported template, see the first template, template.Title is the title of the template, MS-MEETS_R206 can be verified directly when it returns.
            Site.CaptureRequirementIfIsNotNull(
                template.Title,
                206,
                @"[In GetMeetingsInformationResponse]Title: The title of the site template.");

            // The templates returned in above are all server supported template, see the first template, template.Id is the identification number of the template, MS-MEETS_R207 can be verified directly when it returns.
            Site.CaptureRequirementIfIsNotNull(
                template.Id,
                207,
                @"[In GetMeetingsInformationResponse]Id: The identification number of the site template.");

             // The templates returned in above are all server supported template, see the first template, template.Description is the description of the template, MS-MEETS_R208 can be verified directly when it returns.
            Site.CaptureRequirementIfIsNotNull(
                template.Description,
                208,
                @"[In GetMeetingsInformationResponse]Description: A description of the site template.");

             // The templates returned in above are all server supported template, see the first template, template.ImageUrl is the URL of an image or icon of the template, MS-MEETS_R209 can be verified directly when it returns.
            Site.CaptureRequirementIfIsNotNull(
                template.ImageUrl,
                209,
                @"[In GetMeetingsInformationResponse]ImageUrl: The URL of an image or icon of the site template.");

            // Create a new workspace with all parameters specified.
            TimeZoneInf timeZoneInf = new TimeZoneInf();
            timeZoneInf.standardDate = new SysTime();
            timeZoneInf.standardDate.year = 2013;
            timeZoneInf.standardDate.month = 10;
            timeZoneInf.standardDate.dayOfWeek = 4;
            timeZoneInf.standardDate.day = 5;
            timeZoneInf.standardDate.hour = 23;
            timeZoneInf.standardDate.minute = 10;
            timeZoneInf.standardDate.second = 10;
            timeZoneInf.standardDate.milliseconds = 100;
            timeZoneInf.standardBias = 1;
            timeZoneInf.bias = 1;
            timeZoneInf.daylightBias = 2;
            timeZoneInf.daylightDate = new SysTime();
            timeZoneInf.standardDate.year = 0;
            timeZoneInf.standardDate.month = 10;
            timeZoneInf.standardDate.dayOfWeek = 4;
            timeZoneInf.standardDate.day = 5;
            timeZoneInf.standardDate.hour = 23;
            timeZoneInf.standardDate.minute = 10;
            timeZoneInf.standardDate.second = 10;
            timeZoneInf.standardDate.milliseconds = 100;
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(TestSuiteBase.GetUniqueWorkspaceTitle(), template.Name, lcid, timeZoneInf);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // In above CreateWorkspace step, the lcid value input is gained from getWorkspaceInfoResult, when CreateWorkspace succeed, MS-MEETS_R2935 can be verified. 
            Site.CaptureRequirement(
                2935,
                @"[In CreateWorkspace]If a value for this parameter [lcid] is provided, it MUST be a LCID contained in the GetMeetingsInformationResponse response after the client protocol sends the GetMeetingsInformation message with the requestFlags parameter set to 2.");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test create workspace without optional parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC14_CreateWorkspaceWithoutOptionalParameters()
        {
            // Create a new workspace without optional parameters specified.
            SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> createWorkspaceResult = this.meetsAdapter.CreateWorkspace(TestSuiteBase.GetUniqueWorkspaceTitle(), null, null, null);
            Site.Assert.IsNull(createWorkspaceResult.Exception, "Create workspace should succeed");

            // Clean up the SUT.
            this.meetsAdapter.Url = createWorkspaceResult.Result.CreateWorkspace.Url + Common.GetConfigurationPropertyValue("EntryUrl", this.Site);
            SoapResult<Null> deleteResult = this.meetsAdapter.DeleteWorkspace();
            Site.Assert.IsNull(deleteResult.Exception, "DeleteWorkspace should succeed");
        }

        /// <summary>
        /// This test case is used to test GetMeetingInformation with all parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC15_GetMeetingInformationWithAllParametersSpecified()
        {
            // Check whether creating workspace is supported and query available languages.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.AllowCreate | MeetingInfoTypes.QueryLanguages, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assume.AreEqual<string>("true", getWorkspaceInfoResult.Result.MeetingsInformation.AllowCreate.ToLower(CultureInfo.CurrentCulture), "Site must support creating workspace");
            Site.Assume.AreNotEqual<int>(0, getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages.Length, "ListTemplateLanguages must contain one or more items.");
            string lcidString = getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages[0];
            uint lcid = 0;
            Site.Assert.IsTrue(uint.TryParse(lcidString, out lcid), "LCID must be integer");

            getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryLanguages, lcid);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");

            // If GetMeetingInformation executed succeed and the ListTemplateLanguages is returned. All the returned values of ListTemplateLanguages are supported languages(LCID), MS-MEETS_R198, MS-MEETS_R200 can be verified.
            Site.CaptureRequirementIfIsNotNull(
                getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages,
                198,
                @"[In GetMeetingsInformationResponse]ListTemplateLanguages: Lists the site template languages supported.");

            Site.CaptureRequirementIfIsNotNull(
                getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplateLanguages[0],
                200,
                @"[In GetMeetingsInformationResponse]LCID: The LCID of the available site template.");

            // GetMeetingInformation with all parameters specified.
            getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(MeetingInfoTypes.QueryTemplates, lcid);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
            Site.Assert.AreNotEqual<int>(0, getWorkspaceInfoResult.Result.MeetingsInformation.ListTemplates.Length, "ListTemplates must contain one or more items.");
        }

        /// <summary>
        /// This test case is used to test GetMeetingInformation without optional parameters specified.
        /// </summary>
        [TestCategory("MSMEETS"), TestMethod()]
        public void MSMEETS_S01_TC16_GetMeetingInformationWithoutOptionalParameters()
        {
            // GetMeetingInformation without optional parameters specified.
            SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> getWorkspaceInfoResult = this.meetsAdapter.GetMeetingsInformation(null, null);
            Site.Assert.IsNull(getWorkspaceInfoResult.Exception, "GetMeetingsInformation should succeed");
        }

        #endregion

        #region Test case initialization and cleanup

        /// <summary>
        /// Test case initialize method.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.meetsAdapter = this.Site.GetAdapter<IMS_MEETSAdapter>();
            Common.CheckCommonProperties(this.Site, true);

            this.meetsAdapter.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);

            this.sutControlAdapter = Site.GetAdapter<IMS_MEETSSUTControlAdapter>();

            // Make sure the test environment is clean before test case run.
            bool isClean = this.sutControlAdapter.PrepareTestEnvironment(this.meetsAdapter.Url);
            this.Site.Assert.IsTrue(isClean, "The specified site should not have meeting workspaces.");

            // Initialize the TestSuiteBase
            TestSuiteBase.Initialize(this.Site);
        }

        /// <summary>
        /// Test case cleanup method.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            this.meetsAdapter.Reset();
        }
        #endregion 
    }
}