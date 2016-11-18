namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Linq;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WWSP S02 GetForItem.
    /// </summary>
    [TestClass]
    public class S02_GetForItem : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
           TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Class clean up
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region TestCase

        #region MSWWSP_S02_TC01_GetTemplatesForItem_TemplateData
        /// <summary>
        /// This test case is used to verify the element TemplateData when GetTemplatesForItem operation successful.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S02_TC01_GetTemplatesForItem_TemplateData()
        {
            // Upload a file.
            string docItemUrl = this.UploadFileToSut(DocLibraryName);

            // If there are any existing task for new uploaded file, this method will throw an exception. 
            this.VerifyTaskDataOfNewUploadFile(docItemUrl);

            // Call method GetTemplatesForItem to get a set of workflow associations for an existing document. 
            GetTemplatesForItemResponseGetTemplatesForItemResult getTemplatesResult = ProtocolAdapter.GetTemplatesForItem(docItemUrl);

            if (getTemplatesResult == null)
            {
                this.Site.Assert.Fail("GetTemplatesForItem operation is failed.");
            }

            this.Site.Assume.IsNotNull(getTemplatesResult.TemplateData.WorkflowTemplates, "Could not get the WorkflowTemplates.");
            this.Site.Assume.IsTrue(getTemplatesResult.TemplateData.WorkflowTemplates.Length >= 1, "The length of the WorkflowTemplates is 0");

            // If the GetTemplatesForItem operation is succeed, SUT return response then R150 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R150
            Site.CaptureRequirement(
                150,
                @"[In GetTemplatesForItem] The protocol client sends a GetTemplatesForItemSoapIn request message, and the protocol server responds with a GetTemplatesForItemSoapOut response message.");

            // If the GetTemplatesForItem operation is succeed, SUT return TemplateDatas element in response, then R31, R147, R153, R162 and R408 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R31
            Site.CaptureRequirement(
                31,
                @"[In Elements] TemplateData specifies a set of workflow associations.");

            // Verify MS-WWSP requirement: MS-WWSP_R147
            Site.CaptureRequirement(
                147,
                @"[In GetTemplatesForItem] This operation obtains a set of workflow associations for a new or existing document.");

            // Verify MS-WWSP requirement: MS-WWSP_R153
            Site.CaptureRequirement(
                153,
                @"[In Messages] GetTemplatesForItemSoapOut specifies the response to a request for a set of workflow associations for a new or existing document.");

            // Verify MS-WWSP requirement: MS-WWSP_R162
            Site.CaptureRequirement(
                162,
                @"[In Elements] GetTemplatesForItemResponse contains the response to a request for a set of workflow associations for a new or existing document.");

            // Verify MS-WWSP requirement: MS-WWSP_R408
            Site.CaptureRequirement(
                408,
                @"[In Message Processing Events and Sequencing Rules] GetTemplatesForItem obtains a set of workflow associations that can be started on a new or existing document in a specified list.");

            if (getTemplatesResult.TemplateData.Web.Title != null)
            {
                string currentWebTitle = SutController.GetCurrentWebTitle();
                this.Site.Assume.IsNotNull(currentWebTitle, "Could not get the current web title.");

                // If the value of the Web.Title is equal to the specify value, then R384 should be covered.
                // Verify MS-WWSP requirement: MS-WWSP_R384
                Site.CaptureRequirementIfAreEqual<string>(
                    currentWebTitle,
                    getTemplatesResult.TemplateData.Web.Title,
                    384,
                    @"[In TemplateData] Web.Title: The title of the site for this set of workflow associations.");
            }

            if (getTemplatesResult.TemplateData.Web.Url != null)
            {
                string currentWebUrl = SutController.GetCurrentWebUrl().ToString();
                this.Site.Assume.IsNotNull(currentWebUrl, "The current web URL is " + currentWebUrl);

                // If the value of the Web.URL is equal to the specify value, then R385 should be covered.        
                // Verify MS-WWSP requirement: MS-WWSP_R385
                Site.CaptureRequirementIfAreEqual<string>(
                    currentWebUrl,
                    getTemplatesResult.TemplateData.Web.Url,
                    385,
                    @"[In TemplateData] Web.Url: A site URL for this set of workflow associations.");
            }

            if (getTemplatesResult.TemplateData.List.Title != null)
            {
                // If the value of the List.Title is equal to the specify value, then R386 should be covered.
                // Verify MS-WWSP requirement: MS-WWSP_R386
                string titleUrlValueInResponse = getTemplatesResult.TemplateData.List.Title.ToLower();
                string currentDocLibraryNameValue = DocLibraryName.ToLower();
                Site.CaptureRequirementIfAreEqual<string>(
                    currentDocLibraryNameValue,
                    titleUrlValueInResponse,
                    386,
                    @"[In TemplateData] List.Title: The title of the list for this set of workflow associations.");
            }

            if (getTemplatesResult.TemplateData.List.Url != null)
            {
                string listUrl = SutController.GetListUrlByName(DocLibraryName);
                this.Site.Assume.IsNotNull(listUrl, "The value of the List Url is" + listUrl);

                // If the value of the List.URL is equal to the specify value, then R387 should be covered.
                // Verify MS-WWSP requirement: MS-WWSP_R387
                listUrl = listUrl.ToLower();
                string urlValueInresponse = getTemplatesResult.TemplateData.List.Url.ToLower();
                Site.CaptureRequirementIfAreEqual<string>(
                    listUrl,
                    urlValueInresponse,
                    387,
                    @"[In TemplateData] List.Url: A list URL for this set of workflow associations.");
            }

            TemplateDataWorkflowTemplate templateItem = this.GetTemplateItemByName(getTemplatesResult.TemplateData.WorkflowTemplates, WorkflowAssociationName);

            // If the value of the TemplateId is equal to the specify value, then R389 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R389
            this.Site.Assert.IsNotNull(
                              templateItem.WorkflowTemplateIdSet.TemplateId,
                              "Could not get the template id for workflow association[{0}]",
                              TestSuiteBase.WorkflowAssociationName);

            this.Site.CaptureRequirementIfAreEqual(
                new Guid(WorkflowAssociationId),
                templateItem.WorkflowTemplateIdSet.TemplateId,
                389,
                @"[In TemplateData] WorkflowTemplates.WorkflowTemplate.WorkflowTemplateIdSet.TemplateId: A GUID identifying this workflow association.");

            // Get workflowassociation name
            string baseId = SutController.GetBaseIdOfWorkFlowAssociation(DocLibraryName, WorkflowAssociationName);
            this.Site.Assume.IsNotNull(baseId, "The value of the BaseId id " + baseId);

            // If the value of the BaseId is equal to the specify value, then R390 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R390
            this.Site.Assert.IsNotNull(
                         templateItem.WorkflowTemplateIdSet.BaseId,
                         "Could not get the BaseId for workflow association[{0}]",
                         TestSuiteBase.WorkflowAssociationName);

            this.Site.CaptureRequirementIfAreEqual(
                        new Guid(baseId),
                        templateItem.WorkflowTemplateIdSet.BaseId,
                        390,
                        @"[In TemplateData] WorkflowTemplates.WorkflowTemplate.WorkflowTemplateIdSet.BaseId: A GUID identifying the workflow template upon which this workflow association is based.");

            // If the value of the Name is equal to the specify value, then R391 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R391
            this.Site.Assert.IsFalse(
                        string.IsNullOrEmpty(templateItem.Name),
                        "Could not get the Name for workflow association[{0}]",
                        TestSuiteBase.WorkflowAssociationName);

            string templateNameInResponse = templateItem.Name.ToLower();
            string expectedName = WorkflowAssociationName.ToLower();
            this.Site.CaptureRequirementIfAreEqual(
                expectedName,
                        templateNameInResponse,
                        391,
                        @"[In TemplateData] WorkflowTemplates.WorkflowTemplate.Name: The name of this workflow association.");

            // If the value of the Description is not null, then R392 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R392
            this.Site.CaptureRequirementIfIsNotNull(
                        getTemplatesResult.TemplateData.WorkflowTemplates[0].Description,
                        392,
                        @"[In TemplateData] WorkflowTemplates.WorkflowTemplate.Description: The description of this workflow association.");
        }
        #endregion

        #region MSWWSP_S02_TC02_GetToDosForItem_ToDoData
        /// <summary>
        /// This test case is used to verify GetToDosForItem operation to a get a set of workflow tasks for a document.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S02_TC02_GetToDosForItem_ToDoData()
        {
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);

            // If there are any existing task for new uploaded file, this method will throw an exception. 
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start a normal work flow
            this.StartATaskWithNewFile(uploadFileUrl, false);

            // Call method GetToDosForItem  to get a set of Workflow Tasks for a document. 
            GetToDosForItemResponseGetToDosForItemResult getTodosResult = ProtocolAdapter.GetToDosForItem(uploadFileUrl);

            if (getTodosResult == null || getTodosResult.ToDoData == null || getTodosResult.ToDoData.xml == null 
                || getTodosResult.ToDoData.xml.data == null || getTodosResult.ToDoData.xml.data.Any == null)
            {
                this.Site.Assert.Fail("GetToDosForItem operation is failed.");
            }

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // If the number of the workflow tasks is greater than or equal to 1, then R409, R173, R179, R188 should be covered.        
            // Verify MS-WWSP requirement: MS-WWSP_R409
            Site.CaptureRequirementIfIsTrue(
                getTodosResult.ToDoData.xml.data.Any.Length >= 1,
                409,
                @"[In Message Processing Events and Sequencing Rules] GetToDosForItem obtains a set of workflow tasks for an existing document.");

            // Verify MS-WWSP requirement: MS-WWSP_R173
            Site.CaptureRequirementIfIsTrue(
                getTodosResult.ToDoData.xml.data.Any.Length >= 1,
                173,
                @"[In GetToDosForItem] This operation obtains a set of workflow tasks for a document.");

            // Verify MS-WWSP requirement: MS-WWSP_R179
            Site.CaptureRequirementIfIsTrue(
                getTodosResult.ToDoData.xml.data.Any.Length >= 1,
                179,
                @"[In Messages] GetToDosForItemSoapOut specifies the response to a request for a set of workflow tasks for a document.");

            // Verify MS-WWSP requirement: MS-WWSP_R188
            Site.CaptureRequirementIfIsTrue(
                getTodosResult.ToDoData.xml.data.Any.Length >= 1,
                188,
                @"[In Elements] GetToDosForItemResponse contains the response to a request for a set of workflow tasks for a document.");

            // If the response from the GetToDosForItem operation is not null, then R176 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R176
            Site.CaptureRequirementIfIsNotNull(
                getTodosResult,
                176,
                @"[In GetToDosForItem] The protocol client sends a GetToDosForItemSoapIn request message, and the protocol server responds with a GetToDosForItemSoapOut response message.");

            // If the element of the ToDoData is not null, then R32 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R32
            Site.CaptureRequirementIfIsNotNull(
                getTodosResult.ToDoData,
                32,
                @"[In Elements] ToDoData specifies a set of workflow tasks.");
        }
        #endregion

        #region MSWWSP_S02_TC03_GetWorkflowDataForItem
        /// <summary>
        /// This test case is used to verify GetWorkflowDataForItem operation, workflow associations, workflow tasks, and workflows should be returned.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S02_TC03_GetWorkflowDataForItem()
        {
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);

            // Start a normal work flow
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, false);
            this.Site.Assert.IsNotNull(taskIdValue, "Starting a workflow task should succeed.");
           
            // Call method GetWorkflowDataForItem to query a set of workFlow associations, workFlow tasks, and workFlows for a document. 
            GetWorkflowDataForItemResponseGetWorkflowDataForItemResult workflowDataResult = ProtocolAdapter.GetWorkflowDataForItem(uploadFileUrl);
            
            if (workflowDataResult == null || workflowDataResult.WorkflowData == null)
            {
                this.Site.Assume.Fail("GetWorkflowDataForItem operation is failed, the response is not be returned.");
            }

            this.Site.Assert.IsTrue(workflowDataResult.WorkflowData.DefaultWorkflows != null && 
                workflowDataResult.WorkflowData.DefaultWorkflows.DefaultWorkflow != null, "DefaultWorkflow should be present.");

            // Verify MS-WWSP requirement: MS-WWSP_R290
            Site.CaptureRequirementIfAreEqual<string>(
                "OnCheckInMajor",
                workflowDataResult.WorkflowData.DefaultWorkflows.DefaultWorkflow.Event,
                290,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.DefaultWorkflows.DefaultWorkflow.Event: MUST be set to ""OnCheckInMajor"".");

            // Verify MS-WWSP requirement: MS-WWSP_R198
            // If the response from the GetWorkflowDataForItem operation is not null, then R198 should be covered.
            Site.CaptureRequirementIfIsNotNull(
                workflowDataResult.WorkflowData,
                198,
                @"[In GetWorkflowDataForItem] The protocol client sends a GetWorkflowDataForItemSoapIn request message, and the protocol server responds with a GetWorkflowDataForItemSoapOut response message.");

            // If the element of the ActiveWorkflowsData is not null, then R220 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R220
            Site.CaptureRequirementIfIsNotNull(
                workflowDataResult.WorkflowData.ActiveWorkflowsData,
                220,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData: A set of workflows running on the document.");

            if (workflowDataResult.WorkflowData.ToDoData == null || workflowDataResult.WorkflowData.ToDoData.xml.data.Any.Length == 0)
            {
                this.Site.Assume.Fail("GetWorkflowDataForItem operation is failed, the element ToDoData is not be returned.");
            }

            if (workflowDataResult.WorkflowData.TemplateData == null)
            {
                this.Site.Assume.Fail("GetWorkflowDataForItem operation is failed, the element TemplateData is not be returned.");
            }

            // If the element of the StatusPageUrl is not null, then R221 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R221
            Site.CaptureRequirementIfIsNotNull(
                workflowDataResult.WorkflowData.ActiveWorkflowsData.Workflows[0].StatusPageUrl,
                221,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData.Workflows.Workflow.StatusPageUrl: The URL of a Web page that can show the status of a workflow.");

            // If the element ToDoData, TemplateData and ActiveWorkflowsData are all not null, then R410 should be covered.
            Site.CaptureRequirement(
                410,
                @"[In Message Processing Events and Sequencing Rules] GetWorkflowDataForItem obtains an aggregated set of workflows, workflow associations, and workflow tasks for an existing document.");

            // If the element ToDoData, TemplateData and ActiveWorkflowsData are all not null, then R195 should be covered.
            Site.CaptureRequirement(
                195,
                @"[In GetWorkflowDataForItem] This operation[GetWorkflowDataForItem] queries a set of workflow associations, workflow tasks, and workflows for a document.");

            // If the element ToDoData, TemplateData and ActiveWorkflowsData are all not null, then R201 should be covered.
            Site.CaptureRequirement(
                201,
                @"[In Messages] GetWorkflowDataForItemSoapOut specifies the response to a request to query a set of workflow associations, workflow tasks, and workflows for a document.");

            // If the element ToDoData, TemplateData and ActiveWorkflowsData are all not null, then R210 should be covered.
            Site.CaptureRequirement(
                210,
                @"[In Elements] GetWorkflowDataForItemResponse contains the response to a request to query a set of workflow associations, workflow tasks, and workflows for a document.");

            // If the TemplateId is equal to the specific value, then R224 and R225 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R224
            Site.CaptureRequirementIfAreEqual(
                new Guid(WorkflowAssociationId),
                workflowDataResult.WorkflowData.ActiveWorkflowsData.Workflows[0].TemplateId,
                224,
                @"[In GetWorkflowDataForItemResponse] GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData. Workflows.Workflow.TemplateId: A GUID identifying the workflow association of a workflow.");
        
            // Verify MS-WWSP requirement: MS-WWSP_R225
            Site.CaptureRequirementIfAreEqual(
                new Guid(WorkflowAssociationId),
                workflowDataResult.WorkflowData.ActiveWorkflowsData.Workflows[0].TemplateId,
                225,
                @"[In GetWorkflowDataForItemResponse] This[GetWorkflowDataForItemResult.WorkflowData.ActiveWorkflowsData. Workflows.Workflow.TemplateId] MUST be the workflow association of the workflow specified by Workflow.Id.");
        }
        #endregion

        #region MSWWSP_S02_TC04_GetWorkflowTaskData_Success
        /// <summary>
        /// This test case is used to verify GetWorkflowTaskData operation, retrieve data about a single workflow task successfully.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S02_TC04_GetWorkflowTaskData_Success()
        {
            // Upload a file. 
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start a normal work flow 
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, false);

            // Verify whether the task is assign to expected user group for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // Call method GetWorkflowTaskData to retrieve data about a single workflow task. 
            GetWorkflowTaskDataResponseGetWorkflowTaskDataResult askDataResult = ProtocolAdapter.GetWorkflowTaskData(uploadFileUrl, int.Parse(taskIdValue), new Guid(TaskListId));

            // If the response from the GetWorkflowTaskData is not null, then R295 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R295
            Site.CaptureRequirementIfIsNotNull(
                askDataResult,
                295,
                @"[In GetWorkflowTaskData] The protocol client sends a GetWorkflowTaskDataSoapIn request message, and the protocol server responds with a GetWorkflowTaskDataSoapOut response message.");

            // For a new upload file, start only one task currently, if the GetWorkflowTaskData operation return only one record of the task, then R411, R292, R196, R298 and R307 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R411
            Site.CaptureRequirementIfAreEqual(
                1,
                askDataResult.Any.Length,
                411,
                @"[In Message Processing Events and Sequencing Rules] GetWorkflowTaskData retrieves data about a single workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R292
            Site.CaptureRequirementIfAreEqual(
                1,
                askDataResult.Any.Length,
                292,
                @"[In GetWorkflowTaskData] This operation retrieves data about a single workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R196
            Site.CaptureRequirementIfAreEqual(
                1,
                askDataResult.Any.Length,
                196,
                @"[In GetWorkflowDataForItem] This operation[GetWorkflowDataForItem] retrieves some of the same data that GetToDosForItem and GetTemplatesForItem retrieve, as well as additional data.");

            // Verify MS-WWSP requirement: MS-WWSP_R298
            Site.CaptureRequirementIfAreEqual(
                1,
                askDataResult.Any.Length,
                298,
                @"[In Messages] GetWorkflowTaskDataSoapOut specifies the response to a request to retrieve data about a single workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R307
            Site.CaptureRequirementIfAreEqual(
                1,
                askDataResult.Any.Length,
                307,
                @"[In Elements] GetWorkflowTaskDataResponse contains the response to a request to retrieve data about a single workflow task.");
        }
        #endregion

        #region MSWWSP_S02_TC05_GetWorkflowTaskData_IgnoreItem
        /// <summary>
        /// This test case is used to verify if set the different string as the item value, server reply same when the site of the SOAP request URL contains a list with the specified ListId
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S02_TC05_GetWorkflowTaskData_IgnoreItem()
        {
            // Upload a file. 
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start a normal work flow 
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, false);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // Call method GetWorkflowTaskData to retrieve data about a single workflow task. 
            GetWorkflowTaskDataResponseGetWorkflowTaskDataResult askDataResult = ProtocolAdapter.GetWorkflowTaskData(uploadFileUrl, int.Parse(taskIdValue), new Guid(TaskListId));

            if (askDataResult == null || askDataResult.Any.Length == 0)
            {
                this.Site.Assume.Fail("GetWorkflowTaskData operation is failed, the response is null.");
            }

            // Initialize an invalid file URL.
            string notExistFileUrl = this.GenerateRandomValue().ToString();

            GetWorkflowTaskDataResponseGetWorkflowTaskDataResult askDataResultInvalieURL = ProtocolAdapter.GetWorkflowTaskData(notExistFileUrl, int.Parse(taskIdValue), new Guid(TaskListId));
            if (askDataResultInvalieURL == null || askDataResultInvalieURL.Any.Length == 0)
            {
                this.Site.Assume.Fail("GetWorkflowTaskData operation is failed when the file URL is not existing.");
            }

            // Call method GetWorkflowTaskData to retrieve data about a single workflow task, with an invalid URL.
            // If the server replay the same, then R313 should be covered.
            // Verify MS-WWSP requirement: MS-WWSP_R313
            Site.CaptureRequirementIfAreEqual(
                askDataResult.Any.Length,
                askDataResultInvalieURL.Any.Length,
                313,
                @"[In GetWorkflowTaskData] Set the different string as the item value, server reply same if the site (2) of the SOAP request URL contains a list with the specified ListId.");
        }
        #endregion

        #endregion

        #region private method

        /// <summary>
        /// Get the template item from TemplateDataWorkflowTemplate array by specified name.
        /// </summary>
        /// <param name="templates">A parameter represents the TemplateDataWorkflowTemplate array which the method perform on.</param>
        /// <param name="templateName">A parameter represent the name of the expected template item.</param>
        /// <returns>A return value represents the template item matched the specified name.</returns>
        private TemplateDataWorkflowTemplate GetTemplateItemByName(TemplateDataWorkflowTemplate[] templates, string templateName)
        {
            if (null == templates)
            {
                throw new ArgumentNullException("templates");
            }

            if (0 == templates.Length)
            {
                throw new ArgumentException("The templates' collection should contain at least one template item.", "templates");
            }

            if (string.IsNullOrEmpty(templateName))
            {
                throw new ArgumentNullException("templateName");
            }

            var templateItemMatchName = from templateItem in templates
                                        where templateItem.Name.Equals(WorkflowAssociationName)
                                        select templateItem;

            this.Site.Assert.AreNotEqual<int>(
                                        0,
                                        templateItemMatchName.Count(),
                                        "The templates' collection should contain at least one template item with name[{0}]",
                                        templateName);

            TemplateDataWorkflowTemplate expectedTemplateItem = templateItemMatchName.ElementAt<TemplateDataWorkflowTemplate>(0);

            return expectedTemplateItem;
        }
 
        #endregion
    }
}