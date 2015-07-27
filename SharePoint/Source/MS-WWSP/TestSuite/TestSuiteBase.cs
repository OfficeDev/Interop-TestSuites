//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WWSP
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WWSP.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables

        /// <summary>
        /// Gets or sets IMS_WWSPAdapter instance
        /// </summary>
        protected static IMS_WWSPAdapter ProtocolAdapter { get; set; }

        /// <summary>
        /// Gets or sets ISUTControlAdapter instance
        /// </summary>
        protected static IMS_WWSPSUTControlAdapter SutController { get; set; }

        /// <summary>
        /// Gets or sets the list collection used to record the task id which the task is started.
        /// </summary>
        protected static List<string> StartedTaskIdsRecorder { get; set; }

        /// <summary>
        /// Gets or sets the list collection which contains current URLs for uploaded files.
        /// </summary>
        protected static List<string> UploadedFilesUrlRecorder { get; set; }

        /// <summary>
        /// Gets or sets the current name of document library.
        /// </summary>
        protected static string DocLibraryName { get; set; }

        /// <summary>
        /// Gets or sets the current Id of the document library.
        /// </summary>
        protected static string DocListId { get; set; }

        /// <summary>
        ///  Gets or sets the Id of task list which is used by current workflow association.
        /// </summary>
        protected static string TaskListId { get; set; }

        /// <summary>
        /// Gets or sets the name of task list which is used by current workflow association.
        /// </summary>
        protected static string TaskListName { get; set; }

        /// <summary>
        /// Gets or sets the Id of workflow association which is applied on the current document library. 
        /// </summary>
        protected static string WorkflowAssociationId { get; set; }

        /// <summary>
        /// Gets or sets the name of workflow association which is applied on the current document library. 
        /// </summary>
        protected static string WorkflowAssociationName { get; set; }

        /// <summary>
        /// Gets or sets the name of Account who perform protocol's methods.
        /// </summary>
        protected static string CurrentProtocolPerformAccountName { get; set; }

        #endregion Variables

        #region Test Suite Initialization and clean up

        /// <summary>
        /// Initialize the variable for the test suite.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);

            if (null == StartedTaskIdsRecorder)
            {
                StartedTaskIdsRecorder = new List<string>();
            }

            if (null == UploadedFilesUrlRecorder)
            {
                UploadedFilesUrlRecorder = new List<string>();
            }

            if (null == ProtocolAdapter)
            {
                ProtocolAdapter = BaseTestSite.GetAdapter<IMS_WWSPAdapter>();
            }

            if (null == SutController)
            {
                SutController = BaseTestSite.GetAdapter<IMS_WWSPSUTControlAdapter>();
            }
        }

        /// <summary>
        /// A method is used to clean up the test suite.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        
        /// <summary>
        /// A test cases' level initialization method for TestSuiteBase class. It will perform before each test case.
        /// </summary>
        [TestInitialize]
        public void TestSuiteBaseInitialization()
        {
            Common.CheckCommonProperties(this.Site, true);

            // Check if MS-WWSP service is supported in current SUT.
            if (!Common.GetConfigurationPropertyValue<bool>("MS-WWSP_Supported", this.Site))
            {
                SutVersion currentSutVersion = Common.GetConfigurationPropertyValue<SutVersion>("SutVersion", this.Site);
                this.Site.Assert.Inconclusive("This test suite does not supported under current SUT, because MS-WWSP_Supported value set to false in MS-WWSP_{0}_SHOULDMAY.deployment.ptfconfig file.", currentSutVersion);
            }

            // Initialize the variables
            if (string.IsNullOrEmpty(CurrentProtocolPerformAccountName))
            {
                CurrentProtocolPerformAccountName = Common.GetConfigurationPropertyValue("MSWWSPTestAccount", TestSuiteBase.BaseTestSite);
                BaseTestSite.Assert.IsFalse(
                                      string.IsNullOrEmpty(CurrentProtocolPerformAccountName),
                                      @"The [MSWWSPTestAccount] property in [MS-WWSP_TestSuite.deployment.ptfconfig] file should have value.");
            }

            if (string.IsNullOrEmpty(DocLibraryName))
            {
                DocLibraryName = Common.GetConfigurationPropertyValue("CurrentDocLibraryListName", TestSuiteBase.BaseTestSite);
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(DocLibraryName),
                                    "The [CurrentDocLibraryListName] property in [MS-WWSP_TestSuite.deployment.ptfconfig] file should have value.");
            }

            if (string.IsNullOrEmpty(TaskListName))
            {
                TaskListName = Common.GetConfigurationPropertyValue("CurrentTaskListName", TestSuiteBase.BaseTestSite);
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(TaskListName),
                                    "The [CurrentTaskListName] property in [MS-WWSP_TestSuite.deployment.ptfconfig] file should have value.");
            }

            if (string.IsNullOrEmpty(WorkflowAssociationName))
            {
                WorkflowAssociationName = Common.GetConfigurationPropertyValue("WorkflowAssociationName", TestSuiteBase.BaseTestSite);
                this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(WorkflowAssociationName),
                                    "The [WorkflowAssociationName] property in [MS-WWSP_TestSuite.deployment.ptfconfig] file should have value.");
            }

            if (string.IsNullOrEmpty(DocListId))
            {
                DocListId = SutController.GetListIdByName(DocLibraryName);
                this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(DocListId),
                                @"Should get the valid list id of the Document Library type list[""{0}""]",
                                DocLibraryName);
            }

            if (string.IsNullOrEmpty(TaskListId))
            {
                TaskListId = SutController.GetListIdByName(TaskListName);
                this.Site.Assert.IsFalse(
                              string.IsNullOrEmpty(TaskListId),
                              @"Should get the valid list id of the Task type list[""{0}""]",
                              TaskListName);
            }

            if (string.IsNullOrEmpty(WorkflowAssociationId))
            {
                WorkflowAssociationId = SutController.GetWorkflowAssociationIdByName(DocLibraryName, WorkflowAssociationName);
                this.Site.Assert.IsFalse(
                           string.IsNullOrEmpty(WorkflowAssociationId),
                           @"Should get the valid id of workflow Association[{0}] in the  Document Library type list[""{1}""]",
                           WorkflowAssociationName,
                           DocLibraryName);
            }
        }

        /// <summary>
        ///  A test cases' level clean up method for TestSuiteBase class. It will perform before each test case.
        /// </summary>
        [TestCleanup]
        public void TestSuiteBaseCleanUp()
        {
            bool isCleanUpUploadedFilesSucceed = false;
            bool isCleanUpStartedTasksSucceed = false;

            string uploadedfilesUrls = string.Empty;
            if (null != UploadedFilesUrlRecorder && UploadedFilesUrlRecorder.Count > 0)
            {
                StringBuilder strBuilder = new StringBuilder();
                foreach (string urlsOfUploadFileItem in UploadedFilesUrlRecorder)
                {
                    strBuilder.Append(urlsOfUploadFileItem + ",");
                }

                // Remove last "," symbol
                uploadedfilesUrls = strBuilder.ToString(0, strBuilder.Length - 1);
                isCleanUpUploadedFilesSucceed = SutController.CleanUpUploadedFiles(DocLibraryName, uploadedfilesUrls);
                UploadedFilesUrlRecorder.Clear();
            }
            else
            {
                isCleanUpUploadedFilesSucceed = true;
            }

            string taskids = string.Empty;
            if (null != StartedTaskIdsRecorder && StartedTaskIdsRecorder.Count > 0)
            {
                StringBuilder strBuilder = new StringBuilder();
                foreach (string taskIditem in StartedTaskIdsRecorder)
                {
                    strBuilder.Append(taskIditem + ",");
                }

                // Remove last "," symbol
                taskids = strBuilder.ToString(0, strBuilder.Length - 1);
                isCleanUpStartedTasksSucceed = SutController.CleanUpStartedTasks(TaskListName, taskids);
                StartedTaskIdsRecorder.Clear();
            }
            else
            {
                isCleanUpStartedTasksSucceed = true;
            }

            string cleanUpProcessLogs = string.Empty;
            if (!isCleanUpStartedTasksSucceed)
            {
                string taskListName = Common.GetConfigurationPropertyValue("CurrentTaskListName", this.Site);
                cleanUpProcessLogs = string.Format(
                                                "There are some failures when cleaning up below tasks in task list[{0}].\r\nTasks ids:\r\n{1}\r\n",
                                                taskListName,
                                                taskids);
            }

            if (!isCleanUpUploadedFilesSucceed)
            {
                string documentListName = Common.GetConfigurationPropertyValue("CurrentDocLibraryListName", this.Site);
                cleanUpProcessLogs = string.Format(
                                        "{0}\r\nThere are some failures when cleaning up below files in Document Library[{1}].\r\nFiles urls:\r\n{2}",
                                        cleanUpProcessLogs,
                                        documentListName,
                                        uploadedfilesUrls);
            }

            if (!string.IsNullOrEmpty(cleanUpProcessLogs))
            {
                this.Site.Assert.Fail(
                                    "Clean up errors:\r\n{0}",
                                   cleanUpProcessLogs);
            }
        }

        #endregion  Test Suite Initialization and clean up

        #region HelperMethods

        /// <summary>
        /// Get StartWorkflow Parameter from data file.
        /// </summary>
        /// <param name="testsite">Transfer ITestSite into Adapter,Make adapter can use ITestSite's function.</param>
        /// <param name="assignToUser">A parameter represent the value will be replaced to the "PlaceHolder" in startworkflowParameters xml data file. The placeHolder value will be load from configuration file. The task start from this start workflow parameter data, will assign to the value specified by this input parameter.</param>
        /// <returns>A return represents the start workflow parameter data.</returns>
        protected XmlElement GetStartWorkflowParameter(ITestSite testsite, string assignToUser)
        {
            if (null == testsite)
            {
                throw new System.ArgumentException("Parameter [testsite] should not be null.", "testsite");
            }

            if (string.IsNullOrEmpty(assignToUser))
            {
                throw new System.ArgumentException("Parameter [assignToUser] should not be null.", "assignToUser");
            }

            #region Prepare data file name and assign accountID

            // remove .com from the domain value.
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string startworkflowDatafileName = this.GetStartworkflowDatafileName(false);

            // Get user name for normal user type.
            string expectedAssignedToValue = this.GetUserNameFromConfigFile(false);
            expectedAssignedToValue = string.Format(@"{0}\{1}", domainName, expectedAssignedToValue);
            string currentVersion = Common.GetConfigurationPropertyValue("SUTVersion", this.Site);
            if (currentVersion.Equals("SharePointServer2013", StringComparison.OrdinalIgnoreCase))
            {
                expectedAssignedToValue = string.Format("{0}{1}", @"i:0#.w|", expectedAssignedToValue);
            }

            #endregion

            XmlDocument doc = new XmlDocument();
            doc.Load(startworkflowDatafileName);
            if (null == doc.DocumentElement)
            {
                string erromsg = string.Format("Could not load xml data from the file[{0}].", startworkflowDatafileName);
                throw new System.Exception(erromsg);
            }

            string placeHolderOfAssignedTo = Common.GetConfigurationPropertyValue("AssignedToPlaceHolder", this.Site);
            if (string.IsNullOrEmpty(placeHolderOfAssignedTo))
            {
                throw new Exception("Could not get a valid AssignToUser PlaceHolder, it indicate where the user value will be replace in a xml data file.");
            }

            if (doc.DocumentElement.OuterXml.IndexOf(placeHolderOfAssignedTo, StringComparison.OrdinalIgnoreCase) <= 0)
            {
                string errorMsg = string.Format(
                                            @"Could not find expected PlaceHolder ""{0}"" in a xml data file [{1}].",
                                            placeHolderOfAssignedTo,
                                            startworkflowDatafileName);
                throw new Exception(errorMsg);
            }
            else
            {
                string replacedXmlString = doc.DocumentElement.OuterXml.Replace(placeHolderOfAssignedTo, expectedAssignedToValue);
                doc.LoadXml(replacedXmlString);
            }

            return doc.DocumentElement;
        }

        /// <summary>
        /// Get GetStartWorkflowParameter data for Claim used.
        /// </summary>
        /// <param name="testsite">Transfer ITestSite into Adapter,Make adapter can use ITestSite's function.</param>
        /// <returns>A return represents the StartWorkflowParameter data.</returns>
        protected XmlElement GetStartWorkflowParameterForClaim(ITestSite testsite)
        {
            if (null == testsite)
            {
                throw new System.ArgumentException("Parameter [testsite] should not be null.");
            }

            #region Prepare data file name and accountID

            // remove .com from the domain value.
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string startworkflowDatafileName = this.GetStartworkflowDatafileName(true);

            // Get Display name for UserGroup type.
            string expectedAssignedToValue = this.GetUserNameFromConfigFile(true);
            expectedAssignedToValue = string.Format(@"{0}\{1}", domainName, expectedAssignedToValue);
            string currentVersion = Common.GetConfigurationPropertyValue("SUTVersion", this.Site);
            if (currentVersion.Equals("SharePointServer2013", StringComparison.OrdinalIgnoreCase))
            {
                expectedAssignedToValue = string.Format("{0}{1}", @"c:0+.w|", expectedAssignedToValue);
            }

            #endregion
            XmlDocument doc = new XmlDocument();
            doc.Load(startworkflowDatafileName);
            if (null == doc.DocumentElement)
            {
                string erromsg = string.Format("Could not load xml data from the file[{0}].", startworkflowDatafileName);
                throw new System.Exception(erromsg);
            }

            string placeHolderOfAssignedTo = Common.GetConfigurationPropertyValue("AssignedToPlaceHolder", this.Site);
            if (string.IsNullOrEmpty(placeHolderOfAssignedTo))
            {
                throw new Exception("Could not get a valid AssignToUser PlaceHolder, it indicate where the user value will be replace in a xml data file.");
            }

            if (doc.DocumentElement.OuterXml.IndexOf(placeHolderOfAssignedTo, StringComparison.OrdinalIgnoreCase) <= 0)
            {
                string errorMsg = string.Format(
                                            @"Could not find expected PlaceHolder ""{0}"" in a xml data file [{1}].",
                                            placeHolderOfAssignedTo,
                                            startworkflowDatafileName);
                throw new Exception(errorMsg);
            }
            else
            {
                string replacedXmlString = doc.DocumentElement.OuterXml.Replace(placeHolderOfAssignedTo, expectedAssignedToValue);
                doc.LoadXml(replacedXmlString);
            }

            return doc.DocumentElement;
        }

        /// <summary>
        /// Get the AlerTodoTask data from alertToDo xml data file, and replace the "PlaceHolder" in xml data file with specified value. The "PlaceHolder" is specified in configuration file.
        /// </summary>
        /// <param name="testsite">Transfer ITestSite into Adapter,Make adapter can use ITestSite's function.</param>
        /// <param name="alertValue">A parameter represent the value will be replaced to the "PlaceHolder" in alertToDo xml data file. The placeHolder value will be load from configuration file.</param>
        /// <returns>A return represents the AlerToDoTask data.</returns>
        protected XmlElement GetAlerToDoTaskData(ITestSite testsite, string alertValue)
        {
            if (null == testsite)
            {
                throw new System.ArgumentException("Parameter [testsite] should not be null.");
            }

            string alerToDoTaskDataFile = Common.GetConfigurationPropertyValue("AlertToDoDataFile", Site);
            string placeHolder = Common.GetConfigurationPropertyValue("AlertedValuePlaceHolder", Site);

            if (string.IsNullOrEmpty(alerToDoTaskDataFile))
            {
                throw new Exception(@"The [AlertToDoDataFile] property does not have valid value. It should be ""*.xml"" format.");
            }

            if (string.IsNullOrEmpty(placeHolder))
            {
                throw new Exception(@"The [AlertedValuePlaceHolder] property does not have valid value. It should be ""[****]"" format.");
            }

            XmlDocument doc = new XmlDocument();
            doc.Load(alerToDoTaskDataFile);
            if (null == doc.DocumentElement)
            {
                string erromsg = string.Format("Could not load xml data from the file[{0}].", alerToDoTaskDataFile);
                throw new System.Exception(erromsg);
            }

            if (doc.DocumentElement.OuterXml.IndexOf(placeHolder, StringComparison.OrdinalIgnoreCase) <= 0)
            {
                string errorMsg = string.Format(
                                    "The AlertToDo task data file[{0}] should contain a expected PlaceHolder[{1}]",
                                    alerToDoTaskDataFile,
                                    placeHolder);
                throw new Exception(errorMsg);
            }

            // Replace the placeHolder with the specified value
            string replacedXmlString = doc.DocumentElement.OuterXml.Replace(placeHolder, alertValue);
            doc.LoadXml(replacedXmlString);

            return doc.DocumentElement;
        }

        /// <summary>
        /// Generate a Random value base on GUID format, this value will not be duplicated.
        /// </summary>
        /// <returns>A return represents a randomValue</returns>
        protected string GenerateRandomValue()
        {
            Guid guidTemp = Guid.NewGuid();
            string randomString = string.Format("{0}{1}", "MSWWSPTestUpdateValue", guidTemp.ToString("N"));
            return randomString;
        }

        /// <summary>
        /// Get the user name of account from configuration file according the accountType.
        /// </summary>
        /// <param name="isUserGroup">A parameter represents whether is UserGroup account type, true means it get a user group data.</param>
        /// <returns>A return represents the display name of account</returns>
        protected string GetUserNameFromConfigFile(bool isUserGroup)
        {
            string expectedUserName = string.Empty;

            if (isUserGroup)
            {
                expectedUserName = Common.GetConfigurationPropertyValue("UserGroupOnSUT", this.Site);
            }
            else
            {
                expectedUserName = Common.GetConfigurationPropertyValue("MSWWSPTestAccount", this.Site);
            }

            return expectedUserName;
        }

        /// <summary>
        /// Get StartworkflowData file Name from configuration file according startworkflow type.
        /// </summary>
        /// <param name="isForClaim">A parameter represents whether the startworkflow is for claim, true means it is for claim.</param>
        /// <returns>A return value represents the startworkflowData file Name</returns>
        protected string GetStartworkflowDatafileName(bool isForClaim)
        {
            string startworkflowDatafileName = string.Empty;
            if (isForClaim)
            {
                startworkflowDatafileName = Common.GetConfigurationPropertyValue("startworkflowParameterDataFileForClaim", this.Site);
            }
            else
            {
                startworkflowDatafileName = Common.GetConfigurationPropertyValue("startworkflowParameterDataFile", this.Site);
            }

            // Process the startworkflowParameter DataFile for different SUT in Microsoft Products
            if (startworkflowDatafileName.IndexOf(@"[SUTVersionShortName]", StringComparison.OrdinalIgnoreCase) > 0)
            {
                startworkflowDatafileName = startworkflowDatafileName.ToLower();
                string expectedSutPlaceHolderValue = string.Empty;
                string currentVersion = Common.GetConfigurationPropertyValue("SUTVersion", this.Site);
                if (currentVersion.Equals("SharePointServer2007", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2007";
                }
                else if (currentVersion.Equals("SharePointServer2010", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2010";
                }
                else if (currentVersion.Equals("SharePointServer2013", StringComparison.OrdinalIgnoreCase))
                {
                    expectedSutPlaceHolderValue = "2013";
                }
                else
                {
                    throw new Exception("Could Not Generate correct startworkflowParameter DataFile file name.");
                }

                startworkflowDatafileName = startworkflowDatafileName.Replace("[SUTVersionShortName]".ToLower(), expectedSutPlaceHolderValue);
            }

            return startworkflowDatafileName;
        }

        /// <summary>
        /// upload a file to the specified document library and record the file URL in order to cleanup the file in test suite clean up process.
        /// </summary>
        /// <param name="documentLibraryTitle">A parameter represents the title of a document library where the file upload</param>
        /// <returns>A return value represents the URL of the uploaded file.</returns>
        protected string UploadFileToSut(string documentLibraryTitle)
        {
            if (string.IsNullOrEmpty(documentLibraryTitle))
            {
                throw new ArgumentException("Should specify valid value to indicate the title of target document library", "documentLibraryTitle");
            }

            string uploadFileUrl = SutController.UploadFileToDocumentLibrary(documentLibraryTitle);
            if (string.IsNullOrEmpty(uploadFileUrl))
            {
                this.Site.Assert.Fail("Upload file to [{0}] list fail.", documentLibraryTitle);
            }

            UploadedFilesUrlRecorder.Add(uploadFileUrl);
            return uploadFileUrl;
        }
        
        /// <summary>
        /// Start workflow task for specified document item and record the task id in order to cleanup the task in test suite clean up process.
        /// </summary>
        /// <param name="uploadedFileUrl">A parameter represents a URL of document item where the task starts.</param>
        /// <param name="isClaim">A parameter represents a Boolean value indicate whether the task is started for Claim usage.</param>
        /// <returns>A return value represents the id of the task started by this method.</returns>
        protected string StartATaskWithNewFile(string uploadedFileUrl, bool isClaim)
        {
            if (string.IsNullOrEmpty(uploadedFileUrl))
            {
                throw new ArgumentException("Should specify valid value to indicate the URL of a uploaded file", "uploadedFileUrl");
            }

            GetToDosForItemResponseGetToDosForItemResult todosInfo = null;
            try
            {   
                XmlElement startWorkFlowData = null;
                if (isClaim)
                {   
                    startWorkFlowData = this.GetStartWorkflowParameterForClaim(this.Site);
                }
                else
                {
                    string assignToUserValue = Common.GetConfigurationPropertyValue("MSWWSPTestAccount", this.Site);
                    startWorkFlowData = this.GetStartWorkflowParameter(this.Site, assignToUserValue);
                }

                Guid currentWorkflowAsscociationId = new Guid(WorkflowAssociationId);
                ProtocolAdapter.StartWorkflow(uploadedFileUrl, currentWorkflowAsscociationId, startWorkFlowData);
                todosInfo = ProtocolAdapter.GetToDosForItem(uploadedFileUrl);
            }
            catch (SoapException soapEx)
            {
                throw new Exception("There are errors generated during start workflow task and GetToDos process.", soapEx);
            }

            // Get the task id from the response. there have only one task data in the response, because test suite only start one for new upload file.
            string taskId = this.GetZrowAttributeValueFromGetToDosForItemResponse(todosInfo, 0, "ows_ID");

            if (string.IsNullOrEmpty(taskId))
            {
                this.Site.Assert.Fail("Could Not get the id of started task from the document item[{0}]", uploadedFileUrl);
            }

            StartedTaskIdsRecorder.Add(taskId);
            return taskId;
        }

        /// <summary>
        /// Verify a new upload file's task data, if there are any task data for a new upload file, method will throw a Assert.Fail exception.
        /// </summary>
        /// <param name="uploadedFileUrl">A parameter represents a URL of document item which the method check.</param>
        protected void VerifyTaskDataOfNewUploadFile(string uploadedFileUrl)
        {  
            if (string.IsNullOrEmpty(uploadedFileUrl))
            {
                throw new ArgumentException("Should specify valid value to indicate the URL of a uploaded file", "uploadedFileUrl");
            }

            GetToDosForItemResponseGetToDosForItemResult gettodosResult = ProtocolAdapter.GetToDosForItem(uploadedFileUrl);
            this.Site.Assert.IsNotNull(gettodosResult, "The response of GetToDosForItem operation should have instance.");

            // Verify there are no any task data for the new upload file.
            if (gettodosResult.ToDoData != null && gettodosResult.ToDoData.xml != null && gettodosResult.ToDoData.xml != null
                && gettodosResult.ToDoData.xml.data != null && gettodosResult.ToDoData.xml.data.Any != null && gettodosResult.ToDoData.xml.data.Any.Length > 0)
            {
                this.Site.Assert.Fail(
                                "The response of GetToDosForItem operation should not contain any task data for a new uploaded file, actual task number:[{0}]",
                                 gettodosResult.ToDoData.xml.data.Any.Length);
            }
        }

        /// <summary>
        /// Verify a new upload file's task data, if there are any task data for a new upload file, method will throw a Assert.Fail exception.
        /// </summary>
        /// <param name="uploadedFileUrl">A parameter represents a URL of document item which the method check.</param>
        /// <param name="taskIndex">A parameter represents the index of a z:row item in a zrow collection. Each z:row item means a task item.It start on "Zero".</param>
        protected void VerifyAssignToValueForSingleTodoItem(string uploadedFileUrl, int taskIndex)
        {
            GetToDosForItemResponseGetToDosForItemResult todosAfterStartTask = ProtocolAdapter.GetToDosForItem(uploadedFileUrl);
            if (null == todosAfterStartTask || null == todosAfterStartTask.ToDoData || null == todosAfterStartTask.ToDoData.xml
                || null == todosAfterStartTask.ToDoData.xml.data || null == todosAfterStartTask.ToDoData.xml.data.Any)
            {
                this.Site.Assert.Fail("The response of GetToDosForItem operation should contain valid data.");
            }

            this.Site.Assert.AreEqual(1, todosAfterStartTask.ToDoData.xml.data.Any.Length, "The response of GetToDosForItem operation for new file should contain only one task data.");
            string expectedAssignToValue = Common.GetConfigurationPropertyValue("KeyWordForAssignedToField", this.Site).ToLower();

            // Get the assign to value from actual response, "ows_AssignedTo" means get required field "AssignedTo" of task list.
            string actualAssignToValue = Common.GetZrowAttributeValue(todosAfterStartTask.ToDoData.xml.data.Any, taskIndex, "ows_AssignedTo");
            this.Site.Assert.IsFalse(string.IsNullOrEmpty(actualAssignToValue), "The response of GetToDosForItem should contain valid value for [AssignedTo] field.");

            bool isassignToExpectedValue = actualAssignToValue.IndexOf(expectedAssignToValue, StringComparison.OrdinalIgnoreCase) >= 0;
            this.Site.Assert.IsTrue(
                                isassignToExpectedValue,
                                @"The actual ""AssignedTo"" value [{0}] should contain to expected keyword value[{1}]",
                                actualAssignToValue,
                                expectedAssignToValue);
        }

        /// <summary>
        /// Get z:row attribute value by specified attribute name from a GetToDosForItemResponse
        /// </summary>
        /// <param name="getToDosForItemResponse">A parameter represents a GetToDosForItemResponse which contain zrow Attribute Values.</param>
        /// <param name="index">A parameter represents the index of a zrow item in a zrow collection.It start on "Zero".</param>
        /// <param name="expectedAttibuteName">A parameter represents the attributeName of which value will be return.</param>
        /// <returns>A return value represents the attribute value.</returns>
        protected string GetZrowAttributeValueFromGetToDosForItemResponse(GetToDosForItemResponseGetToDosForItemResult getToDosForItemResponse, int index, string expectedAttibuteName)
        {
            if (null == getToDosForItemResponse || null == getToDosForItemResponse.ToDoData
                || null == getToDosForItemResponse.ToDoData.xml)
            {
                throw new ArgumentException("The parameter should contain valid ToDoData.", "getToDosForItemResponse");
            }

            XmlElement[] zrowDatas = getToDosForItemResponse.ToDoData.xml.data.Any;
            return Common.GetZrowAttributeValue(zrowDatas, index, expectedAttibuteName);
        }

        #endregion
    }
}