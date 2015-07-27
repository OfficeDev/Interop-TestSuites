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
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WWSP S01 StartGetWorkflow.
    /// </summary>
    [TestClass]
    public class S01_StartWorkflow : TestSuiteBase
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

        #region Test cases
        /// <summary>
        /// This test case is used to verify StartWorkflow operation,  starts a new workflow successfully.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S01_TC01_StartWorkflow_Success()
        {   
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start work flow
            string assigntoUserName = CurrentProtocolPerformAccountName;
            XmlElement starworkflowParameter = this.GetStartWorkflowParameter(this.Site, assigntoUserName);
            Guid targetWorkflowAssociationGuid = new Guid(WorkflowAssociationId);
            ProtocolAdapter.StartWorkflow(uploadFileUrl, targetWorkflowAssociationGuid, starworkflowParameter);

            GetToDosForItemResponseGetToDosForItemResult todosAfterStartTask = ProtocolAdapter.GetToDosForItem(uploadFileUrl);
            if (null == todosAfterStartTask || null == todosAfterStartTask.ToDoData || null == todosAfterStartTask.ToDoData.xml
                || null == todosAfterStartTask.ToDoData.xml.data || null == todosAfterStartTask.ToDoData.xml.data.Any)
            {
                this.Site.Assert.Fail("The response of GetToDosForItem operation should contain valid data.");
            }

            // Record the started tasks
            string taskId = this.GetZrowAttributeValueFromGetToDosForItemResponse(todosAfterStartTask, 0, "ows_ID");
            StartedTaskIdsRecorder.Add(taskId);

            this.Site.Assert.AreEqual(1, todosAfterStartTask.ToDoData.xml.data.Any.Length, "The response of GetToDosForItem operation for new file should contain only one task data.");
            assigntoUserName = assigntoUserName.ToLower();

            // Get the assign to value from actual response, "ows_AssignedTo" means get required field "AssignedTo" of task list.
            string actualAssignToValue = Common.GetZrowAttributeValue(todosAfterStartTask.ToDoData.xml.data.Any, 0, "ows_AssignedTo");
            this.Site.Assert.IsFalse(string.IsNullOrEmpty(actualAssignToValue), "The response of GetToDosForItem should contain valid value for [AssignedTo] field.");

            bool isassignToExpectedValue = actualAssignToValue.IndexOf(assigntoUserName, StringComparison.OrdinalIgnoreCase) >= 0;

            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: assigntoUserName[{0}], actualAssignToValue[{1}] for requirement #R412, #R324, #R327, #R330, #R339",
                assigntoUserName,
                actualAssignToValue);

            // If start a new workflow instance and assign to expected user, capture #R412, R324, R327, R330, R339
            this.Site.CaptureRequirementIfIsTrue(
                                            isassignToExpectedValue,
                                            412,
                                            @"[In Message Processing Events and Sequencing Rules] StartWorkflow
 instantiates a new workflow for an existing document and a workflow association.");

            this.Site.CaptureRequirementIfIsTrue(
                                             isassignToExpectedValue,
                                             324,
                                             @"[In StartWorkflow] This operation[StartWorkflow] starts a new workflow, generating a workflow from a workflow association.");

            this.Site.CaptureRequirementIfIsTrue(
                                         isassignToExpectedValue,
                                         327,
                                         @"[In StartWorkflow] The protocol client sends a StartWorkflowSoapIn request message, and the protocol server responds with a StartWorkflowSoapOut response message.");

            this.Site.CaptureRequirementIfIsTrue(
                                       isassignToExpectedValue,
                                       330,
                                       @"[In Messages] StartWorkflowSoapOut specifies the response to a request to start a new workflow.");

            this.Site.CaptureRequirementIfIsTrue(
                                        isassignToExpectedValue,
                                        339,
                                        @"[In Elements] StartWorkflowResponse contains the response to a request to start a new workflow.");
        }

        #endregion Test cases
    }
}