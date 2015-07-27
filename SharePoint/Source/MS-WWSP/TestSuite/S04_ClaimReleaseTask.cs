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
    using System.Web.Services.Protocols;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The TestSuite of MS-WWSP S04 ClaimReleaseTask.
    /// </summary>
    [TestClass]
    public class S04_ClaimReleaseTask : TestSuiteBase
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
        /// This test case is used to verify ClaimReleaseTask operation when the document URL is correct.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S04_TC01_ClaimReleaseTask_CorrectURL()
        {
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start work flow task for Claim operation, and method will start a task assign to a value which is equal to "UserGroupOnSUT" property in PTF configuration file.
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, true);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            string expectedAssigntoGroupName = Common.GetConfigurationPropertyValue("UserGroupOnSUT", this.Site);
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // Execute the Claim operation.
            Guid targetTaskListId = new Guid(TaskListId);
            ClaimReleaseTaskResponseClaimReleaseTaskResult claimResult = null;
            try
            {
                claimResult = ProtocolAdapter.ClaimReleaseTask(
                                       uploadFileUrl,
                                       int.Parse(taskIdValue),
                                       targetTaskListId,
                                       true);
            }
            catch (SoapException sopaEx)
            {
                this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "Generate a SoapException during calling ClaimReleaseTask operation.Error:[{0}],\r\n Target file[{1}],\r\n Task AssignTo[{2}]",
                        sopaEx.Message,
                        uploadFileUrl,
                        expectedAssigntoGroupName);
                throw sopaEx;
            }

            // If ClaimReleaseTask operation is succeed, SUT return a response, then capture R407, R110, R116, R125, R113
            // Verify MS-WWSP requirement: MS-WWSP_R407
            Site.CaptureRequirementIfIsNotNull(
                claimResult,
                407,
                @"[In Message Processing Events and Sequencing Rules] ClaimReleaseTask claims (1) or releases a claim (1) on a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R110
            Site.CaptureRequirementIfIsNotNull(
                claimResult,
                110,
                @"[In ClaimReleaseTask] This operation[ClaimReleaseTask] claims (1) or releases a claim (1) on workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R116
            Site.CaptureRequirementIfIsNotNull(
                claimResult,
                116,
                @"[In Messages] ClaimReleaseTaskSoapOut specifies the response to a request to claim (1) or release a claim (1) on a workflow task.");

            // Verify MS-WWSP requirement: MS-WWSP_R125
            Site.CaptureRequirementIfIsNotNull(
                claimResult,
                125,
                @"[In Elements] ClaimReleaseTaskResponse contains the response to a request to claim (1) or release a claim (1) on a workflow task.");

            // Verify MS-WWSP requirement: R113
            this.Site.CaptureRequirementIfIsNotNull(
                                            claimResult,
                                            113,
                                            @"[In ClaimReleaseTask] The protocol client sends a ClaimReleaseTaskSoapIn request message, and the protocol server responds with a ClaimReleaseTaskSoapOut response message.");

            // If the itemId in response equal to the value specified in the request, then capture R145
            this.Site.CaptureRequirementIfAreEqual(
                                            int.Parse(taskIdValue),
                                            claimResult.TaskData.ItemId,
                                            145,
                                            @"[In ClaimReleaseTaskResponse] ClaimReleaseTaskResult.TaskData.ItemId: A list item identifier of a workflow task.");

            // If the ListId in response equal to the value of current task list id which is specified in request, then R146
            this.Site.CaptureRequirementIfAreEqual(
                                          targetTaskListId,
                                          claimResult.TaskData.ListId,
                                          146,
                                          @"[In ClaimReleaseTaskResponse] ClaimReleaseTaskResult.TaskData.ListId: The list identifier of the workflow task.");
        }

        /// <summary>
        /// This test case is used to verify ClaimReleaseTask operation. The client sets the different strings as the item values, and the server replies the same if the site of the SOAP request URL contains a list with the specified ListId.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S04_TC02_ClaimReleaseTask_IgnoreItem()
        {
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start work flow task for Claim operation, and method will start a task assign to a value which is equal to "UserGroupOnSUT" property in PTF configuration file.
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, true);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            string expectedAssigntoGroupName = Common.GetConfigurationPropertyValue("UserGroupOnSUT", this.Site);
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // Execute the Claim operation with correct item value.
            Guid targetTaskListId = new Guid(TaskListId);
            ClaimReleaseTaskResponseClaimReleaseTaskResult claimResultWithCorrectItemValue = null;
            claimResultWithCorrectItemValue = ProtocolAdapter.ClaimReleaseTask(
                                      uploadFileUrl,
                                      int.Parse(taskIdValue),
                                      targetTaskListId,
                                      true);

            // Release the claimed task so that next calling of ClaimReleaseTask operation could claim the task again.
            ClaimReleaseTaskResponseClaimReleaseTaskResult claimResultOfRelease = null;
            claimResultOfRelease = ProtocolAdapter.ClaimReleaseTask(
                                      uploadFileUrl,
                                      int.Parse(taskIdValue),
                                      targetTaskListId,
                                      false);

            // Verify whether release successfully
            this.Site.Assert.IsNotNull(claimResultOfRelease, "The response of ClaimReleaseTask operation should return valid response.");
            bool isassignToValueCorrespond = claimResultOfRelease.TaskData.AssignedTo.IndexOf(expectedAssigntoGroupName, StringComparison.OrdinalIgnoreCase) > 0;
            this.Site.Assert.IsTrue(
                                      isassignToValueCorrespond,
                                      @"Release claimed task action should set the ""assignTo"" value correspond back to the original assignment[{0}], but actual[{1}]",
                                      expectedAssigntoGroupName,
                                      claimResultOfRelease.TaskData.AssignedTo);

            // Execute the Claim operation with invalid item value.
            string invalidItemValue = this.GenerateRandomValue();
            ClaimReleaseTaskResponseClaimReleaseTaskResult claimResultwithInvalidItemValue = null;
            claimResultwithInvalidItemValue = ProtocolAdapter.ClaimReleaseTask(
                                      invalidItemValue,
                                      int.Parse(taskIdValue),
                                      targetTaskListId,
                                      true);

            bool checkResult = claimResultwithInvalidItemValue.AreEquals(claimResultWithCorrectItemValue);

            // Verify MS-WWSP requirement: MS-WWSP_R131
            bool isVerifyR131 = checkResult;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR131,
                131,
                @"[In ClaimReleaseTask] Set the different string as the item value, server reply same if the site (2) of the SOAP request URL  contains a list with the specified ListId.");
        }

        /// <summary>
        /// This test case is used to verify the element AssignedTo when call ClaimReleaseTask operation.
        /// </summary>
        [TestCategory("MSWWSP"), TestMethod()]
        public void MSWWSP_S04_TC03_ClaimReleaseTask_AssignedTo()
        {
            // Upload a file.
            string uploadFileUrl = this.UploadFileToSut(DocLibraryName);
            this.VerifyTaskDataOfNewUploadFile(uploadFileUrl);

            // Start work flow task for Claim operation, and method will start a task assign to a value which is equal to "UserGroupOnSUT" property in PTF configuration file.
            string taskIdValue = this.StartATaskWithNewFile(uploadFileUrl, true);

            // Verify whether the task is assign to expected user group. for new uploaded file, only have one task currently.
            this.VerifyAssignToValueForSingleTodoItem(uploadFileUrl, 0);

            // Claim a task, and the "assignToValue" will assign to a user in the group.
            // Execute the Claim operation with correct item value.
            Guid targetTaskListId = new Guid(TaskListId);
            ClaimReleaseTaskResponseClaimReleaseTaskResult resultOfClaimTask = null;
            resultOfClaimTask = ProtocolAdapter.ClaimReleaseTask(
                                      uploadFileUrl,
                                      int.Parse(taskIdValue),
                                      targetTaskListId,
                                      true);

            if (null == resultOfClaimTask || null == resultOfClaimTask.TaskData || string.IsNullOrEmpty(resultOfClaimTask.TaskData.AssignedTo))
            {
                this.Site.Assert.Fail("The response of ClaimReleaseTask operation should contain valid data in claim mode.");
            }

            string actualAssignToValue = resultOfClaimTask.TaskData.AssignedTo;

            // After Claim operation for the task, the task should assign to the user in the UserGroup.
            string expectedAssignToUserValue = CurrentProtocolPerformAccountName;
            bool isassignToExpectedUserForClaim = actualAssignToValue.IndexOf(expectedAssignToUserValue, StringComparison.OrdinalIgnoreCase) >= 0;

            // After Claim, call ClaimReleaseTask operation to release a claim task.
            resultOfClaimTask = ProtocolAdapter.ClaimReleaseTask(
                                   uploadFileUrl,
                                   int.Parse(taskIdValue),
                                   targetTaskListId,
                                   false);

            if (null == resultOfClaimTask || null == resultOfClaimTask.TaskData || string.IsNullOrEmpty(resultOfClaimTask.TaskData.AssignedTo))
            {
                this.Site.Assert.Fail("The response of ClaimReleaseTask operation should contain valid data in release mode.");
            }

            // Verify MS-WWSP requirement: MS-WWSP_R142
            bool isVerifyR142 = actualAssignToValue.IndexOf(expectedAssignToUserValue, StringComparison.OrdinalIgnoreCase) >= 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR142,
                142,
                @"[In ClaimReleaseTaskResponse] ClaimReleaseTaskResult.TaskData.AssignedTo: The user to whom this workflow task is now assigned.");

            // Verify MS-WWSP requirement: MS-WWSP_R143
            bool isVerifyR143 = actualAssignToValue.IndexOf(expectedAssignToUserValue, StringComparison.OrdinalIgnoreCase) >= 0;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR143,
                143,
                @"[In ClaimReleaseTaskResponse] This[ClaimReleaseTaskResult.TaskData.AssignedTo] MUST be the user authenticated in section 1.5 if this operation is a claim and the protocol server requires authentication.");
        }

        #endregion Test cases
    }
}