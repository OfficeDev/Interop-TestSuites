namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, updating, movement, retrieving, copy and deletion of the task items with or without optional elements in the server.
    /// </summary>
    [TestClass]
    public class S06_OperateTaskItemWithOptionalElements : TestSuiteBase
    {
        #region Class initialize and clean up

        /// <summary>
        /// Initialize the test class.
        /// </summary>
        /// <param name="context">Context to initialize.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        /// Clean up the test class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases
        /// <summary>
        /// This test case is intended to validate the success response of operating task item without optional elements, returned by CreateItem and GetItem operations.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S06_TC01_OperateTaskItemWithoutOptionalElements()
        {
            #region Client calls CreateItem operation to create a task item without optional elements.
            // Save the ItemId of task item got from the createItem response.
            ItemIdType[] createItemIds = this.CreateTasks(new TaskType());
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls GetItem operation to get the task item.
            this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            #endregion

            #region Client calls UpdateItem operation to update the value of taskCompanies element of task item.
            ItemIdType[] updateItemIds = this.UpdateTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This update response status should be success!", null);
            ItemIdType updateItemId = updateItemIds[0];
            #endregion

            #region Client calls CopyItem operation to copy the created task item.
            ItemIdType[] copyItemIds = this.CopyTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This copy response status should be success!", null);
            ItemIdType copyItemId = copyItemIds[0];
            #endregion

            #region Client calls MoveItem operation to move the task item.
            ItemIdType[] moveItemIds = this.MoveTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This move response status should be success!", null);
            ItemIdType moveItemId = moveItemIds[0];
            #endregion

            #region Client calls DeleteItem to delete the task items created in the previous steps.
            this.DeleteTasks(copyItemId, moveItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the success response of operating task item with optional elements, returned by CreateItem and GetItem operations.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S06_TC02_OperateTaskItemWithOptionalElements()
        {
            #region Client calls CreateItem operation to create a task item with optional elements.

            // All the optional elements in task item are set in this method.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, TaskStatusType.Completed);
           
            // Save the ItemId of task item got from the createItem response.
            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion
         
            #region Client call GetItem operation to get the task item.
            TaskType[] retrievedTaskItems = this.GetTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItem = retrievedTaskItems[0];
            #endregion

            #region Verify the related requirements about sub-elements of TaskType
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R42");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R42
            this.Site.CaptureRequirementIfAreEqual<int>(
                sentTaskItem.ActualWork,
                retrievedTaskItem.ActualWork,
                42,
                @"[In t:TaskType Complex Type] ActualWork: Specifies an integer value that specifies the actual amount of time that is spent on a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R44");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R44
            this.Site.CaptureRequirementIfAreEqual<string>(
                sentTaskItem.BillingInformation,
                retrievedTaskItem.BillingInformation,
                44,
                @"[In t:TaskType Complex Type] BillingInformation: Specifies a string value that contains billing information for a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R47, Expected value:" + sentTaskItem.Companies[0].ToString() + " " + sentTaskItem.Companies[1].ToString() + "Actual value:" + retrievedTaskItem.Companies[0].ToString() + " " + retrievedTaskItem.Companies[1].ToString());

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R47
            bool isVerifiedR47 = TestSuiteHelper.CompareStringArray(retrievedTaskItem.Companies, sentTaskItem.Companies);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR47,
                47,
                @"[In t:TaskType Complex Type] Companies: Specifies an instance of an array of type string that represents a collection of companies that are associated with a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R49, Expected value" + sentTaskItem.Contacts[0].ToString() + " " + sentTaskItem.Contacts[1].ToString() + "Actual value:" + retrievedTaskItem.Contacts[0].ToString() + " " + retrievedTaskItem.Contacts[1].ToString());

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R49
            bool isVerifiedR49 = TestSuiteHelper.CompareStringArray(retrievedTaskItem.Contacts, sentTaskItem.Contacts);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR49,
                49,
                @"[In t:TaskType Complex Type] Contacts: Specifies an instance of an array of type string that contains a list of contacts that are associated with a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R53");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R53
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                sentTaskItem.DueDate.Date,
                retrievedTaskItem.DueDate.Date,
                53,
                @"[In t:TaskType Complex Type] DueDate: Specifies an instance of the DateTime structure that represents the date when a task is due.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R58");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R58
            this.Site.CaptureRequirementIfAreEqual<string>(
                sentTaskItem.Mileage,
                retrievedTaskItem.Mileage,
                58,
                @"[In t:TaskType Complex Type] Mileage: Specifies a string value that represents the mileage for a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R63");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R63
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                sentTaskItem.StartDate.Date,
                retrievedTaskItem.StartDate.Date,
                63,
                @"[In t:TaskType Complex Type] StartDate: Specifies an instance of the DateTime structure that represents the start date of a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R65");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R65
            this.Site.CaptureRequirementIfAreEqual<TaskStatusType>(
                sentTaskItem.Status,
                retrievedTaskItem.Status,
                65,
                @"[In t:TaskType Complex Type] Status: Specifies one of the valid TaskStatusType simple type enumeration values that represent the status of a task.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R67");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R67
            this.Site.CaptureRequirementIfAreEqual<int>(
                sentTaskItem.TotalWork,
                retrievedTaskItem.TotalWork,
                67,
                @"[In t:TaskType Complex Type] TotalWork: Specifies an integer value that represents the total amount of work that is associated with a task.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R45");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R45
            this.Site.CaptureRequirementIfAreEqual<int>(
                sentTaskItem.ChangeCount+1,
                retrievedTaskItem.ChangeCount,
                45,
                @"[In t:TaskType Complex Type] ChangeCount: Specifies an integer value that specifies the number of times the task has changed since it was created.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R48");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R48
            this.Site.CaptureRequirementIfAreEqual<DateTime>(
                sentTaskItem.CompleteDate.Date,
                retrievedTaskItem.CompleteDate.Date,
                48,
                @"[In t:TaskType Complex Type] CompleteDate: Specifies an instance of the DateTime structure that represents the date on which a task was completed.");

            #endregion

            #region Client calls UpdateItem operation to update the value of taskCompanies element of task item.
            ItemIdType[] updateItemIds = this.UpdateTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This update response status should be success!", null);
            ItemIdType updateItemId = updateItemIds[0];
            #endregion

            #region Client calls CopyItem operation to copy the created task item.
            ItemIdType[] copyItemIds = this.CopyTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This copy response status should be success!", null);
            ItemIdType copyItemId = copyItemIds[0];
            #endregion

            #region Client calls MoveItem operation to move the task item.
            ItemIdType[] moveItemIds = this.MoveTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This move response status should be success!", null);
            ItemIdType moveItemId = moveItemIds[0];
            #endregion

            #region Client calls DeleteItem to delete the task items created in the previous steps.
            this.DeleteTasks(copyItemId, moveItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the success response of operating task item with IsComplete element.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S06_TC03_OperateTaskItemWithIsCompleteElement()
        {
            #region Client calls CreateItem operation to create a task item with task Status equal to Completed.
            // Save the ItemId of task item got from the createItem response.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsFirst = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, TaskStatusType.Completed));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdFirst = createItemIdsFirst[0];
            #endregion

            #region Client call GetItem operation to get the task item.
            TaskType[] retrievedTaskItemsFirst = this.GetTasks(createItemIdFirst);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemFirst = retrievedTaskItemsFirst[0];
            #endregion

            #region Verify the IsComplete element value
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R5555");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R5555
            this.Site.CaptureRequirementIfIsTrue(
                retrievedTaskItemFirst.IsComplete,
                5555,
                @"[In t:TaskType Complex Type] [IsComplete is] True, indicates a task has been completed.");

            #endregion

            #region Client calls CreateItem operation to create a task item with task Status equal to InProgress.
            // Save the ItemId of task item got from the createItem response.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsSecond = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, TaskStatusType.InProgress));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdSecond = createItemIdsSecond[0];
            #endregion

            #region Client call GetItem operation to get the task item.
            TaskType[] retrievedTaskItemsSecond = this.GetTasks(createItemIdSecond);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemSecond = retrievedTaskItemsSecond[0];
            #endregion

            #region Verify the IsComplete element value

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R5556");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R5556
            this.Site.CaptureRequirementIfIsFalse(
                retrievedTaskItemSecond.IsComplete,
                5556,
                @"[In t:TaskType Complex Type] [IsComplete is] False, indicates a task has not been completed.");

            #endregion

            #region Client calls DeleteItem to delete the task items created in the previous steps.
            this.DeleteTasks(createItemIdFirst, createItemIdSecond);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the success response of operating task item with IsRecurring element.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S06_TC04_OperateTaskItemWithIsRecurringElement()
        {
            #region Client calls CreateItem operation to create a task item, which is a recurring task.
            // Configure the DailyRegeneratingPatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateDailyRegeneratingPattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Save the ItemId of task item got from the createItem response.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsFirst = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, taskRecurrence));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdFirst = createItemIdsFirst[0];
            #endregion

            #region Client call GetItem operation to get the task item.
            TaskType[] retrievedTaskItemsFirst = this.GetTasks(createItemIdFirst);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemFirst = retrievedTaskItemsFirst[0];
            #endregion

            #region Verify the IsRecurring element value

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R5666");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R5666
            this.Site.CaptureRequirementIfIsTrue(
                retrievedTaskItemFirst.IsRecurring,
                5666,
                @"[In t:TaskType Complex Type] [IsRecurring is] True, indicates a task is part of a recurring task.");

            #endregion

            #region Client calls CreateItem operation to create a task item, which is not a recurring task.
            // Save the ItemId of task item got from the createItem response.
            subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIdsSecond = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, null));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemIdSecond = createItemIdsSecond[0];
            #endregion

            #region Client call GetItem operation to get the task item.
            TaskType[] retrievedTaskItemsSecond = this.GetTasks(createItemIdSecond);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType retrievedTaskItemSecond = retrievedTaskItemsSecond[0];
            #endregion

            #region Verify the IsRecurring element value

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R5667");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R5667
            this.Site.CaptureRequirementIfIsFalse(
                retrievedTaskItemSecond.IsRecurring,
                5667,
                @"[In t:TaskType Complex Type] [IsRecurring is] False, indicates a task is not part of a recurring task.");

            #endregion

            #region Client calls DeleteItem to delete the task items created in the previous steps.
            this.DeleteTasks(createItemIdFirst, createItemIdSecond);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }
        #endregion
    }
}