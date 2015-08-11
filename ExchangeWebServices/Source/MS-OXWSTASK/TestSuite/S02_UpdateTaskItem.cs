namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test update operation.
    /// </summary>
    [TestClass]
    public class S02_UpdateTaskItem : TestSuiteBase
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
        /// This test case is used to verify whether an occurrence of a task or a master task is deleted.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S02_TC01_VerifyAffectedTaskOccurrencesType()
        {
            #region Client calls CreateItem to create a task item that contains the Recurrence element, which includes the DailyRecurrencePatternType.

            // Configure the DailyRecurrencePatternType.
            TaskRecurrenceType taskRecurrence = TestSuiteHelper.GenerateTaskRecurrence(TestSuiteHelper.GenerateDailyRecurrencePattern, TestSuiteHelper.GenerateNumberedRecurrenceRange);

            // Define a task item.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            TaskType sentTaskItem = TestSuiteHelper.DefineTaskItem(subject, taskRecurrence);

            ItemIdType[] createItemIds = this.CreateTasks(sentTaskItem);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls UpdateItem to update the value of "companies" element of task item.
            ItemIdType[] updateItemIds = this.UpdateTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This update response status should be success!", null);
            ItemIdType updateItemId = updateItemIds[0];
            #endregion

            #region Client calls GetItem to check whether the task item' "companies" element is updated.
            TaskType[] taskItemsAfterUpdate = this.GetTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType taskItemAfterUpdate = taskItemsAfterUpdate[0];
            bool isEqual = TestSuiteHelper.CompareStringArray(taskItemAfterUpdate.Companies, new string[] { "Company3", "Company4" });
            Site.Assert.IsTrue(isEqual, "After updated, the task companies names should be Company3, Company4", null);
            #endregion

            #region Client calls DeleteItem to delete only the current occurrence of a task.
            this.DeleteTasks(AffectedTaskOccurrencesType.SpecifiedOccurrenceOnly, updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion

            #region Client calls GetItem to check whether only current occurrence of task item is deleted.

            TaskType[] taskItemsAfterDelete = GetTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            TaskType taskItemAfterDelete = taskItemsAfterDelete[0];

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R175");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R175
            // After deleting the current occurrence of a task item, the start time of task item is changed to next occurrence.
            Site.CaptureRequirementIfAreEqual<string>(
                ExtractStartTimeOfNextOccurrence(taskItemAfterUpdate),
                taskItemAfterDelete.StartDate.ToShortTimeString(),
                175,
                @"[In t:AffectedTaskOccurrencesType Simple Type]  SpecifiedOccurrenceOnly: Specifies that a DeleteItem operation request, as specified in [MS-OXWSCORE] section 3.1.4.3, deletes only the current occurrence of a task.");

            #endregion

            #region Client calls DeleteItem to delete all recurring tasks.
            this.DeleteTasks(updateItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion

            #region Client calls GetItem to check whether all recurring tasks is deleted.
            this.GetTasks(updateItemId);
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should not be success!", null);
            Site.Assert.AreEqual<ResponseCodeType>(ResponseCodeType.ErrorItemNotFound, (ResponseCodeType)this.ResponseCode[0], "This get response status information should be ErrorItemNotFound!", null);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSTASK_R174");

            // Verify MS-OXWSTASK requirement: MS-OXWSTASK_R174 
            // After deleting all recurring tasks with the AffectedTaskOccurrences element setting to AllOccurrences, 
            // the operation of getting task will return code ErrorItemNotFound. If getting this response code, the following 
            // requirement will be captured directly.
            Site.CaptureRequirement(
                174,
                @"[In t:AffectedTaskOccurrencesType Simple Type] AllOccurrences: Specifies that a DeleteItem operation request, as specified in [MS-OXWSCORE] section 3.1.4.3, deletes the master task and all recurring tasks that are associated with the master task.");

            #endregion
        }
        #endregion

        #region Private methods
        /// <summary>
        /// Extract the start time of next occurrence for a DailyRecurrencePatternType task item.
        /// </summary>
        /// <param name="taskItem">The task item with daily recurrence pattern.</param>
        /// <returns>The start date time string of next occurrence.</returns>
        private static string ExtractStartTimeOfNextOccurrence(TaskType taskItem)
        {
            if (taskItem != null && taskItem.Recurrence != null && taskItem.Recurrence.Item != null && taskItem.Recurrence.Item1 != null)
            {
                DailyRecurrencePatternType dailyRecurrencePattern = taskItem.Recurrence.Item as DailyRecurrencePatternType;
                return taskItem.StartDate.AddDays(dailyRecurrencePattern.Interval).ToShortTimeString();
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}