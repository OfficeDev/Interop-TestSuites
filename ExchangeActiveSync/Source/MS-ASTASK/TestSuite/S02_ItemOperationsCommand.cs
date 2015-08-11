namespace Microsoft.Protocols.TestSuites.MS_ASTASK
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.Common.DataStructures;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using SyncItem = Microsoft.Protocols.TestSuites.Common.DataStructures.Sync;
    using SyncStore = Microsoft.Protocols.TestSuites.Common.DataStructures.SyncStore;

    /// <summary>
    /// This scenario is to test Task class element on the server by using ItemOperations command.
    /// </summary>
    [TestClass]
    public class S02_ItemOperationsCommand : TestSuiteBase
    {
        #region Test Class initialize and clean up

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        /// <summary>
        /// This test case is designed to verify the requirements about processing the ItemOperations command.
        /// </summary>
        [TestCategory("MSASTASK"), TestMethod()]
        public void MSASTASK_S02_TC01_RetrieveTaskItemWithItemOperations()
        {
            #region Call Sync command to create a task item

            Dictionary<Request.ItemsChoiceType8, object> taskItem = new Dictionary<Request.ItemsChoiceType8, object>();
            string subject = Common.GenerateResourceName(Site, "subject");
            taskItem.Add(Request.ItemsChoiceType8.Subject2, subject);

            #endregion

            #region Call Sync command to add the task to the server

            // add task
            SyncStore syncResponse = this.SyncAddTask(taskItem);
            Site.Assert.AreEqual<byte>(1, syncResponse.AddResponses[0].Status, "Adding a task item to server should success.");
            SyncItem task = this.GetChangeItem(this.UserInformation.TasksCollectionId, subject);
            Site.Assert.IsNotNull(task.Task, "The task which subject is {0} should exist in server.", subject);
            ItemsNeedToDelete.Add(subject);

            #endregion

            #region Call ItemOperations command to fetch tasks

            syncResponse = this.SyncChanges(this.UserInformation.TasksCollectionId);

            List<string> serverIds = new List<string>();
            for (int i = 0; i < syncResponse.AddElements.Count; i++)
            {
                serverIds.Add(syncResponse.AddElements[i].ServerId);
            }

            Request.Schema schema = new Request.Schema
            {
                ItemsElementName = new Request.ItemsChoiceType3[1],
                Items = new object[1]
            };
            schema.ItemsElementName[0] = Request.ItemsChoiceType3.Body;
            schema.Items[0] = new Request.Body();

            Request.BodyPreference bodyReference = new Request.BodyPreference { Type = 1 };

            ItemOperationsRequest itemOperationsRequest = TestSuiteHelper.CreateItemOperationsFetchRequest(this.UserInformation.TasksCollectionId, serverIds, null, bodyReference, schema);
            ItemOperationsStore itemOperationsResponse = this.TASKAdapter.ItemOperations(itemOperationsRequest);
            Site.Assert.AreEqual<string>("1", itemOperationsResponse.Status, "The ItemOperations response should be successful.");

            #endregion

            // Get task item that created in this case.
            ItemOperations taskReturnedInItemOperations = null;
            foreach (ItemOperations item in itemOperationsResponse.Items)
            {
                if (task.Task.Body.Data.ToString().Contains(item.Task.Body.Data))
                {
                    taskReturnedInItemOperations = item;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R356");

            // Verify MS-ASTASK requirement: MS-ASTASK_R356
            // If the Task item in response is not null, this requirement will be captured.
            Site.CaptureRequirementIfIsNotNull(
                taskReturnedInItemOperations.Task,
                356,
                @"[In ItemOperations Command Response] When a client uses an ItemOperations command request ([MS-ASCMD] section 2.2.2.8) to retrieve data from the server for one or more specific Task items, as specified in section 3.1.5.1, the server responds with an ItemOperations command response ([MS-ASCMD] section 2.2.2.8).");

            bool otherPropertiesNull = true;

            // Loop to verify if other properties except "Body" are not returned.
            foreach (System.Reflection.PropertyInfo propertyInfo in typeof(Task).GetProperties())
            {
                if (propertyInfo.Name != "Body")
                {
                    object value = propertyInfo.GetValue(taskReturnedInItemOperations.Task, null);
                    if (value != null)
                    {
                        otherPropertiesNull = false;
                        break;
                    }
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R358");

            // Verify MS-ASTASK requirement: MS-ASTASK_R358
            // If the body of the Task item in response is not null, this requirement can be captured.
            Site.CaptureRequirementIfIsTrue(
                otherPropertiesNull,
                358,
                @"[In ItemOperations Command Response] If an itemoperations:Schema element ([MS-ASCMD] section 2.2.3.135) is included in the ItemOperations command request, then the elements returned in the ItemOperations command response MUST be restricted to the elements that were included as child elements of the itemoperations:Schema element in the command request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-ASTASK_R359");

            // Since MS-ASTASK_R358 is captured, this requirement can be captured.
            Site.CaptureRequirement(
                359,
                @"[In ItemOperations Command Response] Top-level Task class elements, as specified in section 2.2, MUST be returned as child elements of the itemoperations:Properties element ([MS-ASCMD] section 2.2.3.128) in the ItemOperations command response.");
        }
    }
}