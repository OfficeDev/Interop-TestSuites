namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to operate multiple task items.
    /// </summary>
    [TestClass]
    public class S05_OperateMultipleTaskItems : TestSuiteBase
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
        /// This test case is used to verify the server behavior when operating multiple task objects at the same time.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S05_TC01_OperateMultipleTaskItems()
        {
            #region Client calls CreateItem to create two task items on server.
            string firstSubject = Common.GenerateResourceName(this.Site, "This is a task", 1);
            string secondSubject = Common.GenerateResourceName(this.Site, "This is a task", 2);
            ItemIdType[] createItemIds = this.CreateTasks(TestSuiteHelper.DefineTaskItem(firstSubject), TestSuiteHelper.DefineTaskItem(secondSubject));
            Site.Assert.IsNotNull(createItemIds, "This create response should have task item id!");
            Site.Assert.AreEqual<int>(2, createItemIds.Length, "There should be 2 task items' ids in response!");
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This create response status of second task item should be success!", null);
            #endregion

            #region Client calls GetItem to get two task items.
            TaskType[] taskItems = this.GetTasks(createItemIds);
            Site.Assert.IsNotNull(taskItems, "This get response should have task items!");
            Site.Assert.AreEqual<int>(2, taskItems.Length, "There should be 2 task items in response!");
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This get response status of second task item should be success!", null);
            #endregion

            #region Client calls UpdateItem to update the value of taskCompanies element of task item.
            ItemIdType[] updateItemIds = this.UpdateTasks(createItemIds);
            Site.Assert.IsNotNull(updateItemIds, "This update response should have task item id!");
            Site.Assert.AreEqual<int>(2, updateItemIds.Length, "There should be 2 task items' ids in response!");
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This update response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This update response status of second task item should be success!", null);
            #endregion

            #region Client calls CopyItem to copy the two task items.
            ItemIdType[] copyItemIds = this.CopyTasks(updateItemIds);
            Site.Assert.IsNotNull(copyItemIds, "This copy response should have task item id!");
            Site.Assert.AreEqual<int>(2, copyItemIds.Length, "There should be 2 task items' ids in response!");
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This copy response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This copy response status of second task item should be success!", null);
            #endregion

            #region Client calls MoveItem to move the task items to deleteditems folder
            ItemIdType[] moveItemIds = this.MoveTasks(updateItemIds);
            Site.Assert.IsNotNull(moveItemIds, "This move response should have task item id!");
            Site.Assert.AreEqual<int>(2, moveItemIds.Length, "There should be 2 task items' ids in response!");
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This move response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This move response status of second task item should be success!", null);
            #endregion

            #region Client calls DeleteItem to delete the task items created in the previous steps.
            ItemIdType[] deleteItemIds = new ItemIdType[copyItemIds.Length + moveItemIds.Length];
            copyItemIds.CopyTo(deleteItemIds, 0);
            moveItemIds.CopyTo(deleteItemIds, copyItemIds.Length);
            this.DeleteTasks(deleteItemIds);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status of first task item should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This delete response status of second task item should be success!", null);
            #endregion
        }
        #endregion
    }
}