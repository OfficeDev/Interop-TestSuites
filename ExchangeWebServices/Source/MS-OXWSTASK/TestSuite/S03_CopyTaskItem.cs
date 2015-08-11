namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test copy operation.
    /// </summary>
    [TestClass]
    public class S03_CopyTaskItem : TestSuiteBase
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
        /// This test case is used to verify CopyItem operation.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S03_TC01_VerifyCopyTaskItem()
        {
            #region Client calls CreateItem to create a task item on server.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIds = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, null));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls CopyItem to copy the task item.
            ItemIdType[] copyItemIds = this.CopyTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This copy response status should be success!", null);
            ItemIdType copyItemId = copyItemIds[0];
            #endregion

            #region Client calls DeleteItem to delete the created and copied task items.
            this.DeleteTasks(createItemId, copyItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[1], "This delete response status should be success!", null);
            #endregion
        }
        #endregion
    }
}