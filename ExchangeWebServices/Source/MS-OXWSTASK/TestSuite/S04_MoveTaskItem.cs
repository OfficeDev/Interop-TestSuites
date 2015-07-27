//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSTASK
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test move operation.
    /// </summary>
    [TestClass]
    public class S04_MoveTaskItem : TestSuiteBase
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
        /// This test case is used to verify MoveItem operation.
        /// </summary>
        [TestCategory("MSOXWSTASK"), TestMethod()]
        public void MSOXWSTASK_S04_TC01_VerifyMoveTaskItem()
        {
            #region Client calls CreateItem to create a task item on server. By default, it is created in task folder.
            string subject = Common.GenerateResourceName(this.Site, "This is a task");
            ItemIdType[] createItemIds = this.CreateTasks(TestSuiteHelper.DefineTaskItem(subject, null));
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This create response status should be success!", null);
            ItemIdType createItemId = createItemIds[0];
            #endregion

            #region Client calls MoveItem to move the task item to deleteditems folder.
            ItemIdType[] moveItemIds = this.MoveTasks(createItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This move response status should be success!", null);
            ItemIdType moveItemId = moveItemIds[0];
            #endregion

            #region Client calls GetItem to check whether the task item is moved.
            this.GetTasks(createItemId);
            Site.Assert.AreNotEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should not be success! The created task has been moved.", null);
            this.GetTasks(moveItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This get response status should be success!", null);
            #endregion

            #region Client calls DeleteItem to delete the task item created in the previous steps.
            this.DeleteTasks(moveItemId);
            Site.Assert.AreEqual<ResponseClassType>(ResponseClassType.Success, (ResponseClassType)this.ResponseClass[0], "This delete response status should be success!", null);
            #endregion
        }

        #endregion
    }
}