//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_WSSREST
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is aimed to implement the multiple operations contained in one HTTP Request to do the retrieving, inserting, updating and deleting on list Item.
    /// </summary>
    [TestClass]
    public class S03_BatchRequests : TestSuiteBase
    {
        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">VSTS test context.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }

        #endregion Test Suite Initialization

        /// <summary>
        /// This test case is used to implement the multiple operations contained in one HTTP Request to do the retrieving, inserting, updating and deleting on list Item.
        /// </summary>
        [TestCategory("MSWSSREST"), TestMethod]
        public void MSWSSREST_S03_TC01_BatchRequests()
        {
            Dictionary<string, string> properties = new Dictionary<string, string>();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "title"));

            // Insert two list items to list
            BatchRequest requestFirst = new BatchRequest();
            requestFirst.OperationType = OperationType.Insert;
            requestFirst.Parameter = this.GeneralListName;
            requestFirst.Content = this.GenerateContent(properties);
            requestFirst.ContentType = "application/atom+xml";
            Entry insertResult1 = this.Adapter.InsertListItem(requestFirst);
            Site.Assert.IsNotNull(insertResult1, "Verify insertResult1 is not null!");
            Entry insertResult2 = this.Adapter.InsertListItem(requestFirst);
            Site.Assert.IsNotNull(insertResult2, "Verify insertResult2 is not null!");

            // The second batch request that operation type is Update.
            BatchRequest requestSecond = new BatchRequest();
            requestSecond.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult1.Properties[Constants.IDFieldName]);
            requestSecond.OperationType = OperationType.Update;
            requestSecond.UpdateMethod = UpdateMethod.MERGE;
            properties.Clear();
            properties.Add(Constants.TitleFieldName, Common.GenerateResourceName(this.Site, "update"));
            requestSecond.Content = this.GenerateContent(properties);
            requestSecond.ETag = insertResult1.Etag;
            requestSecond.ContentType = "application/atom+xml";

            // The third batch request that operation type is Delete.
            BatchRequest requestThird = new BatchRequest();
            requestThird.OperationType = OperationType.Delete;
            requestThird.Parameter = string.Format("{0}({1})", this.GeneralListName, insertResult2.Properties[Constants.IDFieldName]);
            requestThird.ETag = insertResult2.Etag;

            string batchResult = this.Adapter.BatchRequests(new List<BatchRequest>() { requestFirst, requestSecond, requestThird });
            Site.Assert.IsNotNull(batchResult, "Verify batchResult is not null");

            // Delete all list items in GeneralList.
            this.DeleteListItems(this.GeneralListName);
        }
    }
}