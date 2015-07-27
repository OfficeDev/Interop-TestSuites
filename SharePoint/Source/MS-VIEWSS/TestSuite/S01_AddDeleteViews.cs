//--------------------------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized
// to use this sample source code. For the terms of the license, please see the
// license agreement between you and Microsoft.
//--------------------------------------------------------------------------------------
namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System;
    using System.Web.Services.Protocols;
    using Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the AddView and DeleteView operations.
    /// </summary>
    [TestClass]
    public class S01_AddDeleteViews : TestSuiteBase
    {
        #region Initialization and clean up
        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }
        #endregion

        #region Test Cases
        /// <summary>
        /// A test case used to test AddView method successfully with valid parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC01_AddView_Success()
        {
            // Call AddView method to add a list view with valid parameters.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool. 
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // If the protocol server creates list view successfully include the resulting View element, then the following requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                addViewResponseAddViewResult.View,
                102,
                @"[In AddViewResponse] It[The protocol server] MUST create the list view, and include the resulting View element when the operation succeeds.");
        }

        /// <summary>
        /// A test case used to test AddView method with an empty viewFields parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC02_AddView_EmptyViewFields()
        {
            // Call AddView method to add a list view with an empty viewFields.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Html.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");            

            // If the protocol server creates list view with no fields included when the value of the viewFields element is empty, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                0,
                addViewResponseAddViewResult.View.ViewFields.Length,
                93,
                @"[In AddView] When the value of the viewFields element is empty, the protocol server MUST create the list view with no fields (2) included.");
        }

        /// <summary>
        /// A test case used to test AddView method with an empty query parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC03_AddView_EmptyQuery()
        {
            // Call AddView method to add a list view with an empty query.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.Query, "There should be a query in the view.");
            this.Site.Assert.IsNull(addViewResponseAddViewResult.View.Query.OrderBy, "The OrderBy clause of Query must be null.");
            this.Site.Assert.IsNull(addViewResponseAddViewResult.View.Query.Where, "The Where clause of Query must be null.");

            // Get the count of the items in the view.
            int itemCountOfEmptyQuery = TestSuiteBase.SutControlAdapter.GetItemsCount(TestSuiteBase.ListGUID, viewName);

            int expectAllItemsCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // If the protocol server returns all the items of the list in the view, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectAllItemsCount,
                itemCountOfEmptyQuery,
                94,
                @"[In AddView] When the value of the query element is empty, the protocol server MUST create the list view without any additional restriction.");            
        }

        /// <summary>
        /// A test case used to test AddView method with an empty rowLimit parameter and a query having Where condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC04_AddView_EmptyRowLimit()
        {
            // Call AddView method to add a list view with an empty rowLimit and a query having Where condition.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.RowLimit, "There should be a RowLimit element in the view.");
            
            // If the protocol server returns the default value of 0x0064 when the rowLimit element is not present, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                0x0064,
                addViewResponseAddViewResult.View.RowLimit.Value,
                9503,
                @"[In AddView] When the value of the rowLimit element is empty, the server MUST use the default value of 0x0064.");

            // If the protocol server returns the list view support page-by-page displaying of items when the rowLimit element is not present, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                bool.TrueString.ToUpper(),
                addViewResponseAddViewResult.View.RowLimit.Paged.ToUpper(),
                9504,
                @"[In AddView] When the value of the rowLimit element is empty, the list view MUST support page-by-page displaying of items.");
        }

        /// <summary>
        /// A test case used to test AddView method with an empty rowLimit parameter and a query without Where condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC05_AddView_EmptyRowLimitWithoutWhere()
        {
            // Call AddView method to add a list view with an empty rowLimit and a query without Where condition.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForGroupBy(true);

            AddViewRowLimit rowLimit = new AddViewRowLimit();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.RowLimit, "There should be the RowLimit element in the view.");
            
            // Call GetItemsCount method to get the count of the list items in the specified view.
            int queryActualItemCount = TestSuiteBase.SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // If the rowLimit is not provided and there is a query condition, the number of view's items should be equal to the number of query items, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCount,
                queryActualItemCount,
                "MS-WSSCAML",
                787,
                @"[In Child Elements] If [the content of RowLimitDefinition is ]not specified, the list schema consumer MUST return all items that meet the filter condition.");
        }

        /// <summary>
        /// A test case used to test AddView method when there are no child elements in LogicalJoinDefinition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC06_AddView_LogicalJoinDefinitionWithoutChild()
        {
            // Call AddView method to add a list view for the specified list on the server when there are no child elements in LogicalJoinDefinition.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            AddViewQuery addViewQuery = new AddViewQuery();            
            addViewQuery.Query = this.GetCamlQueryRoot(Query.EmptyQueryInfo, false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, null, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // Call GetItemsCount method to get the count of the list items in the specified view.
            int itemCountWithoutLogicalJoinDefinition = TestSuiteBase.SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCountWithoutLogicalJoinDefinition = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // If there are no child elements in LogicalJoinDefinition, that is to say the query is an empty query, then the number of view's items should be equal to the number of all items in the list, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCountWithoutLogicalJoinDefinition,
                itemCountWithoutLogicalJoinDefinition,
                "MS-WSSCAML",
                2501,
                @"[In LogicalJoinDefinition Type] When there are no child elements[in the element of LogicalJoinDefinition type], no additional conditions apply to the query.");
        }

        /// <summary>
        /// A test case used to test AddView method when there are child elements in LogicalJoinDefinition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC07_AddView_LogicalJoinDefinitionPresent()
        {
            // Call AddView method to add a list view for the specified list on the server when there are child elements in LogicalJoinDefinition.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(true);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();           

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, null, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // Call GetItemsCount method to get the count of the list items in the specified view.
            int itemCountWithoutLogicalJoinDefinition = TestSuiteBase.SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCountWithLogicalJoinDefinition = int.Parse(Common.GetConfigurationPropertyValue("QueryItemsCount", this.Site));

            // If there is a query condition, the number of view's items should be equal to the number of items queried, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCountWithLogicalJoinDefinition,
                itemCountWithoutLogicalJoinDefinition,
                "MS-WSSCAML",
                25,
                @"[In LogicalJoinDefinition Type] When this element[LogicalJoinDefinition] is present and has child elements, the server MUST return only list items that satisfy the conditions specified by those child elements.");       
        }

        /// <summary>
        /// A test case used to test AddView method with null rowLimit parameter and collapse is set to false in the GroupBy query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC08_AddView_NullRowLimitAndGroupByFalseCollapse()
        {
            // Call AddView method to add a list view with null rowLimit and a query having GroupBy condition and Collapse as false.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRoot(Query.IsNotCollapse, false);

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, null, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // Call GetItemsCount method to get the count of the list items in the specified view.
            int itemNotCollapseCount = TestSuiteBase.SutControlAdapter.GetItemsCount(listName, viewName);

            // When the query's Collapse attribute is false, even there are field values that can be grouped up, the number of view's items is equal to the number of all items in the list, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site)),
                itemNotCollapseCount,
                "MS-WSSCAML",
                67,
                @"[In Attributes] Otherwise[In GroupByDefinition: If Collapse is false], the number of rows in the result set MUST NOT be affected by the GroupBy element.");    
        }

        /// <summary>
        /// A test case used to test AddView method with null rowLimit parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC09_AddView_NullRowLimit()
        {
            // Call AddView method to add a list view with null rowLimit.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, null, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.RowLimit, "There should be a RowLimit element in the view.");

            // If the protocol server returns the default value of 0x0064 when the rowLimit element is not present, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                0x0064,
                addViewResponseAddViewResult.View.RowLimit.Value,
                9501,
                @"[In AddView] When the rowLimit element is not present, the server MUST use the default value of 0x0064.");

            // If the protocol server returns the list view support page-by-page displaying of items when the rowLimit element is not present, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                bool.TrueString.ToUpper(),
                addViewResponseAddViewResult.View.RowLimit.Paged.ToUpper(),
                9502,
                @"[In AddView] When the rowLimit element is not present, the list view MUST support page-by-page displaying of items.");
        }

        /// <summary>
        /// A test case used to test AddView method with makeViewDefault parameter is set true.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC10_AddView_MakeViewDefaultTrue()
        {
            // Call AddView method to add a list view with makeViewDefault set true.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);
            
            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, true);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // If the new default view is added successfully, the original default view lost its default view position.
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.DefaultView, "The response element \"addViewResponseAddViewResult.View.DefaultView\" should not be null.");
            this.Site.Assert.AreEqual("true", addViewResponseAddViewResult.View.DefaultView.ToLower(), "The added view should be a default view.");
            if (TestSuiteBase.OriginalDefaultViewName != null)
            {
                if (TestSuiteBase.OriginalDefaultViewLost == false)
                {
                    TestSuiteBase.OriginalDefaultViewLost = true;
                }
            }

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // If the protocol server create the list view as the default list view when the value of makeViewDefault is set to "true", then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                bool.TrueString.ToUpper(),
                addViewResponseAddViewResult.View.DefaultView.ToUpper(),
                96,
                @"[In AddView] The protocol server MUST create the list view as the default list view if ""true"" [the value of makeViewDefault element] is specified.");
        }

        /// <summary>
        /// A test case used to test AddView method with null type parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC11_AddView_NullType()
        {
            // Call AddView method to add a list view with null type.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, null, true);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // If the new default view is added successfully, the original default view lost its default view position.
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.DefaultView, "The response element \"addViewResponseAddViewResult.View.DefaultView\" should not be null.");
            this.Site.Assert.AreEqual("true", addViewResponseAddViewResult.View.DefaultView.ToLower(), "The added view should be a default view.");
            if (TestSuiteBase.OriginalDefaultViewName != null)
            {
                if (TestSuiteBase.OriginalDefaultViewLost == false)
                {
                    TestSuiteBase.OriginalDefaultViewLost = true;
                }
            }

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // If the protocol server set the value of type is "Html" when it is not present, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                ViewType.Html.ToString().ToUpper(),
                addViewResponseAddViewResult.View.Type.ToString().ToUpper(),
                10001,
                @"[In type] When this element[type] is not present, the protocol server MUST take it with a value of ""Html"".");
        }

        /// <summary>
        /// A test case used to test AddView method with an empty type parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC12_AddView_EmptyType()
        {
            // Call AddView method to add a list view with an empty type.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, string.Empty, true);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // If the new default view is added successfully, the original default view lost its default view position.
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.DefaultView, "The response element \"addViewResponseAddViewResult.View.DefaultView\" should not be null.");
            this.Site.Assert.AreEqual("true", addViewResponseAddViewResult.View.DefaultView.ToLower(), "The added view should be a default view.");
            if (TestSuiteBase.OriginalDefaultViewName != null)
            {
                if (TestSuiteBase.OriginalDefaultViewLost == false)
                {
                    TestSuiteBase.OriginalDefaultViewLost = true;
                }
            }

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // If the protocol server set the value of type is "Html" when it is an empty value, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                ViewType.Html.ToString().ToUpper(),
                addViewResponseAddViewResult.View.Type.ToString().ToUpper(),
                10002,
                @"[In type] When this element[type] has an empty value, the protocol server MUST take it with a value of ""Html"".");
        }

        /// <summary>
        /// A test case used to test AddView method with invalid type parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC13_AddView_InvalidType()
        {
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = this.GenerateRandomString(5);
            type = "a" + type;

            bool caughtSoapException = false;

            // Call AddView method to add a list view with invalid type.
            try
            {
                 Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns an exception when the type element is not empty and is not one of the values Calendar, Grid or Html, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    101,
                    @"[In type] When the value of the element [the type element] is not empty and is not one of the values listed in the table [Calendar, Grid, Html], the protocol server MUST throw a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response.");
        }

        /// <summary>
        /// A test case used to test AddView method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC14_AddView_InvalidListName()
        {
            string invalidListName = this.GenerateRandomString(10);
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();

            bool caughtSoapException = false; 

            // Call AddView method to add a list view with an invalid listName.
            try
            {             
                Adapter.AddView(invalidListName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns an exception when the listName element is not the name or GUID of a list, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    @"[In listName] If the value of listName element is not the name or GUID of a list (1), the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test AddView method without computed fields in viewFields and collapse is set to true in the query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC15_AddView_TrueCollapse_NoComputedFields()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1507, this.Site), @"The test case is executed only when R1507Enabled is set to true.");

            // Call AddView method to add a list view without computed fields in viewFields and collapse is set to true in the query condition.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRootForGroupBy(true);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = ViewType.Grid.ToString();               

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, rowLimit, type, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // Call GetItemsCount method to get the count of the list items in the specified view.
            int itemCollapseCount = SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCollapseCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // When the collapse attribute is true, if the number of view's items is the same as the number of all items in the list, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCollapseCount,
                itemCollapseCount,
                1507,
                @"[In Appendix B: Product Behavior] Implementation does not restrict the number of rows present in the result set to the number of unique tuples where a tuple is a set of field values when there aren't any computed fields in the ViewFields section if Collapse is true.(Windows SharePoint Services 2.0 and above products follow this behavior.)");
        }

        /// <summary>
        /// A test case used to test AddView method without the optional parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC16_AddView_WithoutOptionalParameters()
        {
            // Call AddView method without the optional parameters to add a list view.
            string listName = TestSuiteBase.ListGUID;
            string displayName = this.GenerateRandomString(5);
            AddViewViewFields viewFields = new AddViewViewFields();
            AddViewQuery addViewQuery = new AddViewQuery();

            AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields, addViewQuery, null, null, false);
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult, "The added view should be got successfully.");
            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The server should return a View element that specifies the list view when the AddView method is successful!");

            // Put the newly added view into ViewPool.
            string viewName = addViewResponseAddViewResult.View.Name;
            TestSuiteBase.ViewPool.Add(viewName);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The created list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(displayName, getView.View.DisplayName, "The list view added in the step above should be got successfully!");

            // If there is a View element that specifies the list view is returned, then the following requirement can be captured.
            Site.CaptureRequirement(
                102,
                @"[In AddViewResponse] It[The protocol server] MUST create the list view, and include the resulting View element when the operation succeeds.");
        }

        /// <summary>
        /// A test case used to test DeleteView method successful with valid parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC17_DeleteView_Success()
        {
            // Call AddView method to add a list view for the specified list on the server.           
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);
            string listName = TestSuiteBase.ListGUID;

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);            
            this.Site.Assert.IsNotNull(getView.View, "The list view added in the step above should be got successfully!");

            // Call DeleteView method with valid parameters to delete the list view created above.
            Adapter.DeleteView(listName, viewName);

            bool caughtSoapException = false; 

            // Call GetView method to get the list view deleted above.
            try
            {
                Adapter.GetView(listName, viewName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true;

                // If the protocol server returns an exception that means the specified view does not exist anymore, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    106,
                    @"[In DeleteViewResponse] The protocol server MUST delete the list view and respond with a DeleteViewResponse element if the operation [DeleteView] succeeded.");

                // If ViewPool contains this view, remove the view from ViewPool.
                if (ViewPool.Contains(viewName))
                {
                    ViewPool.Remove(viewName);
                }
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test DeleteView method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC18_DeleteView_InvalidListName()
        {
            // Call AddView method to add a list view for the specified list on the server.           
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);
            string listName = TestSuiteBase.ListGUID;

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);            
            this.Site.Assert.IsNotNull(getView.View, "The list view added in the step above should be got successfully!");

            bool caughtSoapException = false; 

            // Call DeleteView method to delete the list view with an invalid listName.
            try
            {
                string invalidListName = this.GenerateRandomString(10);
                Adapter.DeleteView(invalidListName, viewName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns an exception when the listName element is not the name or GUID of a list, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    @"[In listName] If the value of listName element is not the name or GUID of a list (1), the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test DeleteView method without the optional parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC19_DeleteView_WithoutOptionalParameters()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string defaultViewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call DeleteView method with a null viewName.
            Adapter.DeleteView(listName, null);

            bool caughtSoapException = false; 

            // Call GetView method to get the default list view deleted above.
            try
            {
                Adapter.GetView(listName, defaultViewName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true;

                // If ViewPool contains this default view, delete it from ViewPool.
                if (ViewPool.Contains(defaultViewName))
                {
                    ViewPool.Remove(defaultViewName);
                }

                // If server returns an exception that means the specified view does not exist anymore, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    2301,
                    @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list (1).");

                // If server returns an exception that means the specified view does not exist anymore, then the following requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    106,
                    @"[In DeleteViewResponse] The protocol server MUST delete the list view and respond with a DeleteViewResponse element if the operation [DeleteView] succeeded.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test DeleteView operation to delete the default view when there is no default view in the list.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S01_TC20_DeleteView_NoDefaultView()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string defaultViewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default view.
            Adapter.DeleteView(listName, string.Empty);

            // If ViewPool contains this default view, delete it from ViewPool.
            if (ViewPool.Contains(defaultViewName))
            {
                ViewPool.Remove(defaultViewName);
            }

            bool caughtSoapException = false;

            // Call DeleteView operation to delete the default view with an empty view name, when there is no default view.
            try
            {
                Adapter.DeleteView(listName, string.Empty);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true;

                // If server returns a SOAP fault, then capture the following requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    2302,
                    @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list (1).");

                // If server returns a SOAP fault, then capture the following requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    24,
                    @"[In viewName] If the default list view does not exist, the protocol server MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response.");
        }
        #endregion
    }
}