namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System;
    using System.Web.Services.Protocols;
    using Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the UpdateViewHtml2 operation.
    /// </summary>
    [TestClass]
    public class S05_MaintainViewPropertiesWithOpenApplicationExtension : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up
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
        /// A test case used to test UpdateViewHtml2 operation to update the display name of a view, with all input parameters present.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC01_UpdateViewHtml2_Success()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            // Add a view.
            string viewName = this.AddView(false, Query.EmptyQueryInfo, ViewType.Html);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Recursive);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,  
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            // Call UpdateViewHtml2 to update the view including display properties and the possibility to be rendered with extended application.
            // All optional parameters in the request of UpdateViewHtml2 are present in this call.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                       TestSuiteBase.ListGUID,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       toolbar,
                                                                                       viewHeader,
                                                                                       viewBody,
                                                                                       viewFooter,
                                                                                       viewEmpty,
                                                                                       rowLimitExceeded,
                                                                                       queryValue,
                                                                                       viewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue,
                                                                                       openApplicationExtension);

            this.Site.Assert.IsNotNull(updateViewHtml2Re, "There should be a result element returned from the UpdateViewHtml2 operation.");

            // Verify Requirement MS-VIEWSS_R145, if the server returns a View element in the update result.
            Site.CaptureRequirementIfIsNotNull(updateViewHtml2Re.View, 145, @"[In UpdateViewHtml2Response] UpdateViewHtml2Result: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View.DisplayName, "The response element \"updateViewHtml2Re.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(viewProperties.View.DisplayName, updateViewHtml2Re.View.DisplayName, "The display name should be the same as updated.");

            // Verify Requirement MS-VIEWSS_R8016, if the server returns a View element in the update result that indicates the UpdateViewHTML2 operation succeed on the server.
            Site.CaptureRequirementIfIsNotNull(updateViewHtml2Re.View, 8016, @"[In Appendix B: Product Behavior] Implementation does support this method[UpdateViewHtml2]. (Windows SharePoint Services 3.0 and above products follow this behavior.)");
            // Verify requirement: MS-VIEWSS_R8015
            if (Common.IsRequirementEnabled(8015, this.Site))
            {
                string listName = TestSuiteBase.ListGUID;
                string displayName = this.GenerateRandomString(5);
                string type = ViewType.Html.ToString();
                AddViewViewFields viewFields1 = new AddViewViewFields();
                viewFields.ViewFields = this.GetViewFields(true);
                AddViewQuery addViewQuery = new AddViewQuery();
                addViewQuery.Query = this.GetCamlQueryRootForWhere(false);
                AddViewRowLimit rowLimit = new AddViewRowLimit();
                rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

                AddViewResponseAddViewResult addViewResponseAddViewResult = Adapter.AddView(listName, displayName, viewFields1, addViewQuery, rowLimit, type, false);

                UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re2 = Adapter.UpdateViewHtml2(
                                                                                          listName,
                                                                                          viewName,
                                                                                          viewProperties,
                                                                                          toolbar,
                                                                                          viewHeader,
                                                                                          viewBody,
                                                                                          viewFooter,
                                                                                          viewEmpty,
                                                                                          rowLimitExceeded,
                                                                                          queryValue,
                                                                                          viewFields,
                                                                                          aggregations,
                                                                                          formats,
                                                                                          rowLimitValue,
                                                                                          openApplicationExtension);
                GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
                Site.CaptureRequirementIfAreEqual(
                   openApplicationExtension,
                    getView.View.Text[0],
                    8015,
                    @"[In Appendix B: Product Behavior] Implementation does return the value of OpenApplicationExtension as the value of the View element when GetView is called after UpdateViewHtml2 (section 3.1.4.8) and the type of the view is HTML. <3> Section 3.1.4.3.2.2: In SharePoint Foundation 2010 and SharePoint Foundation 2013, when this method is called after UpdateViewHtml2 (section 3.1.4.8) and the type of the view is HTML, the value of OpenApplicationExtension is returned as the value of the View element. (SharePoint Server 2010 and above follow this hebavior.)");
            }
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation concerning the LogicalJoin query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC02_UpdateViewHtml2_LogicalJoin()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            // Add a view.
            string viewName = this.AddView(false, Query.EmptyQueryInfo, ViewType.Html);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Item);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();

            // Make the query having LogicalJoinDefinition element with available child in it.
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            // Call UpdateViewHtml2 to update the view with available LogicalJoinDefinition query condition.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                       TestSuiteBase.ListGUID,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       toolbar,
                                                                                       viewHeader,
                                                                                       viewBody,
                                                                                       viewFooter,
                                                                                       viewEmpty,
                                                                                       rowLimitExceeded,
                                                                                       queryValue,
                                                                                       viewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue,
                                                                                       openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");

            // Get the count of the items in the view.
            int itemCountInTheView = TestSuiteBase.SutControlAdapter.GetItemsCount(TestSuiteBase.ListGUID, viewName);

            int expectIQueryItemsCount = int.Parse(Common.GetConfigurationPropertyValue("QueryItemsCount", this.Site));

            // Verify Requirement MS-WSSCAML_R25, if the item count in the view returned from the server is the same with the query expectation.
            Site.CaptureRequirementIfAreEqual(expectIQueryItemsCount, itemCountInTheView, "MS-WSSCAML", 25, @"[In LogicalJoinDefinition Type] When this element[LogicalJoinDefinition] is present and has child elements, the server MUST return only list items that satisfy the conditions specified by those child elements.");

            // Make the query having no child element in LogicalJoinDefinition, an empty query.
            queryValue.Query = this.GetCamlQueryRoot(Query.EmptyQueryInfo, false);

            // Call UpdateViewHtml2 to update the view with an empty query condition.
            updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                          TestSuiteBase.ListGUID,
                                          viewName,
                                          viewProperties,
                                          toolbar,
                                          viewHeader,
                                          viewBody,
                                          viewFooter,
                                          viewEmpty,
                                          rowLimitExceeded,
                                          queryValue,
                                          viewFields,
                                          aggregations,
                                          formats,
                                          rowLimitValue,
                                          openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");

            // Get the count of the items in the view.
            itemCountInTheView = TestSuiteBase.SutControlAdapter.GetItemsCount(TestSuiteBase.ListGUID, viewName);
            int expectAllItemsCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // Verify Requirement MS-WSSCAML_R2501, if the item count in the view returned from the server is the same with the count of all list items.
            Site.CaptureRequirementIfAreEqual(expectAllItemsCount, itemCountInTheView, "MS-WSSCAML", 2501, @"[In LogicalJoinDefinition Type] When there are no child elements[in the element of LogicalJoinDefinition type], no additional conditions apply to the query.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation concerning the GroupBy and Collapse query conditions.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC03_UpdateViewHtml2_GroupByAndCollapse()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1507, this.Site), @"The test case is executed only when R1507Enabled is set to true.");

            // Add a view.
            string viewName = this.AddView(false, Query.EmptyQueryInfo, ViewType.Grid);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.RecursiveAll);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();

            // Make the query have GroupBy condition and Collapse as false.
            queryValue.Query = this.GetCamlQueryRoot(Query.IsNotCollapse, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            // Call UpdateViewHtml2 to update the view with a query having GroupBy condition and Collapse as false.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                       TestSuiteBase.ListGUID,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       toolbar,
                                                                                       viewHeader,
                                                                                       viewBody,
                                                                                       viewFooter,
                                                                                       viewEmpty,
                                                                                       rowLimitExceeded,
                                                                                       queryValue,
                                                                                       viewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue,
                                                                                       openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");

            // Get the count of the items in the view.
            int itemGroupByCount = TestSuiteBase.SutControlAdapter.GetItemsCount(TestSuiteBase.ListGUID, viewName);

            int expectAllItemsCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // Verify Requirement MS-WSSCAML_R67, if the item count in the view returned from the server is the same with the count of all list items.
            Site.CaptureRequirementIfAreEqual(expectAllItemsCount, itemGroupByCount, "MS-WSSCAML", 67, @"[In Attributes] Otherwise[In GroupByDefinition: If Collapse is false], the number of rows in the result set MUST NOT be affected by the GroupBy element.");

            // Make the query have GroupBy condition and Collapse as true, while the referenced view fields have no computed fields.
            queryValue.Query = this.GetCamlQueryRoot(Query.IsCollapse, false);

            // Call UpdateViewHtml2 to update the view with a query having GroupBy condition and Collapse as true, while the referenced view fields have no computed fields.
            updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                        TestSuiteBase.ListGUID,
                                                        viewName,
                                                        viewProperties,
                                                        toolbar,
                                                        viewHeader,
                                                        viewBody,
                                                        viewFooter,
                                                        viewEmpty,
                                                        rowLimitExceeded,
                                                        queryValue,
                                                        viewFields,
                                                        aggregations,
                                                        formats,
                                                        rowLimitValue,
                                                        openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");

            int itemCollapseCount = TestSuiteBase.SutControlAdapter.GetItemsCount(TestSuiteBase.ListGUID, viewName);
            int expectItemCollapseCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // Verify Requirement MS-VIEWSS_R1507, if the item count in the view returned from the server is the same with the count of all items in the list.
            Site.CaptureRequirementIfAreEqual<int>(expectItemCollapseCount, itemCollapseCount, 1507, @"[In Appendix B: Product Behavior] Implementation does not restrict the number of rows present in the result set to the number of unique tuples where a tuple is a set of field values when there aren't any computed fields in the ViewFields section if Collapse is true.(Windows SharePoint Services 2.0 and above products follow this behavior.)");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation with view name not present.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC04_UpdateViewHtml2_NullViewName()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.FilesOnly);                        
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);
            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            // Call UpdateViewHtml2 operation with view name not present.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                   listName,
                                                                                   null,
                                                                                   viewProperties,
                                                                                   toolbar,
                                                                                   viewHeader,
                                                                                   viewBody,
                                                                                   viewFooter,
                                                                                   viewEmpty,
                                                                                   rowLimitExceeded,
                                                                                   queryValue,
                                                                                   viewFields,
                                                                                   aggregations,
                                                                                   formats,
                                                                                   rowLimitValue,
                                                                                   openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View.DefaultView, "The response element \"updateViewHtml2Re.View.DefaultView\" should not be null.");
            bool isDefaultView = Convert.ToBoolean(updateViewHtml2Re.View.DefaultView);

            // Verify Requirement MS-VIEWSS_R2301, if the server returns the default view back.
            Site.CaptureRequirementIfIsTrue(isDefaultView, 2301, @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation with an empty view name.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC05_UpdateViewHtml2_EmptyViewName()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Recursive);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.IsNotCollapse, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            // Call UpdateViewHtml2 operation with empty view name.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                       listName,
                                                                                       string.Empty,
                                                                                       viewProperties,
                                                                                       toolbar,
                                                                                       viewHeader,
                                                                                       viewBody,
                                                                                       viewFooter,
                                                                                       viewEmpty,
                                                                                       rowLimitExceeded,
                                                                                       queryValue,
                                                                                       viewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue,
                                                                                       openApplicationExtension);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View.DefaultView, "The response element \"updateViewHtml2Re.View.DefaultView\" should not be null.");
            bool isDefaultView = Convert.ToBoolean(updateViewHtml2Re.View.DefaultView);

            // Verify Requirement MS-VIEWSS_R2302, if the server returns the default view back.
            Site.CaptureRequirementIfIsTrue(isDefaultView, 2302, @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation with the least input parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC06_UpdateViewHtml2_Success_LeastInputParameters()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call UpdateViewHtml2 to get the default view with display properties by giving only one input parameter, the list name.
            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Re = Adapter.UpdateViewHtml2(
                                                                                       TestSuiteBase.ListGUID,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null);
            this.Site.Assert.IsNotNull(updateViewHtml2Re, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View, "There should be a view element returned from a successful UpdateViewHtml2 operation.");

            // Verify Requirement MS-VIEWSS_R145, if the server returns a View element in the update result.
            Site.CaptureRequirementIfIsNotNull(updateViewHtml2Re.View, 145, @"[In UpdateViewHtml2Response] UpdateViewHtml2Result: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");

            // Verify Requirement MS-VIEWSS_R8016, if the server returns a View element in the update result that indicates the UpdateViewHTML2 operation succeed on the server.
            Site.CaptureRequirementIfIsNotNull(updateViewHtml2Re.View, 8016, @"[In Appendix B: Product Behavior] Implementation does support this method[UpdateViewHtml2]. (Windows SharePoint Services 3.0 and above products follow this behavior.)");
            this.Site.Assert.IsNotNull(updateViewHtml2Re.View.DefaultView, "The response element \"updateViewHtml2Re.View.DefaultView\" should not be null.");
            bool isDefaultView = Convert.ToBoolean(updateViewHtml2Re.View.DefaultView);

            // Verify Requirement MS-VIEWSS_R2301, if the server returns the default view back.
            Site.CaptureRequirementIfIsTrue(isDefaultView, 2301, @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation with invalid list name.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC07_UpdateViewHtml2_InvalidListName()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            // Add a view.
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Html);

            string listName = this.GenerateRandomString(8);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Recursive);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            bool caughtSoapException = false;

            try
            {
                // Call UpdateViewHtml2 operation with invalid list name.
                Adapter.UpdateViewHtml2(
                    listName,
                    viewName,
                    viewProperties,
                    toolbar,
                    viewHeader,
                    viewBody,
                    viewFooter,
                    viewEmpty,
                    rowLimitExceeded,
                    queryValue,
                    viewFields,
                    aggregations,
                    formats,
                    rowLimitValue,
                    openApplicationExtension);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // Verify Requirement MS-VIEWSS_R13, if the server returns a SOAP fault message.
                Site.CaptureRequirementIfIsNotNull(soapException, 13, @"[In listName] If the value of listName element is not the name or GUID of a list, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation with invalid view name.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC08_UpdateViewHtml2_InvalidViewName()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            string invalidViewName = this.GenerateRandomString(8);

            string listName = TestSuiteBase.ListGUID;

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Recursive);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            bool caughtSoapException = false;

            try
            {
                // Call UpdateViewHtml2 operation with invalid view name.
                Adapter.UpdateViewHtml2(
                    listName,
                    invalidViewName,
                    viewProperties,
                    toolbar,
                    viewHeader,
                    viewBody,
                    viewFooter,
                    viewEmpty,
                    rowLimitExceeded,
                    queryValue,
                    viewFields,
                    aggregations,
                    formats,
                    rowLimitValue,
                    openApplicationExtension);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true;

                // Verify Requirement MS-VIEWSS_R22, if the server returns a SOAP fault message.
                Site.CaptureRequirementIfIsNotNull(soapException, 22, @"[In viewName] If the value of viewName element is not the GUID of a list view, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml2 operation to update the default view when the default view does not exist.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S05_TC09_UpdateViewHtml2_NoDefaultView()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(8016, this.Site), @"The test case is executed only when R8016Enabled is set to true.");

            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string viewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default view.
            this.DeleteView(viewName);

            string updatedDisplayName = this.GenerateRandomString(6);
            UpdateViewHtml2ViewProperties viewProperties = this.GetViewProperties(false, updatedDisplayName, false, ViewScope.Recursive);
            UpdateViewHtml2Toolbar toolbar;
            UpdateViewHtml2ViewHeader viewHeader;
            UpdateViewHtml2ViewBody viewBody;
            UpdateViewHtml2ViewFooter viewFooter;
            UpdateViewHtml2ViewEmpty viewEmpty;
            UpdateViewHtml2RowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRoot(Query.AvailableQueryInfo, false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations = new UpdateViewHtml2Aggregations();
            string aggregationsType = Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site);
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, aggregationsType);

            UpdateViewHtml2Formats formats = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            string openApplicationExtension = Common.GetConfigurationPropertyValue("OpenApplicationExtension", this.Site);

            bool caughtSoapException = false; 

            try
            {
                // Call UpdateViewHtml2 operation with an empty view name, expecting to update the default view.
                Adapter.UpdateViewHtml2(
                    listName,
                    string.Empty,
                    viewProperties,
                    toolbar,
                    viewHeader,
                    viewBody,
                    viewFooter,
                    viewEmpty,
                    rowLimitExceeded,
                    queryValue,
                    viewFields,
                    aggregations,
                    formats,
                    rowLimitValue,
                    openApplicationExtension);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // Verify Requirement MS-VIEWSS_R24, if the server returns a SOAP fault message.
                Site.CaptureRequirementIfIsNotNull(soapException, 24, @"[In viewName] If the default list view does not exist, the protocol server MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 

            caughtSoapException = false;

            try
            {
                // Call UpdateViewHtml2 operation with view name not present, expecting to update the default view.
                Adapter.UpdateViewHtml2(
                    listName,
                    null,
                    viewProperties,
                    toolbar,
                    viewHeader,
                    viewBody,
                    viewFooter,
                    viewEmpty,
                    rowLimitExceeded,
                    queryValue,
                    viewFields,
                    aggregations,
                    formats,
                    rowLimitValue,
                    openApplicationExtension);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // Verify Requirement MS-VIEWSS_R24, if the server returns a SOAP fault message.
                Site.CaptureRequirementIfIsNotNull(soapException, 24, @"[In viewName] If the default list view does not exist, the protocol server MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }
        #endregion

        #region Private Method
        /// <summary>
        /// Get View Properties to update view.
        /// </summary>
        /// <param name="defaultView">Represents whether this is a default view.</param>
        /// <param name="displayName">Represents the display name.</param>
        /// <param name="fpmodified">Represents the whether the view has been modified by a client application.</param>
        /// <param name="viewScope">Represents whether and how files and sub-folders are included in a view.</param>
        /// <returns>The instance of UpdateViewHtml2ViewProperties.</returns>
        protected UpdateViewHtml2ViewProperties GetViewProperties(bool defaultView, string displayName, bool fpmodified, ViewScope viewScope)
        {
            UpdateViewHtml2ViewProperties update2ViewProperties = new UpdateViewHtml2ViewProperties();
            update2ViewProperties.View = new UpdateViewPropertiesDefinition();
            update2ViewProperties.View.DefaultView = defaultView.ToString();
            update2ViewProperties.View.DisplayName = displayName;
            update2ViewProperties.View.FPModified = fpmodified.ToString();
            update2ViewProperties.View.Scope = viewScope;

            return update2ViewProperties;
        }
        #endregion
    }
}