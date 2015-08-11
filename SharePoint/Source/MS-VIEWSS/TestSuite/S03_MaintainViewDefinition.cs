namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System.Web.Services.Protocols;
    using Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the UpdateView and GetView operations. 
    /// </summary>
    [TestClass]
    public class S03_MaintainViewDefinition : TestSuiteBase
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

        #region Test cases
        /// <summary>
        /// A test case used to test UpdateView method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC01_UpdateView_InvalidListName()
        {
            // Call AddView method to add a list view for the specified list on the server.           
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(TestSuiteBase.ListGUID, viewName);
            this.Site.Assert.IsNotNull(getView.View, "The list view added in the step above should be gotten successfully!");

            string invalidListName = this.GenerateRandomString(10);
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();
            
            bool caughtSoapException = false; 

            // Call UpdateView method to update the list view added above with an invalid listName.
            try
            {
                Adapter.UpdateView(
                        invalidListName,
                        viewName,
                        viewProperties,
                        updateViewQuery,
                        updateViewFields,
                        aggregations,
                        formats,
                        updateRowLimit);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns an exception when the listName element is not the name or GUID of a list, then capture below requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    @"[In listName] If the value of listName element is not the name or GUID of a list (1), the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateView method with invalid viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC02_UpdateView_InvalidViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            string invalidViewName = this.GenerateRandomString(10);
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            bool caughtSoapException = false; 

            // Call UpdateView method with invalid viewName.
            try
            {
                Adapter.UpdateView(
                                    listName,
                                    invalidViewName,
                                    viewProperties,
                                    updateViewQuery,
                                    updateViewFields,
                                    aggregations,
                                    formats,
                                    updateRowLimit);
            }
            catch (SoapException soapException) 
            {
                caughtSoapException = true;

                // If server returns an exception when the viewName element is not the name or GUID of a list, then capture below requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    22,
                    @"[In viewName] If the value of viewName element is not the GUID of a list view, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test GetView method with null viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC03_GetView_NullViewName()
        {         
            // Call GetViewCollection to see if there is a default list view in the server.
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call GetView method with a null viewName.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, null);
            this.Site.Assert.IsNotNull(getView, "The list view got with null viewName parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DefaultView, "The response element \"getView.View.DefaultView\" should not be null.");

            // If server refers to the default list view of the list, then the below requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                "true",
                getView.View.DefaultView.ToLower(),
                2301,
                @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list (1).");
        }

        /// <summary>
        /// A test case used to test GetView method with an empty viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC04_GetView_EmptyViewName()
        {            
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call GetView method with an empty viewName.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, string.Empty);
            this.Site.Assert.IsNotNull(getView, "The list view got with an empty viewName parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DefaultView, "The response element \"getView.View.DefaultView\" should not be null.");

            // If server refers to the default list view of the list, then the below requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                "true",
                getView.View.DefaultView.ToLower(),
                2302,
                @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list (1).");
        }

        /// <summary>
        /// A test case used to test GetView method with an empty viewName parameter when the default list view does not exist.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC05_GetView_EmptyViewName_NoDefaultView()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string defaultViewName = this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default list view.
            this.DeleteView(defaultViewName);

            bool caughtSoapException = false; 

            try
            {
                // Call protocol adapter method GetView with an empty viewName.
                Adapter.GetView(listName, string.Empty);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns a SOAP fault, then capture below requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    24,
                    @"[In viewName] If the default list view does not exist, the protocol server MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test GetView method successful with valid parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC06_GetView_Success()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string listName = TestSuiteBase.ListGUID;           
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);

            // Call GetView method to get the list view created above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The list view got with valid parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The protocol server should return a View element that contains the details of the specified list view!");
            this.Site.Assert.IsNotNull(getView.View.Name, "The response element \"getView.View.Name\" should not be null.");
            this.Site.Assert.AreEqual(viewName, getView.View.Name, "Server should return the list view created above!");

            // If server returns a View element that contains the details of the specified list view in the response, then below requirement can be captured.
            Site.CaptureRequirement(
                109,
                @"[In GetViewResponse] The protocol server MUST return a View element [in the GetViewResult element] that contains the details of the specified list view when the operation [GetView] succeeds.");            
        }

        /// <summary>
        /// A test case used to test UpdateView method to update the FPModified field successfully with all input parameters present.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC07_UpdateView_AllParameters()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string listName = TestSuiteBase.ListGUID;       
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);

            // Call UpdateView method to update the FPModified.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.FPModified = "true";

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                       listName,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       updateViewQuery,
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");

            // Call GetView method to get the list view updated above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.FPModified, "The response element \"getView.View.FPModified\" should not be null.");
            this.Site.Assert.AreEqual("true", getView.View.FPModified.ToLower(), "The updated FPModified should be gotten successfully!");

            // If the protocol server updates the list view successfully, and returns a View element that specifies the list view, then the following requirement can be captured.
            Site.CaptureRequirement(
                121,
                @"[In UpdateViewResponse] UpdateViewResult: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");
        }

        /// <summary>
        /// A test case used to test UpdateView method without computed fields and collapse is set to true in the query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC08_UpdateView_TrueCollapse_NoComputedFields()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1507, this.Site), @"The test case is executed only when R1507Enabled is set to true.");
            
            // Call AddView method to add a list view for the specified list on the server.
            string listName = TestSuiteBase.ListGUID;
            string viewName = this.AddView(false, Query.IsCollapse, ViewType.Grid);

            // Call UpdateView method to update the display name of the list view created above.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForGroupBy(true);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                        listName,
                                                                                        viewName,
                                                                                        viewProperties,
                                                                                        updateViewQuery,
                                                                                        updateViewFields,
                                                                                        aggregations,
                                                                                        formats,
                                                                                        updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");

            // Call GetView method to get the list view updated above.
            GetViewResponseGetViewResult getView = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getView, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(getView.View, "The response element \"getView.View\" should not be null.");
            this.Site.Assert.IsNotNull(getView.View.DisplayName, "The response element \"getView.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(viewProperties.View.DisplayName, getView.View.DisplayName, "The fields in the list view updated in the steps above should be got successfully!");

            // Call the SUT control adapter method GetItemsCount to get the count of the list items in the specified view.
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
        /// A test case used to test UpdateView method when LogicalJoinDefinition is present and not empty in the query.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC09_UpdateView_LogicalJoinDefinitionPresent()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string listName = TestSuiteBase.ListGUID;        
            string viewName = this.AddView(false, Query.EmptyQueryInfo, ViewType.Grid);
           
            // Call UpdateView method to update the query condition when LogicalJoinDefinition is present in the query and its child element is "Eq".
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(true);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                       listName,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       updateViewQuery,
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");

            // Call SUT control adapter method GetItemsCount to get the count of the list items in the specified view.
            int itemCountWithLogicalJoinDefinition = SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCountWithLogicalJoinDefinition = int.Parse(Common.GetConfigurationPropertyValue("QueryItemsCount", this.Site));
            
            // If there is a query condition, the number of view's items should be equal to the number of items queried, then MS-WSSCAML_R25 can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCountWithLogicalJoinDefinition,
                itemCountWithLogicalJoinDefinition,
                "MS-WSSCAML",
                25,
                @"[In LogicalJoinDefinition Type] When this element[LogicalJoinDefinition] is present and has child elements, the server MUST return only list items that satisfy the conditions specified by those child elements.");
        }

        /// <summary>
        /// A test case used to test UpdateView method when there are no child elements in LogicalJoinDefinition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC10_UpdateView_LogicalJoinDefinitionWithoutChild()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);
            string listName = TestSuiteBase.ListGUID;
          
            // Call UpdateView method to update the query condition when there are no child elements in LogicalJoinDefinition.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForGroupBy(true);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                       listName,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       updateViewQuery,
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");

            // Call SUT control adapter method GetItemsCount to get the count of the list items in the specified view.
            int itemCountWithoutLogicalJoinDefinition = SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCountWithoutLogicalJoinDefinition = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));
           
            // If there are no child elements in LogicalJoinDefinition, that is to say the query is empty query, then the number of view's items should be equal to the number of all items in the list, then MS-WSSCAML_R2501 can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCountWithoutLogicalJoinDefinition,
                itemCountWithoutLogicalJoinDefinition,
                "MS-WSSCAML",
                2501,
                @"[In LogicalJoinDefinition Type] When there are no child elements[in the element of LogicalJoinDefinition type], no additional conditions apply to the query.");
        }

        /// <summary>
        /// A test case used to test UpdateView method when collapse is set to false in the query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC11_UpdateView_FalseCollapse()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.IsNotCollapse, ViewType.Grid);
            string listName = TestSuiteBase.ListGUID;

            // Call UpdateView method to update the list view to make it as default view.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DefaultView = "TRUE";

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForGroupBy(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                       listName,
                                                                                       viewName,
                                                                                       viewProperties,
                                                                                       updateViewQuery,
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");

            // Call GetView method to get the list view updated above.
            GetViewResponseGetViewResult getViewAgain = Adapter.GetView(listName, viewName);
            this.Site.Assert.IsNotNull(getViewAgain, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(getViewAgain.View, "The response element \"getViewAgain.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewAgain.View.DefaultView, "The response element \"getViewAgain.View.DefaultView\" should not be null.");
            this.Site.Assert.AreEqual("true", getViewAgain.View.DefaultView.ToLower(), "The list view should be updated to default view!");

            // If the view is updated into a default view successfully, the original default view lost its default view position.
            if (TestSuiteBase.OriginalDefaultViewName != null)
            {
                if (TestSuiteBase.OriginalDefaultViewLost == false)
                {
                    TestSuiteBase.OriginalDefaultViewLost = true;
                }
            }

            // Call SUT control adapter method GetItemsCount to get the count of the list items in the specified view.
            int itemCollapseCount = SutControlAdapter.GetItemsCount(listName, viewName);
            int expectItemCollapseCount = int.Parse(Common.GetConfigurationPropertyValue("AllItemsCount", this.Site));

            // When the collapse attribute is false, even there are field values that can be grouped up, the number of view's items is the same as the number of all items in the list, then the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                expectItemCollapseCount,
                itemCollapseCount,
                "MS-WSSCAML",
                67,
                @"[In Attributes] Otherwise[In GroupByDefinition: If Collapse is false], the number of rows in the result set MUST NOT be affected by the GroupBy element.");
        }

        /// <summary>
        /// A test case used to test GetView method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC12_GetView_InvalidListName()
        {
            // Call AddView method to add a list view for the specified list on the server.            
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Grid);

            bool caughtSoapException = false; 

            // Call GetView method with invalid listName to get the list view added above.
            try
            {
                GetViewResponseGetViewResult getView = Adapter.GetView(this.GenerateRandomString(5), viewName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If the server returns SOAP fault, then capture this requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    "[In listName] If the value of listName element is not the name or GUID of a list (1), the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test GetView method with invalid viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC13_GetView_InvalidViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            bool caughtSoapException = false; 

            // Call GetView method with invalid viewName.
            try
            {
                Adapter.GetView(listName, this.GenerateRandomString(10));
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If the server returns SOAP fault, then capture this requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    22,
                    "[In viewName] If the value of viewName element is not the GUID of a list view, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateView method with null viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC14_UpdateView_NullViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call UpdateView method with a null viewName.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                        listName,
                                                                                        null,
                                                                                        viewProperties,
                                                                                        updateViewQuery,
                                                                                        updateViewFields,
                                                                                        aggregations,
                                                                                        formats,
                                                                                        updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds.");
            this.Site.Assert.IsNotNull(updateViewResult.View.DefaultView, "A \"View.DefaultView\" element should exist in the server response.");                   

            // If the protocol server refers to the default list view of the list when viewName element is not present, then capture below requirement.
            Site.CaptureRequirementIfAreEqual(
                "true",
                updateViewResult.View.DefaultView.ToLower(),
                2301,
                @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list (1).");
        }

        /// <summary>
        /// A test case used to test UpdateView method with an empty viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC15_UpdateView_EmptyViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);
             
            // Call UpdateView method to update the display name of the view with empty viewName.
            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                        listName,
                                                                                        string.Empty,
                                                                                        viewProperties,
                                                                                        updateViewQuery,
                                                                                        updateViewFields,
                                                                                        aggregations,
                                                                                        formats,
                                                                                        updateRowLimit);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");
            this.Site.Assert.IsNotNull(updateViewResult.View.DefaultView, "A \"View.DefaultView\" element should exist in the server response."); 

            // If the protocol server refers to the default list view of the list when viewName element is empty, then capture below requirement.
            Site.CaptureRequirementIfAreEqual(
                "true",
                updateViewResult.View.DefaultView.ToLower(),
                2302,
                @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list (1).");
        }

        /// <summary>
        /// A test case used to test UpdateView method with an empty viewName parameter when the default list view does not exist.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC16_UpdateView_EmptyViewName_NoDefaultView()
        {
            string listName = TestSuiteBase.ListGUID;
            
            // Add a default view.
            string defaultViewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default view.
            this.DeleteView(defaultViewName);

            UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewQuery updateViewQuery = new UpdateViewQuery();
            updateViewQuery.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewViewFields updateViewFields = new UpdateViewViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewAggregations aggregations = new UpdateViewAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewFormats formats = new UpdateViewFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewRowLimit updateRowLimit = new UpdateViewRowLimit();
            updateRowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            bool caughtSoapException = false; 

            // Call UpdateView method to update the display name of the view with an empty viewName.
            try
            {
                Adapter.UpdateView(
                                    listName,
                                    string.Empty,
                                    viewProperties,
                                    updateViewQuery,
                                    updateViewFields,
                                    aggregations,
                                    formats,
                                    updateRowLimit);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If server returns an exception when the default list view does not exist, then capture below requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    24,
                    @"[In viewName] If the default list view does not exist, the protocol server MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateView method without the optional parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S03_TC17_UpdateView_WithoutOptionalParameters()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);
                 
            // Call UpdateView method without optional parameters.               
            UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                        listName,
                                                                                        null,
                                                                                        null,
                                                                                        null,
                                                                                        null,
                                                                                        null,
                                                                                        null,
                                                                                        null);
            this.Site.Assert.IsNotNull(updateViewResult, "The updated list view should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewResult.View, "The server should return a View element that specifies the list view when the UpdateView method succeeds!");
                                 
            // If there is a View element that specifies the list view is returned, then capture below requirement.
            Site.CaptureRequirement(
                121,
                @"[In UpdateViewResponse] UpdateViewResult: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");
        }

        #endregion
    }        
}