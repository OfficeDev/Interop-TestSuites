namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System.Web.Services.Protocols;
    using Common;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the UpdateViewHtml and GetViewHtml operations.
    /// </summary>
    [TestClass]
    public class S04_MaintainViewProperties : TestSuiteBase
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
        /// A test case used to test GetViewHtml method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC01_GetViewHtml_InvalidListName()
        {
            // Call AddView method to add a list view for the specified list on the server.       
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Html);

            bool caughtSoapException = false; 

            // Call GetViewHtml method with invalid listName to get the list view added above.
            try
            {
                Adapter.GetViewHtml(this.GenerateRandomString(5), viewName);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If the server returns SOAP fault, then capture this requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    "[In listName] If the value of listName element is not the name or GUID of a list, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml method with invalid listName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC02_UpdateViewHtml_InvalidListName()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Html);

            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            aggregations.Aggregations = this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            bool caughtSoapException = false;

            // Call UpdateViewHtml method with invalid listName.
            try
            {
                Adapter.UpdateViewHtml(
                                        this.GenerateRandomString(10),
                                        viewName,
                                        viewProperties,
                                        toolbar,
                                        viewHeader,
                                        viewBody,
                                        viewFooter,
                                        viewEmpty,
                                        rowLimitExceeded,
                                        queryValue,
                                        updateViewFields,
                                        aggregations,
                                        formats,
                                        rowLimitValue);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true; 

                // If the server returns SOAP fault, then capture this requirement.
                Site.CaptureRequirementIfIsNotNull(
                    soapException,
                    13,
                    "[In listName] If the value of listName element is not the name or GUID of a list, the operation MUST return a SOAP fault message.");
            }

            this.Site.Assert.IsTrue(caughtSoapException, "There should be a SOAP exception in the response."); 
        }

        /// <summary>
        /// A test case used to test GetViewHtml method with invalid viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC03_GetViewHtml_InvalidViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            bool caughtSoapException = false; 

            // Call GetViewHtml method with invalid viewName.
            try
            {
                Adapter.GetViewHtml(listName, this.GenerateRandomString(10));
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
        /// A test case used to test UpdateViewHtml method with invalid viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC04_UpdateViewHtml_InvalidViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            bool caughtSoapException = false; 

            // Call UpdateViewHtml method with invalid viewName.
            try
            {
                Adapter.UpdateViewHtml(
                            listName,
                            this.GenerateRandomString(5),
                            viewProperties,
                            toolbar,
                            viewHeader,
                            viewBody,
                            viewFooter,
                            viewEmpty,
                            rowLimitExceeded,
                            queryValue,
                            updateViewFields,
                            aggregations,
                            formats,
                            rowLimitValue);
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
        /// A test case used to test GetViewHtml method with null viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC05_GetViewHtml_NullViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call GetViewHtml method with a null viewName.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, null);
            this.Site.Assert.IsNotNull(getViewHtml, "The view html got with null viewName parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The response element \"getViewHtml.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewHtml.View.DefaultView, "The response element \"getViewHtml.View.DefaultView\" should not be null.");

            // If server refers to the default list view of the list, then the below requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                "true",
                getViewHtml.View.DefaultView.ToLower(),
                2301,
                @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list.");           
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml method with null viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC06_UpdateViewHtml_NullViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call UpdateViewHtml method with null viewName.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                        updateViewFields,
                                                                                        aggregations,
                                                                                        formats,
                                                                                        rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View.DefaultView, "The response element \"updateViewHtmlResult.View.DefaultView\" should not be null.");

            // If the protocol server refers to the default list view of the list when viewName element is not present, then capture below requirement.
            Site.CaptureRequirementIfAreEqual(
                "true",
                updateViewHtmlResult.View.DefaultView.ToLower(),
                2301,
                @"[In viewName] When viewName element is not present in the message, the protocol server MUST refer to the default list view of the list.");
        }

        /// <summary>
        /// A test case used to test the GetViewHtml method with an empty viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC07_GetViewHtml_EmptyViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call GetViewHtml method with empty viewName.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, string.Empty);
            this.Site.Assert.IsNotNull(getViewHtml, "The view html got with an empty viewName parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The response element \"getViewHtml.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewHtml.View.DefaultView, "The response element \"getViewHtml.View.DefaultView\" should not be null.");

            // If server refers to the default list view of the list, then the below requirement can be captured.
            Site.CaptureRequirementIfAreEqual(
                "true",
                getViewHtml.View.DefaultView.ToLower(),
                2302,
                @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list.");      
        }

        /// <summary>
        /// A test case used to test the UpdateViewHtml method with an empty viewName parameter.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC08_UpdateViewHtml_EmptyViewName()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);
                    
            // Call UpdateViewHtml method to update the Scope of the view with empty viewName.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.Scope = ViewScope.Item;
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                        updateViewFields,
                                                                                        aggregations,
                                                                                        formats,
                                                                                        rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View.DefaultView, "The response element \"updateViewHtmlResult.View.DefaultView\" should not be null.");

            // If the protocol server refers to the default list view of the list when viewName element is empty, then capture below requirement.
            Site.CaptureRequirementIfAreEqual(
                "true",
                updateViewHtmlResult.View.DefaultView.ToLower(),
                2302,
                @"[In viewName] When the value of viewName element is empty, the protocol server MUST refer to the default list view of the list.");
        }

        /// <summary>
        /// A test case used to test GetViewHtml method with an empty viewName parameter when the default list view does not exist.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC09_GetViewHtml_EmptyViewName_NoDefaultView()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string defaultViewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default view.
            this.DeleteView(defaultViewName);

            bool caughtSoapException = false; 

            // Call GetViewHtml method with empty viewName.
            try
            {
                Adapter.GetViewHtml(listName, string.Empty);
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
        /// A test case used to test UpdateViewHtml method successful with all parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC10_UpdateViewHtml_AllParameters()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string listName = TestSuiteBase.ListGUID;          
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Html);
            
            // Call UpdateViewHtml method to update the FPModified.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.FPModified = "true";
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

            // Call GetViewHtml method to get the list view updated above.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, viewName);
            this.Site.Assert.IsNotNull(getViewHtml, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The response element \"getViewHtml.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewHtml.View.FPModified, "The response element \"getViewHtml.View.FPModified\" should not be null.");
            this.Site.Assert.AreEqual("true", getViewHtml.View.FPModified.ToLower(), "The updated FPModified should be got successfully!");

            // The protocol server successfully updates the default list view and return a View element that specifies the list view, so capture this requirement.
                Site.CaptureRequirement(
                    127,
                    @"[In UpdateViewHtmlResponse] UpdateViewHtmlResult: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml method successful without the optional parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC11_UpdateViewHtml_WithoutOptionalParameters()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            this.AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Call UpdateViewHtml method without the optional parameters.
            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
                                                                                        listName,
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
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View.DefaultView, "The response element \"updateViewHtmlResult.View.DefaultView\" should not be null.");
            this.Site.Assert.AreEqual("true", updateViewHtmlResult.View.DefaultView.ToLower(), "The protocol server should refer to the default list view of the list when viewName element is not present!");

            // The protocol server successfully updates the default list view and return a View element that specifies the list view, so capture this requirement.
            Site.CaptureRequirement(
                127,
                @"[In UpdateViewHtmlResponse] UpdateViewHtmlResult: If the protocol server successfully updates the list view, it MUST return a View element that specifies the list view.");
        }

        /// <summary>
        /// A test case used to verify UpdateViewHtml and UpdateViewHtml2 return values are the same when OpenApplicationExtension is not present for UpdateViewHtml2.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC12_CompareUpdateViewHtmlResult()
        {
            // Call AddView method to add a list view for the specified list on the server.           
            string viewName = this.AddView(true, Query.AvailableQueryInfo, ViewType.Html);
            string listName = TestSuiteBase.ListGUID;

            // Call UpdateViewHtml method to update the display name of the view added above.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);
            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtmlViewFields updateViewFields = new UpdateViewHtmlViewFields();
            updateViewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                       updateViewFields,
                                                                                       aggregations,
                                                                                       formats,
                                                                                       rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

            // Call UpdateViewHtml2 method to update the display name of the view when openApplicationExtension is set to null and all other parameters are same as above UpdateViewHtml method.
            UpdateViewHtml2ViewProperties viewProperties2 = new UpdateViewHtml2ViewProperties();
            viewProperties2.View = new UpdateViewPropertiesDefinition();
            viewProperties2.View.DisplayName = viewProperties.View.DisplayName;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtml2Query queryValue2 = new UpdateViewHtml2Query();
            queryValue.Query = this.GetCamlQueryRootForWhere(false);

            UpdateViewHtml2ViewFields viewFields = new UpdateViewHtml2ViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtml2Aggregations aggregations2 = new UpdateViewHtml2Aggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtml2Formats formats2 = new UpdateViewHtml2Formats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtml2RowLimit rowLimitValue2 = new UpdateViewHtml2RowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtml2ResponseUpdateViewHtml2Result updateViewHtml2Result = Adapter.UpdateViewHtml2(
                                                                                       listName,
                                                                                       viewName,
                                                                                       viewProperties2,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       null,
                                                                                       queryValue2,
                                                                                       viewFields,
                                                                                       aggregations2,
                                                                                       formats2,
                                                                                       rowLimitValue2,
                                                                                       null);

            this.Site.Assert.IsNotNull(updateViewHtml2Result, "The updated view html2 should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtml2Result.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

            // Compare the two responses returned by UpdateViewHtml and UpdateViewHtml2.
            bool isSame = this.DoCompare(updateViewHtmlResult.View, updateViewHtml2Result.View);
            
            // If the two responses are equal, then capture this requirement.
            Site.CaptureRequirementIfIsTrue(
                isSame,
                12501,
                @"[In UpdateViewHtml] When processing this call[[UpdateViewHtml]], the protocol server MUST return the same results as for the UpdateViewHtml2 method (section 3.1.4.8) with parameter openApplicationExtension as empty.");
        }

        /// <summary>
        /// A test case used to test GetViewHtml method with valid parameters.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC13_GetViewHtml_Success()
        {           
            // Call AddView method to add a list view for the specified list on the server.             
            string viewName = this.AddView(true, Query.AvailableQueryInfo, ViewType.Html);
            string listName = TestSuiteBase.ListGUID;

            // Call GetViewHtml method to get the list view created above.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, viewName);
            this.Site.Assert.IsNotNull(getViewHtml, "The view html got with valid parameter should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The GetViewHtml response should include the corresponding View element when the operation succeeds");

            // The GetViewHtml response includes the corresponding View element, then the following requirement can be captured.
            Site.CaptureRequirement(
                117,
                @"[In GetViewHtmlResponse] It [GetViewHtmlResult] MUST include the corresponding View element when the operation [GetViewHtml] succeeds.");
        }

        /// <summary>
        /// A test case used to test UpdateViewHtml method without computed fields and collapse is set to true in the query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC14_UpdateViewHtml_TrueCollapse_NoComputedFields()
        {
            this.Site.Assume.IsTrue(Common.IsRequirementEnabled(1507, this.Site), @"The test case is executed only when R1507Enabled is set to true.");

            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.IsCollapse, ViewType.Grid);
            string listName = TestSuiteBase.ListGUID;

            // Call UpdateViewHtml method to update the display name of the list view added above.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForGroupBy(true);

            UpdateViewHtmlViewFields viewFields = new UpdateViewHtmlViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                        rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

            // Call GetViewHtml method to get the list view updated above.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, viewName);
            this.Site.Assert.IsNotNull(getViewHtml, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The response element \"getViewHtml.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewHtml.View.DisplayName, "The response element \"getViewHtml.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(viewProperties.View.DisplayName, getViewHtml.View.DisplayName, "The fields in the list view updated in the step above should be got successfully!");

            // Call SUT control adapter method GetItemsCount to get the count of the list items in the specified view.
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
        /// A test case used to test UpdateViewHtml method when LogicalJoinDefinition is present and not empty in the query.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC15_UpdateViewHtml_LogicalJoinDefinitionPresent()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.EmptyQueryInfo, ViewType.Html);
            string listName = TestSuiteBase.ListGUID;
         
            // Call UpdateViewHtml method to update the query condition when LogicalJoinDefinition is present in the query and its child element is "Eq".
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();

            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForWhere(true);

            UpdateViewHtmlViewFields viewFields = new UpdateViewHtmlViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                       rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

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
        /// A test case used to test UpdateViewHtml method when there are no child elements in LogicalJoinDefinition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC16_UpdateViewHtml_LogicalJoinDefinitionWithoutChild()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.AvailableQueryInfo, ViewType.Html);
            string listName = TestSuiteBase.ListGUID;
          
            // Call UpdateViewHtml method to update the query condition when there are no child elements in LogicalJoinDefinition.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();

            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForGroupBy(false);

            UpdateViewHtmlViewFields viewFields = new UpdateViewHtmlViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                       rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

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
        /// A test case used to test UpdateViewHtml method when Collapse is set to false in the GroupBy query condition.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC17_UpdateViewHtml_FalseCollapse()
        {
            // Call AddView method to add a list view for the specified list on the server.
            string viewName = this.AddView(false, Query.IsNotCollapse, ViewType.Html);
            string listName = TestSuiteBase.ListGUID;

            // Call UpdateViewHtml method to update the display name of the list view added above.
            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForGroupBy(false);

            UpdateViewHtmlViewFields viewFields = new UpdateViewHtmlViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            UpdateViewHtmlResponseUpdateViewHtmlResult updateViewHtmlResult = Adapter.UpdateViewHtml(
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
                                                                                       rowLimitValue);
            this.Site.Assert.IsNotNull(updateViewHtmlResult, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(updateViewHtmlResult.View, "The server should return a View element that specifies the list view when the UpdateViewHtml method succeeds!");

            // Call GetView method to get the list view updated above.
            GetViewHtmlResponseGetViewHtmlResult getViewHtml = Adapter.GetViewHtml(listName, viewName);
            this.Site.Assert.IsNotNull(getViewHtml, "The updated view html should be got successfully.");
            this.Site.Assert.IsNotNull(getViewHtml.View, "The response element \"getViewHtml.View\" should not be null.");
            this.Site.Assert.IsNotNull(getViewHtml.View.DisplayName, "The response element \"getViewHtml.View.DisplayName\" should not be null.");
            this.Site.Assert.AreEqual(getViewHtml.View.DisplayName, viewProperties.View.DisplayName, "The fields in the list view updated in the step above should be got successfully!");

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
        /// A test case used to test UpdateViewHtml method with an empty viewName parameter when the default list view does not exist.
        /// </summary>
        [TestCategory("MSVIEWSS"), TestMethod]
        public void MSVIEWSS_S04_TC18_UpdateViewHtml_EmptyViewName_NoDefaultView()
        {
            string listName = TestSuiteBase.ListGUID;

            // Add a default view.
            string defaultViewName = AddView(true, Query.EmptyQueryInfo, ViewType.Grid);

            // Delete the default view.
            this.DeleteView(defaultViewName);

            UpdateViewHtmlViewProperties viewProperties = new UpdateViewHtmlViewProperties();
            viewProperties.View = new UpdateViewPropertiesDefinition();
            viewProperties.View.DisplayName = this.GenerateRandomString(10);

            UpdateViewHtmlToolbar toolbar;
            UpdateViewHtmlViewHeader viewHeader;
            UpdateViewHtmlViewBody viewBody;
            UpdateViewHtmlViewFooter viewFooter;
            UpdateViewHtmlViewEmpty viewEmpty;
            UpdateViewHtmlRowLimitExceeded rowLimitExceeded;

            this.GetHtmlConfigure(
                                out toolbar,
                                out viewHeader,
                                out viewBody,
                                out viewFooter,
                                out viewEmpty,
                                out rowLimitExceeded);

            UpdateViewHtmlQuery queryValue = new UpdateViewHtmlQuery();
            queryValue.Query = this.GetCamlQueryRootForGroupBy(false);

            UpdateViewHtmlViewFields viewFields = new UpdateViewHtmlViewFields();
            viewFields.ViewFields = this.GetViewFields(false);

            UpdateViewHtmlAggregations aggregations = new UpdateViewHtmlAggregations();
            this.GetAggregationsDefinition(true, true, Common.GetConfigurationPropertyValue("FieldRefAggregations_AggregationsType", this.Site));

            UpdateViewHtmlFormats formats = new UpdateViewHtmlFormats();
            formats.Formats = this.GetViewFormatDefinitions();

            UpdateViewHtmlRowLimit rowLimitValue = new UpdateViewHtmlRowLimit();
            rowLimitValue.RowLimit = this.GetAvailableRowLimitDefinition();

            bool caughtSoapException = false; 

            // Call UpdateViewHtml method to update the display name of the view with an empty viewName.
            try
            {
                Adapter.UpdateViewHtml(
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
                            rowLimitValue);                
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

        #endregion
    }
}