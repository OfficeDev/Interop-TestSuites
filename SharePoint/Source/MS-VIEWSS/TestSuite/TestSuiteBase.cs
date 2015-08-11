namespace Microsoft.Protocols.TestSuites.MS_VIEWSS
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Text;
    using System.Web.Services.Protocols;
    using System.Xml;
    using System.Xml.Serialization;
    using Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Contain test cases designed to test MS-VIEWSS protocol.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Variables
        /// <summary>
        /// A private static string list contains a pool of the view names of which have been added by the test suite.
        /// </summary>
        private static List<string> viewPool = new List<string>();

        /// <summary>
        /// Original view name of the default view of the list.
        /// </summary>
        private static string originalDefaultViewName;

        /// <summary>
        /// True indicates the original default view exists and has been changed into a normal view; False indicates otherwise.
        /// </summary>
        private static bool originalDefaultViewLost = false;

        /// <summary>
        /// True indicates that it is checked whether there is a default view before any test case begins; False indicates otherwise.
        /// </summary>
        private static bool originalDefaultViewChecked = false;

        /// <summary>
        /// An enumeration that indicates the kind of query condition in the view.
        /// </summary>
        protected enum Query
        {
            /// <summary>
            /// Available Query.
            /// </summary>
            AvailableQueryInfo,

            /// <summary>
            /// Indicate the Query field don't contain any information
            /// </summary>
            EmptyQueryInfo,

            /// <summary>
            /// Indicate the Collapse field is true.
            /// </summary>
            IsCollapse,

            /// <summary>
            /// Indicate the Collapse field is false.
            /// </summary>
            IsNotCollapse,
        }

        /// <summary>
        /// An enumeration that indicates the type of view to add.
        /// </summary>
        protected enum ViewType
        {
            /// <summary>
            /// Represents a calendar list view.
            /// </summary>
            Calendar,

            /// <summary>
            /// Represents a datasheet list view.
            /// </summary>
            Grid,

            /// <summary>
            /// Represents a standard HTML list view.
            /// </summary>
            Html
        }

        /// <summary>
        /// Gets or sets MS_VIEWSS Protocol Adapter instance.
        /// </summary>
        protected static IMS_VIEWSSAdapter Adapter { get; set; }

        /// <summary>
        /// Gets or sets MS_VIEWSS SUT Control Adapter instance.
        /// </summary>
        protected static IMS_VIEWSSSUTControlAdapter SutControlAdapter { get; set; }

        /// <summary>
        /// Gets or sets the GUID of the list.
        /// </summary>
        protected static string ListGUID { get; set; }

        /// <summary>
        /// Gets or sets the view pool.
        /// </summary>
        protected static List<string> ViewPool
        {
            get
            {
                return viewPool;
            }

            set
            {
                viewPool = value;
            }
        }

        /// <summary>
        /// Gets or sets the original default view name.
        /// </summary>
        protected static string OriginalDefaultViewName
        {
            get
            {
                return originalDefaultViewName;
            }

            set
            {
                originalDefaultViewName = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the original default view lost its default view position.
        /// </summary>
        protected static bool OriginalDefaultViewLost
        {
            get
            {
                return originalDefaultViewLost;
            }

            set
            {
                originalDefaultViewLost = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the original default view existence is checked.
        /// </summary>
        protected static bool OriginalDefaultViewChecked
        {
            get
            {
                return originalDefaultViewChecked;
            }

            set
            {
                originalDefaultViewChecked = value;
            }
        }
        #endregion Variables

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the class.
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            // A method is used to initialize the variables.
            TestClassBase.Initialize(testContext);

            Adapter = BaseTestSite.GetAdapter<IMS_VIEWSSAdapter>();
            SutControlAdapter = BaseTestSite.GetAdapter<IMS_VIEWSSSUTControlAdapter>();

            string displayListName = Common.GetConfigurationPropertyValue("DisplayListName", BaseTestSite);

            TestSuiteBase.ListGUID = SutControlAdapter.GetListGuidByName(displayListName);

            if (TestSuiteBase.OriginalDefaultViewChecked == false)
            {
                // Get name of the default view.
                TestSuiteBase.OriginalDefaultViewName = SutControlAdapter.GetListAndView(displayListName);
                TestSuiteBase.OriginalDefaultViewChecked = true;
            }
        }

        /// <summary>
        /// Clean up the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion Test Suite Initialization

        #region Test Case Initialization

        /// <summary>
        /// Initialize the test.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
        }

        /// <summary>
        /// Clean up the test.
        /// </summary>
        protected override void TestCleanup()
        {
            this.DeleteViews();

            TestSuiteBase.RestoreOriginalDefaultView();
            base.TestCleanup();
        }

        #endregion Test Case Initialization

        #region Helper Methods

        /// <summary>
        /// A method to get all the HTML and CAML configure elements from the configure file ViewHtmlConfigure.xml 
        /// for the UpdateViewHtml2 operation.
        /// </summary>
        /// <param name="toolBar">The outer UpdateViewHtml2Toolbar instance.</param>
        /// <param name="viewHeader">The outer UpdateViewHtml2ViewHeader instance.</param>
        /// <param name="viewBody">The outer UpdateViewHtml2ViewBody instance.</param>
        /// <param name="viewFooter">The outer UpdateViewHtml2ViewFooter instance.</param>
        /// <param name="viewEmpty">The outer UpdateViewHtml2ViewEmpty instance.</param>
        /// <param name="rowLimitExceeded">The outer UpdateViewHtml2RowLimitExceeded instance.</param>
        protected void GetHtmlConfigure(
               out UpdateViewHtml2Toolbar toolBar,
               out UpdateViewHtml2ViewHeader viewHeader,
               out UpdateViewHtml2ViewBody viewBody,
               out UpdateViewHtml2ViewFooter viewFooter,
               out UpdateViewHtml2ViewEmpty viewEmpty,
               out UpdateViewHtml2RowLimitExceeded rowLimitExceeded)
        {
            toolBar = new UpdateViewHtml2Toolbar();
            viewHeader = new UpdateViewHtml2ViewHeader();
            viewBody = new UpdateViewHtml2ViewBody();
            viewFooter = new UpdateViewHtml2ViewFooter();
            viewEmpty = new UpdateViewHtml2ViewEmpty();
            rowLimitExceeded = new UpdateViewHtml2RowLimitExceeded();

            bool fileExisted = File.Exists("ViewHtmlConfigure.xml");
            this.Site.Assume.IsTrue(fileExisted, "The file \"ViewHtmlConfigure.xml\" should exist in current path.");

            // If the view HTML configure file exists, get the HTML or CAML content from the configuration.
            if (fileExisted)
            {
                toolBar.Toolbar = new ToolbarDefinition();
                viewHeader.ViewHeader = new UpdateViewHtml2ViewHeaderViewHeader();
                viewBody.ViewBody = new UpdateViewHtml2ViewBodyViewBody();
                viewFooter.ViewFooter = new UpdateViewHtml2ViewFooterViewFooter();
                viewEmpty.ViewEmpty = new UpdateViewHtml2ViewEmptyViewEmpty();
                rowLimitExceeded.RowLimitExceeded = new UpdateViewHtml2RowLimitExceededRowLimitExceeded();

                // Load the ViewHtmlConfigure.xml into XmlDocument.
                XmlDocument doc = new XmlDocument();
                doc.Load("ViewHtmlConfigure.xml");

                // Get all the configured HTML or CAML elements.
                toolBar.Toolbar.Any = this.GetConfiguredElements("Toolbar", doc);
                viewHeader.ViewHeader.Any = this.GetConfiguredElements("ViewHeader", doc);
                viewBody.ViewBody.Any = this.GetConfiguredElements("ViewBody", doc);
                viewFooter.ViewFooter.Any = this.GetConfiguredElements("ViewFooter", doc);
                viewEmpty.ViewEmpty.Any = this.GetConfiguredElements("ViewEmpty", doc);
                rowLimitExceeded.RowLimitExceeded.Any = this.GetConfiguredElements("RowLimitExceeded", doc);
            }
        }

        /// <summary>
        /// A method to get all the HTML and CAML configure elements from the configure file ViewHtmlConfigure.xml 
        /// for the UpdateViewHtml operation.
        /// </summary>
        /// <param name="toolBar">The outer UpdateViewHtmlToolbar instance.</param>
        /// <param name="viewHeader">The outer UpdateViewHtmlViewHeader instance.</param>
        /// <param name="viewBody">The outer UpdateViewHtmlViewBody instance.</param>
        /// <param name="viewFooter">The outer UpdateViewHtmlViewFooter instance.</param>
        /// <param name="viewEmpty">The outer UpdateViewHtmlViewEmpty instance.</param>
        /// <param name="rowLimitExceeded">The outer UpdateViewHtmlRowLimitExceeded instance.</param>
        protected void GetHtmlConfigure(
                                    out UpdateViewHtmlToolbar toolBar,
                                    out UpdateViewHtmlViewHeader viewHeader,
                                    out UpdateViewHtmlViewBody viewBody,
                                    out UpdateViewHtmlViewFooter viewFooter,
                                    out UpdateViewHtmlViewEmpty viewEmpty,
                                    out UpdateViewHtmlRowLimitExceeded rowLimitExceeded)
        {
            toolBar = new UpdateViewHtmlToolbar();
            viewHeader = new UpdateViewHtmlViewHeader();
            viewBody = new UpdateViewHtmlViewBody();
            viewFooter = new UpdateViewHtmlViewFooter();
            viewEmpty = new UpdateViewHtmlViewEmpty();
            rowLimitExceeded = new UpdateViewHtmlRowLimitExceeded();

            bool fileExisted = File.Exists("ViewHtmlConfigure.xml");
            this.Site.Assume.IsTrue(fileExisted, "The file \"ViewHtmlConfigure.xml\" should exist in current path.");

            // If the view HTML configure file exists, get the HTML or CAML content from the configuration.
            if (fileExisted)
            {
                toolBar.Toolbar = new ToolbarDefinition();
                viewHeader.ViewHeader = new UpdateViewHtmlViewHeaderViewHeader();
                viewBody.ViewBody = new UpdateViewHtmlViewBodyViewBody();
                viewFooter.ViewFooter = new UpdateViewHtmlViewFooterViewFooter();
                viewEmpty.ViewEmpty = new UpdateViewHtmlViewEmptyViewEmpty();
                rowLimitExceeded.RowLimitExceeded = new UpdateViewHtmlRowLimitExceededRowLimitExceeded();

                // Load the ViewHtmlConfigure.xml into XmlDocument.
                XmlDocument doc = new XmlDocument();
                doc.Load("ViewHtmlConfigure.xml");

                // Get all the configured HTML or CAML elements.
                toolBar.Toolbar.Any = this.GetConfiguredElements("Toolbar", doc);
                viewHeader.ViewHeader.Any = this.GetConfiguredElements("ViewHeader", doc);
                viewBody.ViewBody.Any = this.GetConfiguredElements("ViewBody", doc);
                viewFooter.ViewFooter.Any = this.GetConfiguredElements("ViewFooter", doc);
                viewEmpty.ViewEmpty.Any = this.GetConfiguredElements("ViewEmpty", doc);
                rowLimitExceeded.RowLimitExceeded.Any = this.GetConfiguredElements("RowLimitExceeded", doc);
            }
        }

        /// <summary>
        /// Used to get view's format definitions.
        /// </summary>
        /// <returns>Return the view's format definitions</returns>
        protected ViewFormatDefinitions GetViewFormatDefinitions()
        {
            ViewFormatDefinitions formatDefinitions = new ViewFormatDefinitions();
            formatDefinitions.FormatDef = new FormatDefDefinition[1];
            formatDefinitions.FormatDef[0] = new FormatDefDefinition();
            formatDefinitions.FormatDef[0].Type = Common.GetConfigurationPropertyValue("FormatsGeneral_Type", this.Site);
            formatDefinitions.FormatDef[0].Value = Common.GetConfigurationPropertyValue("FormatsGeneral_Value", this.Site);

            formatDefinitions.Format = new FormatDefinition[1];
            formatDefinitions.Format[0] = new FormatDefinition();
            formatDefinitions.Format[0].Name = Common.GetConfigurationPropertyValue("FormatsField_Name", this.Site);
            formatDefinitions.Format[0].FormatDef = new FormatDefDefinition[1];
            formatDefinitions.Format[0].FormatDef[0] = new FormatDefDefinition();
            formatDefinitions.Format[0].FormatDef[0].Type = Common.GetConfigurationPropertyValue("FormatsField_Type", this.Site);
            formatDefinitions.Format[0].FormatDef[0].Value = Common.GetConfigurationPropertyValue("FormatsField_Value", this.Site);

            return formatDefinitions;
        }

        /// <summary>
        /// Used to get the definition of Aggregations for a view. 
        /// </summary>
        /// <param name="isValueOn">Specify whether the Aggregation's value is On or Off.</param>
        /// <param name="hasValidAggregatingFieldRef">Specify whether there is a valid FieldRef.</param>
        /// <param name="aggregationType">Only meaningful when the hasValidAggregatingFieldRef is True, specify the Aggregation Type</param>
        /// <returns>Return the view's AggregationsDefinition.</returns>
        protected AggregationsDefinition GetAggregationsDefinition(
                        bool isValueOn,
                        bool hasValidAggregatingFieldRef,
                         string aggregationType)
        {
            AggregationsDefinition aggregationDefinitions = new AggregationsDefinition();

            // Set the Aggregation's value to either On or Off.
            aggregationDefinitions.Value = isValueOn ? "On" : "Off";

            // If AggregatingFieldRef is supported, then generate FieldRefAggregations by configuration. 
            if (hasValidAggregatingFieldRef)
            {
                aggregationDefinitions.FieldRef = new FieldRefDefinitionAggregation[1];
                aggregationDefinitions.FieldRef[0] = new FieldRefDefinitionAggregation();
                aggregationDefinitions.FieldRef[0].Name
                    = Common.GetConfigurationPropertyValue("FieldRefAggregations_Name", this.Site);
                aggregationDefinitions.FieldRef[0].Type
                    = aggregationType;
            }

            return aggregationDefinitions;
        }

        /// <summary>
        /// Used to get the fields of the list referenced by the view.
        /// </summary>
        /// <param name="isExplicit">Represents whether it is explicit.</param>
        /// <returns>The fields referenced by the view definition.</returns>
        protected FieldRefDefinitionView[] GetViewFields(bool isExplicit)
        {
            FieldRefDefinitionView[] fieldRefs = new FieldRefDefinitionView[4];
            fieldRefs[0] = new FieldRefDefinitionView();
            fieldRefs[0].Name = Common.GetConfigurationPropertyValue("ViewFields0", this.Site);
            fieldRefs[1] = new FieldRefDefinitionView();
            fieldRefs[1].Name = Common.GetConfigurationPropertyValue("ViewFields1", this.Site);
            fieldRefs[2] = new FieldRefDefinitionView();
            fieldRefs[2].Name = Common.GetConfigurationPropertyValue("ViewFields2", this.Site);
            fieldRefs[3] = new FieldRefDefinitionView();
            fieldRefs[3].Name = Common.GetConfigurationPropertyValue("ViewFields3", this.Site);

            foreach (FieldRefDefinitionView fieldRef in fieldRefs)
            {
                fieldRef.Explicit = isExplicit.ToString().ToUpper();
            }

            return fieldRefs;
        }

        /// <summary>
        /// Used to get the query root element containing the where condition.
        /// </summary>
        /// <param name="isLookupId">Explicitly specify whether the field referenced in the logical test query condition is a look up field.</param>
        /// <returns>The query root element containing a where condition.</returns>
        protected CamlQueryRoot GetCamlQueryRootForWhere(bool isLookupId)
        {
            CamlQueryRoot camlQuery = new CamlQueryRoot();

            // Construct a LogicalTestDefinition based on PTF Configure file.
            LogicalTestDefinition logicalTest = new LogicalTestDefinition();
            FieldRefDefinitionQueryTest fieldRef = new FieldRefDefinitionQueryTest();
            fieldRef.Name = Common.GetConfigurationPropertyValue("FieldRefWhere_Name", this.Site);
            logicalTest.FieldRef = fieldRef;
            logicalTest.FieldRef.LookupId = isLookupId.ToString().ToUpper();
            ValueDefinition value = new ValueDefinition();
            value.Type = Common.GetConfigurationPropertyValue("FieldRefWhere_Type", this.Site);
            value.Text = new string[] { Common.GetConfigurationPropertyValue("FieldRefWhere_Text", this.Site) };
            logicalTest.Value = value;

            // Use Equal to construct a Where condition.
            camlQuery.Where = new LogicalJoinDefinition();
            camlQuery.Where.ItemsElementName = new ItemsChoiceType1[] { ItemsChoiceType1.Eq };
            camlQuery.Where.Items = new LogicalTestDefinition[] { logicalTest };

            // Construct an OrderBy element.
            camlQuery.OrderBy = new OrderByDefinition();
            camlQuery.OrderBy.FieldRef = new FieldRefDefinitionOrderBy[1];
            camlQuery.OrderBy.FieldRef[0] = new FieldRefDefinitionOrderBy();
            camlQuery.OrderBy.FieldRef[0].Ascending = Common.GetConfigurationPropertyValue("FieldRefOrderBy_Ascending", this.Site);
            camlQuery.OrderBy.FieldRef[0].Name = Common.GetConfigurationPropertyValue("FieldRefOrderBy_Name", this.Site);

            return camlQuery;
        }

        /// <summary>
        /// Used to get the RowLimitDefinition of a view.
        /// </summary>
        /// <returns>A constructed RowLimitDefinition instance.</returns>
        protected RowLimitDefinition GetAvailableRowLimitDefinition()
        {
            RowLimitDefinition rowLimitDefinition = new RowLimitDefinition();
            rowLimitDefinition.Value = int.Parse(Common.GetConfigurationPropertyValue("AvailableRowLimit", this.Site));
            rowLimitDefinition.Paged = Common.GetConfigurationPropertyValue("IsRowPaged", this.Site);
            return rowLimitDefinition;
        }

        /// <summary>
        /// A method used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>A return value represents the random generated string.</returns>
        protected string GenerateRandomString(int size)
        {
            Random random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// Used to get the query with GroupBy and OrderBy condition.
        /// </summary>
        /// <param name="isCollapse">Indicate whether the result set is collapsed.</param>
        /// <returns>A constructed query with GroupBy and OrderBy condition.</returns>
        protected CamlQueryRoot GetCamlQueryRootForGroupBy(bool isCollapse)
        {
            CamlQueryRoot camlQuery = new CamlQueryRoot();
            GroupByDefinition groupBy = new GroupByDefinition();

            FieldRefDefinitionGroupBy fieldRefGroupBy = new FieldRefDefinitionGroupBy();
            fieldRefGroupBy.Ascending = Common.GetConfigurationPropertyValue("FieldRefGroupBy_Ascending", this.Site);
            fieldRefGroupBy.Name = Common.GetConfigurationPropertyValue("FieldRefGroupBy_Name", this.Site);
            FieldRefDefinitionGroupBy[] fieldRefs = { fieldRefGroupBy };

            groupBy.FieldRef = fieldRefs;
            groupBy.Collapse = isCollapse.ToString();
            groupBy.GroupLimit = int.Parse(Common.GetConfigurationPropertyValue("FieldRefGroupBy_RowLimit", this.Site));

            camlQuery.GroupBy = groupBy;

            return camlQuery;
        }

        /// <summary>
        /// Used to get the query of a view.
        /// </summary>
        /// <param name="query">Specify the type of Query.</param>
        /// <param name="isLookupId">Explicitly specify whether the field referenced in the logical test query condition is a look up field.</param>
        /// <returns>A constructed query root element.</returns>
        protected CamlQueryRoot GetCamlQueryRoot(Query query, bool isLookupId)
        {
            CamlQueryRoot camlQueryRoot = null;
            switch (query)
            {
                case Query.AvailableQueryInfo:
                    camlQueryRoot = this.GetCamlQueryRootForWhere(isLookupId);
                    break;

                case Query.EmptyQueryInfo:
                    camlQueryRoot = new CamlQueryRoot();
                    break;

                case Query.IsCollapse:
                    camlQueryRoot = this.GetCamlQueryRootForGroupBy(true);
                    break;

                case Query.IsNotCollapse:
                    camlQueryRoot = this.GetCamlQueryRootForGroupBy(false);
                    break;

                default:
                    Site.Debug.Fail("Not supported Query type {0}", query.ToString());
                    break;
            }

            return camlQueryRoot;
        }

        /// <summary>
        /// Add a view to the list.
        /// </summary>
        /// <param name="isDefault">Represents whether this view is a default view.</param>
        /// <param name="queryType">Represents the type of query condition of the view.</param>
        /// <param name="viewType">Represents the type of the view to be created.</param>
        /// <returns>The GUID of the view, which is also called the view name.</returns>
        protected string AddView(bool isDefault, Query queryType, ViewType viewType)
        {
            string listName = TestSuiteBase.ListGUID;

            string viewName = this.GenerateRandomString(5);

            AddViewViewFields viewFields = new AddViewViewFields();
            viewFields.ViewFields = this.GetViewFields(true);

            AddViewQuery addViewQuery = new AddViewQuery();
            addViewQuery.Query = this.GetCamlQueryRoot(queryType, false);

            AddViewRowLimit rowLimit = new AddViewRowLimit();
            rowLimit.RowLimit = this.GetAvailableRowLimitDefinition();

            string type = string.Empty;
            switch (viewType)
            {
                case ViewType.Calendar:
                    type = "Calendar";
                    break;
                case ViewType.Grid:
                    type = "Grid";
                    break;
                case ViewType.Html:
                    type = "Html";
                    break;
                default:
                    Site.Debug.Fail("Not supported view type {0}", viewType.ToString());
                    break;
            }

            AddViewResponseAddViewResult addViewResponseAddViewResult = TestSuiteBase.Adapter.AddView(
                                                                                          listName,
                                                                                          viewName,
                                                                                          viewFields,
                                                                                          addViewQuery,
                                                                                          rowLimit,
                                                                                          type,
                                                                                          isDefault);

            this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View, "The call of the AddView operation SHOULD be successful.");

            string viewGUID = addViewResponseAddViewResult.View.Name;
            
            if (isDefault)
            {
                this.Site.Assert.IsNotNull(addViewResponseAddViewResult.View.DefaultView, "The response element \"addViewResponseAddViewResult.View.DefaultView\" should not be null.");
                this.Site.Assert.AreEqual("true", addViewResponseAddViewResult.View.DefaultView.ToLower(), "The added view should be a default view.");

                // If the new default view is added successfully, the original default view lost its default view position.
                if (OriginalDefaultViewName != null)
                {
                    if (TestSuiteBase.OriginalDefaultViewLost == false)
                    {
                        TestSuiteBase.OriginalDefaultViewLost = true;
                    }
                }
            }

            TestSuiteBase.ViewPool.Add(viewGUID);
            return addViewResponseAddViewResult.View.Name;
        }

        /// <summary>
        /// Delete the specific view.
        /// </summary>
        /// <param name="viewname">The GUID of the view.</param>
        protected void DeleteView(string viewname)
        {
            if (TestSuiteBase.ViewPool.Contains(viewname))
            {
                try
                {
                    TestSuiteBase.Adapter.DeleteView(TestSuiteBase.ListGUID, viewname);
                    TestSuiteBase.ViewPool.Remove(viewname);
                }
                catch (SoapException soapExc)
                {
                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        @"There is an exception generated when calling [DeleteView] method:\r\n{0}",
                        soapExc.Message);
                    throw;
                }
            }
        }

        /// <summary>
        /// Performs the comparing based on instance's public contents between
        /// two instances, these two instances must be declared as the same class.
        /// </summary>
        /// <param name="viewDefinition1">The first instance to be compared.</param>
        /// <param name="viewDefinition2">The second instance to be compared.</param>
        /// <returns>Return true if they are equal, else return false.</returns>
        protected bool DoCompare(ViewDefinition viewDefinition1, ViewDefinition viewDefinition2)
        {
            string viewString1 = SerializerHelp(viewDefinition1, typeof(ViewDefinition));
            string viewString2 = SerializerHelp(viewDefinition2, typeof(ViewDefinition));
            return viewString1 == viewString2;
        }
        #endregion

        #region Private helper methods

        /// <summary>
        /// Helper function to serialize an object to string.
        /// </summary>
        /// <param name="targetObject">The serialized target object.</param>
        /// <param name="type">The serialized type.</param>
        /// <returns>The serialized result string.</returns>
        private static string SerializerHelp(object targetObject, Type type)
        {
            XmlSerializer serializer = new XmlSerializer(type);
            StringWriter sw = new StringWriter();
            XmlTextWriter writer = new XmlTextWriter(sw);
            serializer.Serialize(writer, targetObject);
            return sw.ToString();
        }

        /// <summary>
        /// Restore original default view's position.
        /// </summary>
        private static void RestoreOriginalDefaultView()
        {
            if (TestSuiteBase.OriginalDefaultViewName != null && TestSuiteBase.OriginalDefaultViewLost)
            {
                // Call UpdateView method to restore original default view's position.
                UpdateViewViewProperties viewProperties = new UpdateViewViewProperties();
                viewProperties.View = new UpdateViewPropertiesDefinition();
                viewProperties.View.DefaultView = "TRUE";

                // Call UpdateView method without optional parameters.               
                UpdateViewResponseUpdateViewResult updateViewResult = Adapter.UpdateView(
                                                                                            TestSuiteBase.ListGUID,
                                                                                            TestSuiteBase.OriginalDefaultViewName,
                                                                                            viewProperties,
                                                                                            null,
                                                                                            null,
                                                                                            null,
                                                                                            null,
                                                                                            null);

                BaseTestSite.Assert.IsNotNull(updateViewResult.View, "The response element \"updateViewResult.View\" should not be null.");
                BaseTestSite.Assert.IsNotNull(updateViewResult.View.DefaultView, "The response element \"updateViewResult.View.DefaultView\" should not be null.");
                BaseTestSite.Assert.AreEqual("true", updateViewResult.View.DefaultView.ToLower(), "The original default view should be updated to default view.");
                TestSuiteBase.OriginalDefaultViewLost = false;
            }
        }

        /// <summary>
        /// Delete the views created in the test case.
        /// </summary>
        private void DeleteViews()
        {
            List<string> ids = new List<string>(TestSuiteBase.ViewPool);
            foreach (string id in ids)
            {
                try
                {
                    TestSuiteBase.Adapter.DeleteView(TestSuiteBase.ListGUID, id);
                    TestSuiteBase.ViewPool.Remove(id);
                }
                catch (SoapException soapException)
                {
                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        @"There is an exception generated when calling [DeleteView] method:\r\n{0}",
                        soapException.Message);

                    throw;
                }
            }
        }

        /// <summary>
        /// Used to get all the XmlElements under the configured elementName.
        /// </summary>
        /// <param name="elementName">Specify the configured elementName.</param>
        /// <param name="xmlDoc">Specify the XmlDocument instance.</param>
        /// <returns>The XmlElements under the configured elementName.</returns>
        private XmlElement[] GetConfiguredElements(string elementName, XmlDocument xmlDoc)
        {
            XmlNodeList nodes = xmlDoc.GetElementsByTagName(elementName);

            if (nodes.Count > 1)
            {
                Site.Debug.Fail(
                    "The each element in the ViewHtmlConfigure.xml MUST occurs no more than once, but the element {0} occurs {1} times",
                    elementName,
                    nodes.Count);
            }

            // if the element does not exist in the file, then just
            // return null, indicate nothing send in the request.
            if (nodes.Count == 0)
            {
                return null;
            }

            XmlNodeList childNodes = nodes[0].ChildNodes;
            List<XmlElement> list = new List<XmlElement>();

            foreach (XmlNode node in childNodes)
            {
                XmlElement element = node as XmlElement;
                if (element != null)
                {
                    list.Add(element);
                }
            }

            return list.ToArray();
        }

        #endregion
    }
}