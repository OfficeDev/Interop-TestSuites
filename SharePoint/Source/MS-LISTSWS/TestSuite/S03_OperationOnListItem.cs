namespace Microsoft.Protocols.TestSuites.MS_LISTSWS
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Globalization;
    using System.Linq;
    using System.Web.Services.Protocols;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test the following operations with valid or invalid parameters.
    /// <list type="bullet">
    ///     <item>GetListItems</item>
    ///     <item>GetListItemChanges</item>
    ///     <item>GetListItemChangesSinceToken</item>
    ///     <item>GetListItemChangesWithKnowledge</item>
    ///     <item>UpdateListItems</item>
    ///     <item>UpdateListItemsWithKnowledge</item>
    ///     <item>GetVersionCollection</item>
    /// </list>
    /// </summary>
    [TestClass]
    public class S03_OperationOnListItem : TestClassBase
    {
        #region Private member variables

        /// <summary>
        /// Protocol adapter.
        /// </summary>
        private IMS_LISTSWSAdapter listswsAdapter;
        #endregion

        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Class initialization.
        /// </summary>
        /// <param name="testContext">The context of test case.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Class clean up.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }

        #endregion

        #region Test cases

        #region AddDiscussionBoardItem

        /// <summary>
        /// This test case is used to test when the listName does not correspond to a list in AddDiscussionBoardItem operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC01_AddDiscussionBoardItem_Fail_ListNameNotExists()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2751, this.Site), @"Test is executed only when R2751Enabled is set to true.");

            // Construct an invalid ListName does not exist on the server.
            string notExistslistName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string errorCode = string.Empty;
            bool caughtSoapException = false;

            // Construct the value of the message for AddDiscussionBoardItem. 
            byte[] message = TestSuiteHelper.GetMessageDataForAddDiscussionBoardItem();

            try
            {
                // Call AddDiscussionBoardItem to add new discussion items to a specified Discussion Board, with the ListName does not exist on the server.
                this.listswsAdapter.AddDiscussionBoardItem(notExistslistName, message);
            }
            catch (SoapException exp)
            {
                caughtSoapException = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'AddDiscussionBoardItem' with the ListName does not exist on the server.");

            // If error code is "0x82000006", capture R2751.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2751,
                @"[In AddDiscussionBoardItem operation]Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to test when the list is not a discussion board in AddDiscussionBoardItem operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC02_AddDiscussionBoardItem_Fail_NoDiscussionBoard()
        {
            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();

            // Create a list by using the specified list name.
            string listId = TestSuiteHelper.CreateList(listName);

            // Construct the value of the message for AddDiscussionBoardItem.
            byte[] message = TestSuiteHelper.GetMessageDataForAddDiscussionBoardItem();

            // Identify whether the soap exception is returned.
            bool isSoapFaultReturned = false;
            string errorcode = string.Empty;
            try
            {
                // Call AddDiscussionBoardItem to add new discussion items to a specified Discussion Board, with the list which is not a discussion board.
                this.listswsAdapter.AddDiscussionBoardItem(listId, message);
            }
            catch (SoapException exp)
            {
                isSoapFaultReturned = true;
                errorcode = TestSuiteHelper.GetErrorCode(exp);
            }

            #region CaptureRequirement 319,1573

            Site.Log.Add(
                    LogEntryKind.Debug,
                    "Capture requirements #R319 #R1573: isSoapFaultReturned[{0}]; ErrorCode[{1}]",
                    isSoapFaultReturned,
                    string.IsNullOrEmpty(errorcode) ? "NullOrEmpty" : errorcode);

            // If the soap exception is returned, then R319 and R1573 should be covered.
            Site.CaptureRequirementIfIsTrue(
                    isSoapFaultReturned && string.IsNullOrEmpty(errorcode),
                    319,
                    @"[In AddDiscussionBoardItem operation] If the list is not a discussion board, the "
                      + "protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                      isSoapFaultReturned && string.IsNullOrEmpty(errorcode),
                    1573,
                    @"[In AddDiscussionBoardItem operation] [If the list is not a discussion board, the "
                        + "protocol server MUST return a SOAP fault.]There is no error code returned for "
                        + "this fault.");
            #endregion
        }

        /// <summary>
        /// This test case is used to verify AddDiscussionBoardItem operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in the current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC03_AddDiscussionBoardItem_InvalidGUIDAndNotCorrespond_SP3WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2429, this.Site), @"Test is executed only when R2429Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool isSoapFaultGenerated = false;
            string errorstring = string.Empty;
            try
            {
                // Call AddDiscussionBoardItem method to add new discussion items to a specified Discussion Board.
                this.listswsAdapter.AddDiscussionBoardItem(invalidListName, null);
            }
            catch (SoapException soapException)
            {
                isSoapFaultGenerated = true;
                errorstring = TestSuiteHelper.GetErrorString(soapException);
            }

            // If the value of the error is null and the detailed message of SoapException is equal to the 
            // specific value, then R2429 should be covered.
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultGenerated[{0}],errorstring[{1}] for requirement #R2429",
                    isSoapFaultGenerated,
                    string.IsNullOrEmpty(errorstring) ? "NullOrEmpty" : errorstring);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated
                && "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).".Equals(errorstring, StringComparison.OrdinalIgnoreCase),
                2429,
                @"[In Appendix B: Product Behavior]Implementation does not return a SOAP fault with error code 0x82000006. (<24> Section 3.1.4.2: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
        }

        /// <summary>
        /// This test case is used to test when listName is not a valid GUID but corresponds to the list title of a list on the server in AddDiscussionBoardItem operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC04_AddDiscussionBoardItem_Success_ListNameIsInvalidGuid()
        {
            // Construct a ListName. It is an invalid Guid, but it corresponds to the list title of a list.
            string listName = TestSuiteHelper.GetUniqueListName();

            // Construct a template id.
            int discussionBoardTemplateId = (int)TemplateType.Discussion_Board;

            // Create a list by using specified list name and template Id.
            string listId = TestSuiteHelper.CreateList(listName, discussionBoardTemplateId);

            // Construct the value of the message for AddDiscussionBoardItem.
            byte[] message = TestSuiteHelper.GetMessageDataForAddDiscussionBoardItem();

            // Call AddDiscussionBoardItem to add new discussion items to a specified Discussion Board, 
            // with the ListName is an invalid Guid, but it exists on the server.
            AddDiscussionBoardItemResponseAddDiscussionBoardItemResult result = null;
            result = this.listswsAdapter.AddDiscussionBoardItem(listName, message);

            Site.Assert.IsNotNull(result, "AddDiscussionBoardItem operation is successful.");

            if (Common.IsRequirementEnabled(3101, this.Site))
            {
                // Verify requirement R3101.
                // If there is no assert fail, it means AddDiscussionBoardItem operation is successful. R3101 can be captured. 
                Site.CaptureRequirement(
                    3101,
                    @"Implementation does support this method[AddDiscussionBoardItem]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // Call GetListItems to retrieve details about list items in a list that satisfies specified criteria.           
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // extract a row element to a DataTable, it is used for "z:row" data
            DataTable data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            #region Capture Requirements 316,317

            // If the value of ItemCount from the GetListItems operation is equal to 1, then the 
            // DiscussionBoardItem is added successfully, R315 and R320 should be covered.
            Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    data.Rows.Count,
                    316,
                    @"[In AddDiscussionBoardItem operation] If the specified listName is not a valid "
                    + "GUID, check whether the listName corresponds to the list title of a list on the "
                    + "site and if so, use that list.");

            Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    data.Rows.Count,
                    317,
                    @"[In AddDiscussionBoardItem operation] If the specified listName does not "
                    + "correspond to the ID of a list on the site, check whether the listName "
                    + "corresponds to the list title of a list on the site and if so, use that list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the AddDiscussionBoardItem operation when listName is valid GUID. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC05_AddDiscussionBoardItem_Success_ListNameIsValidGuid()
        {
            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();

            // Construct a template id.
            int discussionBoardTemplateId = (int)TemplateType.Discussion_Board;

            // Create a list by using specified list name and template Id.
            string listId = TestSuiteHelper.CreateList(listName, discussionBoardTemplateId);

            // Construct the value of the message for AddDiscussionBoardItem.
            byte[] message = TestSuiteHelper.GetMessageDataForAddDiscussionBoardItem();

            // Call AddDiscussionBoardItem to add new discussion items to a specified Discussion Board.
            AddDiscussionBoardItemResponseAddDiscussionBoardItemResult result = null;
            result = this.listswsAdapter.AddDiscussionBoardItem(listId, message);

            Site.Assert.IsNotNull(result, "AddDiscussionBoardItem operation is successful.");

            // Call GetListItems to retrieve details about list items in a list that satisfies specified criteria.
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // extract a row element to a DataTable, it is used for "z:row" data
            DataTable data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            #region CaptureRequiremen 315,320

            // If the value of ItemCount from the GetListItems operation is equal to 1, then the 
            // DiscussionBoardItem is added successfully, R315 and R320 should be covered.
            Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    data.Rows.Count,
                    315,
                    @"[In AddDiscussionBoardItem operation] If the specified listName is a valid GUID "
                    + "and corresponds to the identification of a list on the site, use that list.");

            Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    data.Rows.Count,
                    320,
                    @"[In AddDiscussionBoardItem operation] If both parameters [listName and "
                        + "message] are valid, the protocol server MUST add a discussion item to the "
                        + "discussion board.");
            #endregion
        }

        #endregion

        #region GetListItemChangesSinceToken

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when the change token is empty. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC06_GetListItemChangesSinceToken_EmptyChangeToken()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(3032, this.Site), @"Test is executed only when R3032Enabled is set to true.");

            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesSinceToken = null;

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Add two items.
            TestSuiteHelper.AddListItems(listId, 2);

            // Call GetListItemChangesSinceToken with empty row limit when the SUT is not WSS3.0.
            getListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                                    listId,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    string.Empty,
                                                                                    null);
            DataTable data = AdapterHelper.ExtractData(getListItemChangesSinceToken.listitems.data.Any);
            Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    data.Rows.Count,
                    3032,
                    "[In GetListItemChangesSinceToken operation] Implementation does ignore the changeToken parameter, if the value of the changeToken parameter is empty. (SharePoint Foundation 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// The test case is used to verify GetListItemChangesSinceToken Operation when querying the MetaInfo field and Properties attribute equals to true.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC07_GetListItemChangesSinceToken_MetaInfo()
        {
            // Create list.
            string listGuid = TestSuiteHelper.CreateList();

            #region Add a list item and add property on the metaInfo field

            // Add a property on the metaInfo field
            string uniquePropertyName = TestSuiteHelper.GetUniqueMetaInfoFieldPropertyName();
            string propertyValue = Guid.NewGuid().ToString("N");
            TestSuiteHelper.AddListItemWithMetaInfoProperty(listGuid, uniquePropertyName, propertyValue);
            #endregion

            // Get the ViewFields whose Properties is true and reference field is MetaInfo.
            // "MetaInfo" is required Field in [MS-LISTSWS]
            string metaInfoFieldName = "MetaInfo";
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { metaInfoFieldName });

            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemResult = null;
            getListItemResult = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                   listGuid,
                                                                    null,
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R743
            if (getListItemResult == null || getListItemResult.listitems == null
                || getListItemResult.listitems.data == null || getListItemResult.listitems.data.Any == null)
            {
                this.Site.Assert.Fail("Get list item response error.");
            }

            if (getListItemResult.listitems.data.Any.Length == 0)
            {
                this.Site.Assert.Fail("Get list item data error");
            }

            // A return value represents the schema definition which is read from specified xsd file.
            XmlNode[] anyItems = getListItemResult.listitems.data.Any;

            // "ows_MetaInfo_" is required prefix in [MS-LISTSWS]
            string expectedPrefix = "ows_MetaInfo_";
            DataTable zrowData = AdapterHelper.ExtractData(anyItems);

            // Pick up the items which begin with "ows_MetaInfo_"
            var matchColumn = from DataColumn columnItem in zrowData.Columns
                              where 0 == columnItem.ColumnName.IndexOf(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                              select columnItem;

            // Find out whether there is an item that contains the added property name in the column items.
            bool isContainAddedPropertyName = matchColumn.Any(founder => (0 < founder.ColumnName.IndexOf(uniquePropertyName, StringComparison.OrdinalIgnoreCase)));

            Site.CaptureRequirementIfIsTrue(
                isContainAddedPropertyName,
                2338,
                @"[In GetListItemChangesSinceToken]  [In viewFields element] If the Properties attribute of the ViewFields structure is set to TRUE and the MetaInfo field is referenced in FieldRef element, then properties related to the list MUST be returned with the prefix ""ows_MetaInfo_"".");
        }

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when there are more than 100 changes after a valid change token.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC08_GetListItemChangesSinceToken_MoreThan100Changes()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Try to get a valid change token.
            string token = this.GetChangeToken(listId);

            // Insert more than 100 items.
            TestSuiteHelper.AddListItems(listId, 120);

            // Call GetListItemChangesSinceToken with a valid change token.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesSinceToken = null;
            getListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                                    listId,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    token,
                                                                                    null);

            // If the changeToken parameter is specified and the total number of list items that have been inserted or updated is greater than 100, only the first 100 changes are returned.
            // If R6981 is enable, protocol SUT would return first 100 items that changed are returned.
            if (Common.IsRequirementEnabled(6981, this.Site))
            {
                Site.Assert.IsNotNull(getListItemChangesSinceToken, "The number of list items of returned list is 100, so the data table cannot be null");
                int count = int.Parse(getListItemChangesSinceToken.listitems.data.ItemCount);

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R698.
                Site.CaptureRequirementIfAreEqual<int>(
                    100,
                    count,
                    6981,
                    @"[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In changeToken field] "
                   + @"Implementation does only return the first 100 items that changed, if the changeToken parameter is specified "
                   + @"and the total number of list items that have been inserted or updated is greater than 100."
                   + @"(The 2007 Microsoft® Office system and Windows® SharePoint® Services 3.0 follow this behavior.)");
            }

            // Here all the 120 changes cannot be returned at once, so MoreChanges is TRUE.
            // Then can capture MS-LISTSWS_R713 MS-LISTSWS_R987.
            Site.CaptureRequirementIfAreEqual<string>(
                bool.TrueString.ToUpper(CultureInfo.CurrentCulture),
                getListItemChangesSinceToken.listitems.Changes.MoreChanges.ToUpper(CultureInfo.CurrentCulture),
                713,
                @"[In GetListItemChangesSinceToken operation] [In "
                    + "GetListItemChangesSinceToken element] [In "
                    + "GetListItemChangesSinceTokenResult element] [In listitems element] [In "
                    + "Changes element] [MoreChanges attribute] The Changes element in the "
                    + @"response MUST include the attribute MoreChanges set to ""TRUE"" if "
                    + "more changes are known to the protocol server than were returned to the "
                    + "protocol client.");

            if (Common.IsRequirementEnabled(4006, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<string>(
                    "TRUE",
                    getListItemChangesSinceToken.listitems.Changes.MoreChanges,
                    4006,
                    @"[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In changeToken field] "
                    + "Implementation dose not make another call to GetListItemChangesSinceToken by using the new change token "
                    + "from the most recent response to retrieve subsequent change when the Changes element in the response "
                    + @"MUST include the attribute MoreChanges set to ""TRUE"".(The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// The test case is used to verify GetListItemChangesSinceToken Operation when the change since token is not specified.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC09_GetListItemChangesSinceToken_NoChangeToken()
        {
            string listID = TestSuiteHelper.CreateList();

            // Add 20 list items.
            TestSuiteHelper.AddListItems(listID, 20);

            // Call GetListItemChangesSinceToken without change token and row limit.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultWithoutRowLimit = null;
            resultWithoutRowLimit = this.listswsAdapter.GetListItemChangesSinceToken(
                                listID,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null);

            DataTable dataWithoutRowLimit = AdapterHelper.ExtractData(resultWithoutRowLimit.listitems.data.Any);
            bool isAllItemsReturned = dataWithoutRowLimit.Rows.Count == 20;
            this.Site.Assert.AreEqual<bool>(
                                  true,
                                  isAllItemsReturned,
                                  "When the change token is not specified, the operation GetListItemChangesSinceToken will return all the items.");

            // Call GetListItemChangesSinceToken without change token and a specified row limit.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultWithRowLimit = null;
            resultWithRowLimit = this.listswsAdapter.GetListItemChangesSinceToken(
                                listID,
                                null,
                                null,
                                null,
                                "10",
                                null,
                                null,
                                null);

            // If specified the row limit, the return items must be the same as row limit.
            DataTable dataWithRowLimit = AdapterHelper.ExtractData(resultWithRowLimit.listitems.data.Any);
            bool isItemsReturnedSubjectToRowLimit = dataWithRowLimit.Rows.Count == 10;
            this.Site.Assert.AreEqual<bool>(
                        true,
                        isItemsReturnedSubjectToRowLimit,
                        "When the change token is not specified, the operation GetListItemChangesSinceToken will return all items constrained by row limit.");

            // This requirement is not fully captured, because the case which is subject to the view's default row limit is not tested.
            bool isVerifyR693 = isAllItemsReturned && isItemsReturnedSubjectToRowLimit;
            Site.Log.Add(
                      LogEntryKind.Debug,
                      "The actual value: isVerifyR693[{0}] for requirement #R693",
                      isVerifyR693,
                      null == resultWithRowLimit.listitems.data.Any ? "NullRow" : "NotNullRow");

            Site.CaptureRequirementIfIsTrue(
                isVerifyR693,
                693,
                @"[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In changeToken field] If excluded by the protocol client all items in the list MUST be returned, subject to the default row limit of the view or the overriding rowLimit parameter.");
        }

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when the change token is not empty but invalid.  
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC10_GetListItemChangesSinceToken_NotMatchChangeToken()
        {
            // Use random generate string as an invalid token.
            string invalidToken = TestSuiteHelper.GenerateRandomString(10);

            // Create a list and add two items.
            string listId = TestSuiteHelper.CreateList();
            TestSuiteHelper.AddListItems(listId, 2);

            // Call GetListItemChangesSinceToken with invalid change token.
            bool isSoapFault = false;
            string errorCode = null;
            try
            {
                this.listswsAdapter.GetListItemChangesSinceToken(
                                                                 listId,
                                                                 null,
                                                                 null,
                                                                 null,
                                                                 null,
                                                                 null,
                                                                 invalidToken,
                                                                 null);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isSoapFault = true;
            }

            Site.Assert.IsTrue(
                    isSoapFault,
                    "When the list name is a GUID and the change token is invalid token value, call the operation GetListItemChangesSinceToken will return a soap fault");

            // If SOAP fault returned, capture R664.
            Site.CaptureRequirementIfIsTrue(
                    isSoapFault,
                    664,
                    @"[In GetListItemChangesSinceToken operation] If the value of the changeToken parameter does not match 
                    a change token known to the protocol server, the protocol server MUST return a SOAP fault.");

            // If there is no error code returned in the soap fault, capture R1878.
            Site.CaptureRequirementIfAreEqual<string>(
                    null,
                    errorCode,
                    1878,
                    @"[In GetListItemChangesSinceToken operation] [If the value of the changeToken parameter does not "
                    + "match a change token known to the protocol server, the protocol server MUST return a SOAP fault.]"
                    + "There is no error code for this fault.");
        }

        /// <summary>
        /// This method is used to test GetListItemChangesSinceToken whether the row limit is specified or not.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC11_GetListItemChangesSinceToken_RowLimit()
        {
            string listId = TestSuiteHelper.CreateList();

            // Set the expected row limit number.
            int expectedRowLimit = 10;

            // Insert more rows than the expected row limit number, but not more than 40.
            TestSuiteHelper.AddListItems(listId, 20);

            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;

            // Call GetListItemChangesSinceToken with the specified row limit number.
            result = this.listswsAdapter.GetListItemChangesSinceToken(
                        listId,
                        null,
                        null,
                        null,
                        expectedRowLimit.ToString(),
                        null,
                        null,
                        null);

            if (null == result || null == result.listitems || null == result.listitems.data)
            {
                this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain any List item change.");
            }

            if (null == result.listitems.data.Any || 0 == result.listitems.data.Any.Length)
            {
                this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain valid change item.");
            }

            int actualListItemReturned = AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count;
            bool isRowLimitSpecifyMaxNumber = actualListItemReturned <= expectedRowLimit;

            this.Site.CaptureRequirementIfIsTrue(
                            isRowLimitSpecifyMaxNumber,
                            678,
                            "[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In rowLimit field] The protocol server MUST NOT return more inserted or updated list items than this parameter specifies.");

            if (Common.IsRequirementEnabled(2219, this.Site))
            {
                this.Site.CaptureRequirementIfIsTrue(
                            isRowLimitSpecifyMaxNumber,
                            2219,
                            "Implementation does specify rowLimit to indicate the maximum number of rows of data to return in the response. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(677, this.Site))
            {
                // Call GetListItemChangesSinceToken without row limit number.
                result = this.listswsAdapter.GetListItemChangesSinceToken(
                            listId,
                            null,
                            null,
                            null,
                            null,
                            null,
                            null,
                            null);

                if (null == result || null == result.listitems || null == result.listitems.data)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain any List item change.");
                }

                if (null == result.listitems.data.Any || 0 == result.listitems.data.Any.Length)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain valid change item.");
                }

                actualListItemReturned = AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count;
                Site.Assert.AreEqual<int>(
                                20,
                                actualListItemReturned,
                                "If does not specify row limit, the response should contain all the rows.");

                // Call GetListItemChangesSinceToken without row limit number.
                result = this.listswsAdapter.GetListItemChangesSinceToken(
                            listId,
                            null,
                            null,
                            null,
                            "0",
                            null,
                            null,
                            null);

                if (null == result || null == result.listitems || null == result.listitems.data)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain any List item change.");
                }

                if (null == result.listitems.data.Any || 0 == result.listitems.data.Any.Length)
                {
                    this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain valid change item.");
                }

                // If return all the rows in either unspecified row limit or "0" row limit,
                // capture R777.
                actualListItemReturned = AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count;
                this.Site.CaptureRequirementIfAreEqual<int>(
                                20,
                                actualListItemReturned,
                                677,
                                "[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In rowLimit field] Implementation does retrieve all rows when specifying a value of 0,unless the changeToken parameter is specified. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesSinceToken operation when all input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC12_GetListItemChangesSinceToken_Succeed()
        {
            #region AddList

            // Create a List by using specified list name
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);
            #endregion

            // Add one item into the list.
            int addedItemNumber = 1;
            TestSuiteHelper.AddListItems(listId, addedItemNumber);

            // Invoke the GetListItemChangesSinceToken operation with valid parameters.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesresult = null;
            bool isSoapFaultExisted = false;
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { AdapterHelper.FieldAuthorName });
            CamlQueryOptions queryOptions = CreateDefaultCamlQueryOptions();
            try
            {
                getListItemChangesresult = this.listswsAdapter.GetListItemChangesSinceToken(
                            uniqueListName,
                            null,
                            null,
                            viewFields,
                            null,
                            queryOptions,
                            null,
                            null);
            }
            catch (SoapException)
            {
                isSoapFaultExisted = true;
            }

            string addedItemNumberStrValue = addedItemNumber.ToString();
            if (null == getListItemChangesresult || null == getListItemChangesresult.listitems || null == getListItemChangesresult.listitems.data
                || !addedItemNumberStrValue.Equals(getListItemChangesresult.listitems.data.ItemCount))
            {
                this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain any List item change.");
            }

            #region CaptureRequirement 656,657,665,7081
            // Capture R656, R657 and R665 if the GetListItemChangesWithKnowledge succeeds.
            Site.CaptureRequirementIfIsFalse(
                isSoapFaultExisted,
                656,
                @"[In GetListItemChangesSinceToken operation] If the specified listName is not a "
                + "valid GUID, check if the listName corresponds to the List title of a list on the site "
                + "and if so, use that list.");

            Site.CaptureRequirementIfIsFalse(
               isSoapFaultExisted,
                657,
                @"[In GetListItemChangesSinceToken operation] If the specified listName does not "
                + "correspond to the identification of a list on the site, check if the listName "
                + "corresponds to the List title of a list on the site and if so, use that list.");

            if (null == getListItemChangesresult.listitems.data.Any || 0 == getListItemChangesresult.listitems.data.Any.Length)
            {
                this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain valid change item.");
            }

            int actualListItemReturned = getListItemChangesresult.listitems.data.Any.Length;
            Site.CaptureRequirementIfAreEqual(
               addedItemNumber,
               actualListItemReturned,
                665,
                @"[GetListItemChangesSinceToken]Otherwise, all list items specified by the listName "
                + "parameter that match the criteria passed in MUST be returned,[ unless the "
                + "protocol server encounters an unexpected situation and MUST return a SOAP fault.]");

            if (Common.IsRequirementEnabled(7081, this.Site))
            {
                Site.CaptureRequirementIfIsFalse(
                    getListItemChangesresult.listitems.MaxRecommendedEmbeddedFileSizeSpecified,
                    7081,
                    @"Implementation does not return this attribute[MaxRecommendedEmbeddedFileSize attribute]. [In Appendix B: Product Behavior] <59> Section 3.1.4.22.2.2: This attribute is not returned by Windows SharePoint Services 3.0, and SharePoint Foundation 2010.");
            }
            #endregion

            #region CaptureRequirement 1906

            bool captureR1906 = true;
            foreach (XmlElement row in getListItemChangesresult.listitems.data.Any)
            {
                System.Collections.IEnumerator attributeEnumerator = row.Attributes.GetEnumerator();
                while (attributeEnumerator.MoveNext())
                {
                    XmlAttribute curAttribute = (XmlAttribute)attributeEnumerator.Current;
                    string attributeName = curAttribute.Name;
                    string prefix = attributeName.Substring(0, 4);
                    if (prefix != "ows_")
                    {
                        captureR1906 = false;
                        string errorInfo = string.Format("The prefix of attribute {0} is not \"ows_\"", attributeName);
                        Site.Log.Add(LogEntryKind.CheckFailed, errorInfo);
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                   captureR1906,
                   1906,
                  "[GetListItemChangesSinceTokenResponse]The names of the attributes containing the list item data in inner z:row elements are prefixed by \"ows_\".");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesSinceToken operation while the change type is MoveAway.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC13_GetListItemChangesSinceToken_MoveAway()
        {
            // Create the source document library.
            string srcListName = TestSuiteHelper.GetUniqueListName();
            string srcListId = TestSuiteHelper.CreateList(srcListName, (int)TemplateType.Document_Library);

            // Create the destination document library.
            string destListName = TestSuiteHelper.GetUniqueListName();
            string destListId = TestSuiteHelper.CreateList(destListName, (int)TemplateType.Document_Library);

            // Upload a file to the source document library.
            IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();
            string srcFileUrl = sutControlAdapter.UploadFile(srcListName);
            Site.Assert.IsTrue(
                !string.IsNullOrEmpty(srcFileUrl),
                "URL of the file being uploaded to list {0} should be returned.",
                srcListName);

            // Try to get the base change token.
            string srcToken = this.GetChangeToken(srcListId);

            // Move the document to destination document library.
            string destFileUrl = string.Format(
                "{0}://{1}/sites/{2}/{3}/{4}",
                Common.GetConfigurationPropertyValue("TransportType", Site),
                Common.GetConfigurationPropertyValue("SutComputerName", Site),
                Common.GetConfigurationPropertyValue("SiteCollectionName", Site),
                destListName,
                System.IO.Path.GetFileName(srcFileUrl));
            sutControlAdapter.MoveFile(srcFileUrl, destFileUrl);

            // Call GetListItemChangesSinceToken to get the change information in source document library.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult srcGetListItemChangesSinceToken = null;
            srcGetListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                srcListId,
                null,
                null,
                null,
                null,
                null,
                srcToken,
                null);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R1436");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1436
            Site.CaptureRequirementIfAreEqual<Guid>(
                new Guid(destListId),
                new Guid(srcGetListItemChangesSinceToken.listitems.Changes.Id.AfterListId),
                1436,
                @"[AfterListId:] Specifies the identifier of the list into which the list item was moved.");

            // Get the identifier of the file after it's moved
            GetListItemsResponseGetListItemsResult getListItemsForFileItem = null;
            getListItemsForFileItem = this.listswsAdapter.GetListItems(
                destListId,
                null,
                null,
                null,
                null,
                null,
                null);
            if (null == getListItemsForFileItem || null == getListItemsForFileItem.listitems || null == getListItemsForFileItem.listitems.data
                || !getListItemsForFileItem.listitems.data.ItemCount.Equals("1"))
            {
                this.Site.Assert.Fail("Failed to get the information of the moved file '{0}' with GetListItems operation.", destFileUrl);
            }

            DataTable rowDataOfGetListitem = AdapterHelper.ExtractData(getListItemsForFileItem.listitems.data.Any);
            uint documentItemId = uint.Parse(rowDataOfGetListitem.Rows[0]["ows_ID"].ToString());

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R1438");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1438
            Site.CaptureRequirementIfAreEqual<uint>(
                documentItemId,
                srcGetListItemChangesSinceToken.listitems.Changes.Id.AfterItemId,
                1438,
                @"AfterItemId: Specifies the identifier of the item after the move change.");
        }

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when the list name is a valid GUID. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC14_GetListItemChangesSinceToken_ValidListGUID()
        {
            string listId = TestSuiteHelper.CreateList();

            // Add two items in the created list.
            TestSuiteHelper.AddListItems(listId, 2);

            // Call GetListItemChangesSinceToken with valid GUID value.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesSinceToken = null;
            getListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                                    listId,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R6501.
            // If getting response is successful, it means implementation does support this GetListItemChangesSinceToken method. R6501 can be captured.
            if (Common.IsRequirementEnabled(6501, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    getListItemChangesSinceToken,
                    6501,
                    @"Implementation does support this method[GetListItemChangesSinceToken]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R655.
            // If there is no error and the response list ID equals to the request list ID,
            // capture the MS-LISTSWS_R655 requirement.
            Site.CaptureRequirementIfAreEqual<string>(
                listId,
                getListItemChangesSinceToken.listitems.Changes.List.ID,
                655,
                @"[In GetListItemChangesSinceToken operation] If the specified listName is a "
                    + "valid GUID and corresponds to the identification of a list on the site, use "
                    + "that list.");
        }

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when the list name is list title. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC15_GetListItemChangesSinceToken_ValidListTitle()
        {
            // Read configure value.
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, fieldName);

            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Add two items in the created list.
            List<string> ids = TestSuiteHelper.AddListItems(listId, 2);

            // Call getListItemChangesSinceToken using list title and exclude all other parameters.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult getListItemChangesSinceToken = null;
            getListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                                    listName,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null);

            // If there is no error and the response list ID equals to the request list ID,
            // capture the MS-LISTSWS_R656, MS-LISTSWS_R657, and MS-LISTSWS_R665 requirements.
            Site.CaptureRequirementIfAreEqual<string>(
                listId,
                getListItemChangesSinceToken.listitems.Changes.List.ID,
                656,
                @"[In GetListItemChangesSinceToken operation] If the specified listName is "
                    + "not a valid GUID, check if the listName corresponds to the List title of a "
                    + "list on the site and if so, use that list.");

            Site.CaptureRequirementIfAreEqual<string>(
                listId,
                getListItemChangesSinceToken.listitems.Changes.List.ID,
                657,
                @"[In GetListItemChangesSinceToken operation] If the specified listName does "
                + "not correspond to the identification of a list on the site, check if the "
                + "listName corresponds to the List title of a list on the site and if so, "
                + "use that list.");

            Site.CaptureRequirementIfAreEqual<string>(
                listId,
                getListItemChangesSinceToken.listitems.Changes.List.ID,
                665,
                @"[GetListItemChangesSinceToken]Otherwise, all list items specified by the "
                    + "listName parameter that match the criteria passed in MUST be returned,"
                    + "[ unless the protocol server encounters an unexpected situation and MUST "
                    + "return a SOAP fault.]");

            // Try to get the list definition using GetList operation.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(listId);

            bool isSameListDef = true;

            // If the R7151 is disabled, the SUT equals to Windows SharePoint Services 3.0. The GetListItemChangesSinceToken operation does not return the FileFragment element in the Changes element.
            if (Common.IsRequirementEnabled(7151, this.Site))
            {
                isSameListDef = TestSuiteHelper.DeepCompare(listDef.Fields, getListItemChangesSinceToken.listitems.Changes.List.Fields);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.RegionalSettings, getListItemChangesSinceToken.listitems.Changes.List.RegionalSettings);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.ServerSettings, getListItemChangesSinceToken.listitems.Changes.List.ServerSettings);
                Site.CaptureRequirementIfIsTrue(
                            isSameListDef,
                            716,
                            "[In GetListItemChangesSinceToken operation]Implementation does return the inner XML of the Changes element which contains the full List element from the GetList operation, if no change token was passed in the request.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }
            else
            {
                // Remove FileFragment fields from GetList response.
                listDef.Fields.Field = (from field in listDef.Fields.Field
                                        where !"FileFragment".Equals(field.Name, StringComparison.OrdinalIgnoreCase)
                                        select field).ToArray();

                // If both results contain the same Fields, RegionalSettings and ServerSettings,
                // both list elements can be thought as the same; Then, R716 can be captured.
                isSameListDef = TestSuiteHelper.DeepCompare(listDef.Fields, getListItemChangesSinceToken.listitems.Changes.List.Fields);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.RegionalSettings, getListItemChangesSinceToken.listitems.Changes.List.RegionalSettings);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.ServerSettings, getListItemChangesSinceToken.listitems.Changes.List.ServerSettings);
                Site.CaptureRequirementIfIsTrue(
                            isSameListDef,
                            716,
                            "[In GetListItemChangesSinceToken operation]Implementation does return the inner XML of the Changes element which contains the full List element from the GetList operation, if no change token was passed in the request.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Add one list field to make the list schema change.
            TestSuiteHelper.AddFieldsToList(
                                        listId,
                                        new List<string> { TestSuiteHelper.GetUniqueFieldName() },
                                        new List<string> { "Currency" },
                                        new List<string> { null });

            // After the list schema changes to call GetListItemChangesSinceToken using the specified change token,
            // expect returning the new list definition.
            string changeToken = getListItemChangesSinceToken.listitems.Changes.LastChangeToken;
            getListItemChangesSinceToken = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                                    listName,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    null,
                                                                                    changeToken,
                                                                                    null);

            // Try to get the list definition using GetList operation again.
            listDef = TestSuiteHelper.GetListDefinition(listId);

            // If the SUT equals WSS3.0 (R7151 is enabled), ignore FileFragment.
            if (Common.IsRequirementEnabled(7151, this.Site))
            {
                isSameListDef = TestSuiteHelper.DeepCompare(listDef.Fields, getListItemChangesSinceToken.listitems.Changes.List.Fields);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.RegionalSettings, getListItemChangesSinceToken.listitems.Changes.List.RegionalSettings);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.ServerSettings, getListItemChangesSinceToken.listitems.Changes.List.ServerSettings);
                Site.CaptureRequirementIfIsTrue(
                        isSameListDef,
                        717,
                        @"[In GetListItemChangesSinceToken operation]Implementation does return the inner XML of the Changes element which contains the full List element from the GetList operation, if there have been schema changes made to the list after the change token.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }
            else
            {
                // Remove FileFragment fields from GetList response.
                listDef.Fields.Field = (from field in listDef.Fields.Field
                                        where !"FileFragment".Equals(field.Name, StringComparison.OrdinalIgnoreCase)
                                        select field).ToArray();

                // If both results contain the same Fields, RegionalSettings and ServerSettings,
                // both list elements can be thought as the same; Then R717 can be captured.
                isSameListDef = TestSuiteHelper.DeepCompare(listDef.Fields, getListItemChangesSinceToken.listitems.Changes.List.Fields);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.RegionalSettings, getListItemChangesSinceToken.listitems.Changes.List.RegionalSettings);
                isSameListDef = isSameListDef && TestSuiteHelper.DeepCompare(listDef.ServerSettings, getListItemChangesSinceToken.listitems.Changes.List.ServerSettings);
                Site.CaptureRequirementIfIsTrue(
                        isSameListDef,
                        717,
                        @"[In GetListItemChangesSinceToken operation]Implementation does return the inner XML of the Changes element which contains the full List element from the GetList operation, if there have been schema changes made to the list after the change token.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Construct view fields and use the specified view fields to invoke GetListItemChangesSinceToken without view name.
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(
                                                false,
                                                new List<string> { fieldName });
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultViewNameNotSpecified = null;
            resultViewNameNotSpecified = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Use the specified view fields to invoke GetListItemChangesSinceToken with invalid view name.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultViewNameInvalid = null;
            resultViewNameInvalid = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    Guid.NewGuid().ToString(),
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // If both the results without view name or with invalid view name are the same,
            // capture R2280.
            bool isSame = TestSuiteHelper.DeepCompare(resultViewNameInvalid, resultViewNameNotSpecified);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2280,
                @"if the viewFields parameter is specified, the server response "
                 + "will be same when the request uses different  values of viewName parameter.");
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2281,
                 @"If [both viewName and viewFields are] specified, the protocol server response will be same when the request uses different values of viewName parameter.");

            DataTable data = AdapterHelper.ExtractData(resultViewNameInvalid.listitems.data.Any);
            Site.Assert.AreEqual<int>(
                            2,
                            data.Rows.Count,
                            "There are must have two rows in the current list state.");

            // Get the first row value as the query condition.
            GetListItemChangesSinceTokenQuery query = new GetListItemChangesSinceTokenQuery();
            query.Query = TestSuiteHelper.CreateQueryRoot(fieldName, Convert.ToString(data.Rows[0][columnName]));

            // Call GetListItemChangesSinceToken with query.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultContainNotSpecified = null;
            resultContainNotSpecified = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    query,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Construct Contains element.
            CamlContains contains = new CamlContains();
            contains.Contains = new CamlContainsContains();
            contains.Contains.FieldRef = new CamlContainsContainsFieldRef();

            // Must contains a field called ID.
            contains.Contains.FieldRef.Name = "ID";
            contains.Contains.Value = new CamlContainsContainsValue();
            contains.Contains.Value.Type = "Counter";

            // Use the second added item ID.
            string containsValue = ids[1];
            contains.Contains.Value.Text = new string[] { containsValue };

            // Call GetListItemChangesSinceToken with contains.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultContainSpecified = null;
            resultContainSpecified = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    query,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    contains);

            // If both of the result are same, capture R2285.
            isSame = TestSuiteHelper.DeepCompare(resultContainNotSpecified, resultContainSpecified);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2285,
                @"if query parameter is specified, the server response will be same for different values of contains parameter.");

            // Call GetListItemChangesSinceToken without QueryOptions.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultExcludeOptions = null;
            resultExcludeOptions = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Call GetListItemChangesSinceToken with CamlQueryOptions using the default value
            // specified by the document.
            CamlQueryOptions options = CreateDefaultCamlQueryOptions();
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultIncludeDefaultOptions = null;
            resultIncludeDefaultOptions = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    options,
                                                                    null,
                                                                    null);

            // If both results are same, capture R679.
            isSame = TestSuiteHelper.DeepCompare(resultExcludeOptions, resultIncludeDefaultOptions);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                679,
                 @"[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element] [In queryOptions"
                     + "field] If the queryOptions parameter is excluded, the default values that the protocol server MUST use "
                     + "are the following: RecurrenceOrderBy – True, ViewAttributes – Scope=\"RecursiveAll\", DateInUtc –True,"
                     + "IncludePermissions – True, IncludeAttachmentUrls – True, IncludeAttachmentVersion – True, "
                     + "RecurrencePatternXMLVersion – v3, ExpandUserField – True, MeetingInstanceID – -1.");

            // Make Query option's RecurrenceOrderBy and RecurrencePatternXMLVersion different from the default value.
            options.QueryOptions.RecurrenceOrderBy = "False";
            options.QueryOptions.RecurrencePatternXMLVersion = EnumRecurrencePatternXMLVersion.Item;

            // Call GetListItemChangesSinceToken with the specified QueryOptions.
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultIncludeSpecifiedOptions = null;
            resultIncludeSpecifiedOptions = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listId,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    options,
                                                                    null,
                                                                    null);

            // If both of the result are the same, capture R689.
            isSame = TestSuiteHelper.DeepCompare(resultExcludeOptions, resultIncludeSpecifiedOptions);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                689,
                @"[In GetListItemChangesSinceToken operation] [In GetListItemChangesSinceToken element]"
                    + "[In queryOptions field] The default values for RecurrenceOrderBy and RecurrencePatternXMLVersion"
                    + "MUST NOT affect the response if RecurrenceData field is not part of the list item.");

            // If the viewFields parameter is not specified and the viewName parameter contains a valid GUID but refers to a view that does not exist, capture R660            
            try
            {
                this.listswsAdapter.GetListItemChangesSinceToken(
                                                        listId,
                                                        Guid.NewGuid().ToString(),
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

                Site.Assert.Fail("GetListItemChangesSinceToken operation must return a SOAP fault If the viewFields parameter is not specified and the viewName parameter contains a valid GUID but refers to a view that does not exist");
            }
            catch (SoapException)
            {
                // If a SOAP exception is caught, it means a SOAP fault is returned. So capture R660.
                Site.CaptureRequirement(
                    660,
                    @"If the viewFields parameter is not specified and the viewName parameter contains a valid GUID but refers to a view that does not exist, the protocol server MUST return a SOAP fault. ");
            }

            // If both of the result are the same, capture R6611
            if (Common.IsRequirementEnabled(6611, this.Site))
            {
                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultInvaildViewNameFirst = null;
                string viewNameFirst = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

                resultInvaildViewNameFirst = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                        listId,
                                                                        viewNameFirst,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null);

                GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult resultInvaildViewNameSecond = null;
                string viewNameSecond = string.Empty;

                do
                {
                    viewNameSecond = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                }
                while (viewNameFirst == viewNameSecond);

                resultInvaildViewNameSecond = this.listswsAdapter.GetListItemChangesSinceToken(
                                                                        listId,
                                                                        viewNameSecond,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null,
                                                                        null);

                // If both of the result are same, capture R6611.
                isSame = TestSuiteHelper.DeepCompare(resultInvaildViewNameFirst, resultInvaildViewNameSecond);
                Site.CaptureRequirementIfIsTrue(
                    isSame,
                    6611,
                    @"[In GetListItemChangesSinceToken operation] When the viewFields parameter is not specified,  the server returns same for two different  viewName parameters which contains an invalid GUID."
                        + "(SharePoint Foundation 2010 and above follow this behavior.) ");
            }

            // If the result is a SOAP fault without error code and contains the expected error string, capture R6612
            if (Common.IsRequirementEnabled(6612, this.Site))
            {
                string viewName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                string expectedErrorString = "GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";

                try
                {
                    this.listswsAdapter.GetListItemChangesSinceToken(
                                                            listId,
                                                            viewName,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null);

                    Site.Assert.Fail("GetListItemChangesSinceToken operation must return a SOAP fault if the viewFields parameter is not specified and the viewName parameter contains an invalid GUID.");
                }
                catch (SoapException ex)
                {
                    string errorCode = TestSuiteHelper.GetErrorCode(ex);
                    string errorString = TestSuiteHelper.GetErrorString(ex);

                    // If the result is a SOAP fault without error code and contains the expected error string, capture R6612
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement #R6612",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorString[{0}] for requirement #R6612",
                        string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

                    bool isR6612Verified = errorCode == null && expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase);

                    Site.CaptureRequirementIfIsTrue(
                        isR6612Verified,
                        6612,
                        @"[In GetListItemChangesSinceToken operation] Implementation does return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"", "
                        + @"if the viewFields parameter is not specified and the viewName parameter contains an invalid GUID.  (<56> Section 3.1.4.22:  Windows SharePoint Services 3.0 returns the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"") ");
                }
            }
        }

        /// <summary>
        /// This test case is used to verify GetListItemChangesSinceToken operation when the parameter changeToken is empty.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC16_GetListItemChangesSinceToken_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2483, this.Site), @"Test is executed only when R2483Enabled is set to true.");

            // Initialize a valid ListName.
            string listname = TestSuiteHelper.CreateList();
            bool isSoapFaultExist = false;
            try
            {
                // Call GetListItemChangesSinceToken method to return changes made to a specified list after the event
                // expressed by the change token, if specified, or to return all the list items in the list.
                this.listswsAdapter.GetListItemChangesSinceToken(
                                                                    listname,
                                                                    string.Empty,
                                                                    null,
                                                                    null,
                                                                    string.Empty,
                                                                    null,
                                                                    string.Empty,
                                                                    null);
            }
            catch (SoapException soapException)
            {
                isSoapFaultExist = true;

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2483
                bool isVerifyR2483 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2483 = string.IsNullOrEmpty(error);

                // If the value of the error is not be returned, then R2483 should be covered.
                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: isSoapFaultExist[{0}],error[{1}] for requirement #R2483",
                        isSoapFaultExist,
                        string.IsNullOrEmpty(error) ? "NullOrEmpty" : error);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2483,
                    2483,
                    @"[In Appendix B: Product Behavior]  Implementation does not ignore the changeToken parameter. (<57> Section 3.1.4.22: In Windows SharePoint Services 3.0, if the value of the changeToken parameter is empty, the protocol server will return a SOAP fault. There is no error code for this fault.)");
            }

            this.Site.Assert.IsTrue(isSoapFaultExist, "Call GetListItemChangesSinceToken operation with invalid ListName must generate a Soap Fault");
        }

        /// <summary>
        /// This test case is used to verify that only the first 40 rows will be returned when changes of inserting or updating are more than 40  and the changeToken parameter is specified in GetListItemChangesSinceToken operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC17_GetListItemChangesSinceToken_changeTokenParameterSpecified()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(2484, this.Site) || Common.IsRequirementEnabled(2485, this.Site),
                @"Test is executed only when R2484Enabled is set to true or R2485Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);

            // Get the latest change token
            string token = this.GetChangeToken(listID);

            // Add more than 40 items
            int addedItemUnmber = 45;
            List<string> ids = TestSuiteHelper.AddListItems(listName, addedItemUnmber);

            // Call GetListItemChangesSinceToken without row limit but with the latest change token
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;
            result = this.listswsAdapter.GetListItemChangesSinceToken(
                                            listID,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            token,
                                            null);

            if (Common.IsRequirementEnabled(2484, this.Site))
            {
                // If only return the first 40 rows, then capture R2484.
                Site.CaptureRequirementIfAreEqual<int>(
                    40,
                    AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count,
                    2484,
                    @"Implementation does only return the first 40 items that changed. [In Appendix B: Product Behavior] <58> Section 3.1.4.22.2.1: In SharePoint Foundation 2010 and SharePoint Foundation 2013, if the changeToken parameter is specified and the total number of list items that have been inserted is greater than 40, only the first 40 items that changed are returned.");

                int leftInsertListItemCount = addedItemUnmber - 40;

                // Get all the remaining 5 insert items, this action will dry out the changes.
                token = result.listitems.Changes.LastChangeToken;
                result = this.listswsAdapter.GetListItemChangesSinceToken(
                                                listID,
                                                null,
                                                null,
                                                null,
                                                null,
                                                null,
                                                token,
                                                null);

                // Assert the left insert list item count.
                this.Site.Assert.AreEqual(
                            leftInsertListItemCount,
                            AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count,
                            "All left  inserted list items should be clean up. The remaining [{0}] inserted list items",
                            leftInsertListItemCount);
            }

            if (Common.IsRequirementEnabled(2485, this.Site))
            {
                // Make another 45 items update.
                TestSuiteHelper.UpdateListItems(listName, ids, OnErrorEnum.Return);

                // Get the updated items changes without row limit but with the latest change token
                token = result.listitems.Changes.LastChangeToken;
                result = this.listswsAdapter.GetListItemChangesSinceToken(
                                                listID,
                                                null,
                                                null,
                                                null,
                                                null,
                                                null,
                                                token,
                                                null);

                // If only return the first 40 rows, then capture R2485.
                Site.CaptureRequirementIfAreEqual<int>(
                    40,
                    AdapterHelper.ExtractData(result.listitems.data.Any).Rows.Count,
                    2485,
                    @"Implementation does only return the first 40 items that changed. [In Appendix B: Product Behavior] <58> Section 3.1.4.22.2.1: In SharePoint Foundation 2010 and SharePoint Foundation 2013, if the changeToken parameter is specified and the total number of list items that have been updated is greater than 40, only the first 40 items that changed are returned.");
            }
        }

        /// <summary>
        /// This test case is used to verify the GetListItemChangesSinceToken operation when listName parameter does not correspond to a list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC18_GetListItemChangesSinceToken_ListNameNotCorrespond()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(6581, this.Site) || Common.IsRequirementEnabled(6582, this.Site),
                @"Test is executed only when R6581Enabled is set to true or R6582Enabled is set to true.");

            // Create list
            string listId = TestSuiteHelper.CreateList();

            // Add one item into the list.
            TestSuiteHelper.AddListItems(listId, 1);

            // Call GetVersionCollection.
            bool isSoapFaultExisted = false;
            string errorCode = string.Empty;

            try
            {
                string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapter.GetListItemChangesSinceToken(invalidListName, null, null, null, null, null, null, null);

                Site.Assert.Fail("GetListItemChangesSinceToken operation must return a SOAP fault if the specified listName does not correspond to a list from either of these checks.");
            }
            catch (SoapException ex)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(ex);

                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement R6581 and R6582",
                        string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                // If error code "0x82000006" is returned, capture R6582.
                if (Common.IsRequirementEnabled(6582, this.Site))
                {
                    Site.CaptureRequirementIfAreEqual<string>(
                        "0x82000006",
                        errorCode,
                        6582,
                        @"[In GetListItemChangesSinceToken operation] Implementation does return a SOAP fault with error code 0x82000006, if the specified listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");
                }

                // If a SOAP fault is returned with no error code, capture R6581.
                if (Common.IsRequirementEnabled(6581, this.Site))
                {
                    Site.CaptureRequirementIfIsNull(
                        errorCode,
                        6581,
                        @"[In GetListItemChangesSinceToken operation] Implementation does not return error code in the SOAP fault, if the specified listName does not correspond to a list from either of these checks. (<55> Section 3.1.4.22: In Windows SharePoint Services 3.0, this error code is not returned in the SOAP fault.)");
                }
            }

            this.Site.Assert.IsTrue(isSoapFaultExisted, "There is no Soap Fault generated for requirement R6581 and R6582");
        }

        #endregion

        #region GetListItemChangesWithKnowledge

        /// <summary>
        /// This test case is used to validate the GetListItemChangesWithKnowledge operation with rowLimit parameter empty.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC19_GetListItemChangesWithKnowledge_EmptyRowLimit()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2220, this.Site), @"Test is executed only when R2220Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);

            // Insert 10 items
            TestSuiteHelper.AddListItems(listID, 10);

            // Call GetListItemChangesWithKnowledge with empty row limit and without knowledge
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListKnowledgeResultWithEmptyRowLimit = null;
            getListKnowledgeResultWithEmptyRowLimit = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                 listName,
                                 null,
                                 null,
                                 null,
                                 string.Empty,
                                 null,
                                 null,
                                 null,
                                 null);

            DataTable data = AdapterHelper.ExtractData(getListKnowledgeResultWithEmptyRowLimit.listitems.data.Any);

            // Because insert 10 items, so if the result contains 10 total items,
            // then requirement R22201 can be captured.
            // R22201 is case derived from R2220, capture R22201.
            Site.CaptureRequirementIfAreEqual<int>(
                10,
                data.Rows.Count,
                22201,
                @"[In GetListItemChangesWithKnowledge]  Implementation does retrieve all rows when get list item changes with knowledge excluding this parameter[rowLimit] or passing an empty element, unless the knowledge parameter is specified. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesWithKnowledge operation when the list name is not valid GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC20_GetListItemChangesWithKnowledge_InvalidListGuid()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            bool isSoapFault = false;
            string errorCode = null;
            try
            {
                // Call GetListItemChangesWithKnowledge with random GUID value.
                this.listswsAdapter.GetListItemChangesWithKnowledge(
                                            TestSuiteHelper.GetUniqueListName(),
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isSoapFault = true;
            }

            Site.Assert.IsTrue(
                    isSoapFault,
                    "When the list name is not valid GUID and list title, call the operation GetListItemChangesWithKnowledge will return a soap fault");

            // If SUT is SharePointFoundation2013 or SharePointServer2013, the error code equals 0x82000006.
            string sutVersion = Common.GetConfigurationPropertyValue("SutVersion", this.Site);
            if (sutVersion == "SharePointFoundation2013" || sutVersion == "SharePointServer2013")
            {
                Site.Assert.AreEqual<string>(
                    "0x82000006",
                    errorCode,
                    @"{0} return a SOAP fault with error code 0x82000006, if the specified listName does not correspond to a list from either of these checks.",
                    sutVersion);

                Site.Assert.AreEqual<string>(
                    "0x82000006",
                    errorCode,
                    @"{0} return a SOAP fault with error code 0x82000006, this indicates that the list does not exist or might have been deleted by another user.",
                    sutVersion);
            }

            // If the error code equals 0x82000006, capture R1084 and R2354.
            if (Common.IsRequirementEnabled(10842, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<string>(
                        "0x82000006",
                        errorCode,
                        10842,
                        @"Implementation does return a SOAP fault with error code 0x82000006, if the specified listName does not correspond to a list from either of these checks.(SharePoint Foundation 2010 follow this behavior.)");
            }

            Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    errorCode,
                    2354,
                    "[In GetListItemChangesWithKnowledge operation] [If listName does not correspond to a list from either of these checks, the protocol server MUST return a SOAP fault with error code 0x82000006.] This indicates that the list does not exist or might have been deleted by another user");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesWithKnowledge 
        /// operation when at least one of its input parameters is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC21_GetListItemChangesWithKnowledge_ListNameNotCorrespond()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            #region AddList
            string listId = TestSuiteHelper.CreateList();
            #endregion

            // Add one item into the list.
            TestSuiteHelper.AddListItems(listId, 1);

            // Call GetVersionCollection.
            bool isSoapFaultExisted = false;
            try
            {
                string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                            invalidListName,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null);
            }
            catch (SoapException)
            {
                isSoapFaultExisted = true;
            }

            this.Site.Assert.IsTrue(isSoapFaultExisted, "There is no Soap Fault generated for requirement#658,#1877");
        }

        /// <summary>
        /// The test case is used to verify GetListItemChangesWithKnowledge Operation when querying the MetaInfo field and Properties attribute equals to true. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC22_GetListItemChangesWithKnowledge_MetaInfo()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            // Create list.
            string listGuid = TestSuiteHelper.CreateList();

            #region Add a list item and add property on the metaInfo field

            // Add a property on the metaInfo field
            string propertyValue = Guid.NewGuid().ToString("N");

            // Add a required property "vti_contentchangeunit" by calling updatelistItemswithKnowledge
            UpdateListItemsWithKnowledgeUpdates updates = new UpdateListItemsWithKnowledgeUpdates();
            updates.Batch = new UpdateListItemsWithKnowledgeUpdatesBatch();
            updates.Batch.Method = new UpdateListItemsWithKnowledgeUpdatesBatchMethod[1];
            updates.Batch.Method[0] = new UpdateListItemsWithKnowledgeUpdatesBatchMethod();
            updates.Batch.Method[0].ID = 0;
            updates.Batch.Method[0].Cmd = MethodCmdEnum.New;
            updates.Batch.Method[0].Field = new UpdateListItemsWithKnowledgeUpdatesBatchMethodField[1];
            updates.Batch.Method[0].Field[0] = new UpdateListItemsWithKnowledgeUpdatesBatchMethodField();
            updates.Batch.Method[0].Field[0].Name = "MetaInfo";
            updates.Batch.Method[0].Field[0].Property = "vti_contentchangeunit";
            updates.Batch.Method[0].Field[0].Value = propertyValue;

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult updateListItemsResult = null;
            updateListItemsResult = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listGuid,
                                                                updates,
                                                                null,
                                                                null);

            if (null == updateListItemsResult || null == updateListItemsResult.Results || 0 == updateListItemsResult.Results.Length)
            {
                this.Site.Assert.Fail("Could not update the [vti_contentchangeunit] property for the MetaInfo field");
            }

            #endregion

            // Get the ViewFields whose Properties is true and reference field is MetaInfo.
            // "MetaInfo" is required Field in [MS-LISTSWS]
            string metaInfoFieldName = "MetaInfo";
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { metaInfoFieldName });
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListItemResult = null;
            getListItemResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                    listGuid,
                                                                    null,
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R743
            if (getListItemResult == null || getListItemResult.listitems == null
                || getListItemResult.listitems.data == null || getListItemResult.listitems.data.Any == null)
            {
                this.Site.Assert.Fail("Get list item response error.");
            }

            if (getListItemResult.listitems.data.Any.Length == 0)
            {
                this.Site.Assert.Fail("Get list item data error");
            }

            // A return value represents the schema definition which is read from specified xsd file.
            XmlNode[] anyItems = getListItemResult.listitems.data.Any;

            // "ows_MetaInfo_" is required prefix in [MS-LISTSWS]
            string expectedPrefix = "ows_MetaInfo_";
            DataTable zrowData = AdapterHelper.ExtractData(anyItems);

            // Pick up the items which begin with "ows_MetaInfo_"
            var matchColumn = from DataColumn columnItem in zrowData.Columns
                              where 0 == columnItem.ColumnName.IndexOf(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                              select columnItem;

            // Find out whether there is an item that contains the added property name in the column.
            // vti_contentchangeunit is required property in [MS-LISTSWS]
            string expectedAddedPropertyName = "vti_contentchangeunit";
            bool isContaincontentchangeunit = matchColumn.Any(founder => (0 < founder.ColumnName.IndexOf(expectedAddedPropertyName, StringComparison.OrdinalIgnoreCase)));

            // If contains OWS_Metatinfo prefix
            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2339
            Site.CaptureRequirementIfIsTrue(
                isContaincontentchangeunit,
                2339,
                @"[In GetListItemChangesWithKnowledge] [In viewFields element] If the Properties attribute of the ViewFields structure is set to TRUE and the MetaInfo field is referenced in FieldRef element, then properties related to the list MUST be returned with the prefix ""ows_MetaInfo_"".");

            // If contains vti_contentchangeunit MetoInfo property 
            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1144
            Site.CaptureRequirementIfIsTrue(
                isContaincontentchangeunit,
                1144,
                @"[In GetListItemChangesWithKnowledgeResponse]If the protocol client requests the MetaInfo field for the list items, the protocol server MUST return the vti_contentchangeunit Metainfo property if the property was set in the item or document and the list item or document has not been updated since the time the property was set by a client.");
        }

        /// <summary>
        /// The test case is used to verify GetListItemChangesWithKnowledge Operation when the viewFields parameter is not specified but contains a correct formatted GUID and refers to a view that does not exist.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC23_GetListItemChangesWithKnowledge_NoExistViewNameGuid()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            string listId = TestSuiteHelper.CreateList();

            // Insert 1 item.
            TestSuiteHelper.AddListItems(listId, 1);

            bool isSoapFaultExist = false;
            string errorCode = string.Empty;
            try
            {
                // Generate valid format GUID, it does not refer to any view.
                string invalidviewID = Guid.NewGuid().ToString();
                this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                    listId,
                                                    invalidviewID,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            }
            catch (SoapException exp)
            {
                isSoapFaultExist = true;
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }

            Site.Assert.IsTrue(
                            isSoapFaultExist,
                            "When the list id is a correctly formatted GUID and refers to a view that does not exist, the server will return soap fault exception.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1086
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000005",
                errorCode,
                1086,
                @"[In GetListItemChangesWithKnowledge operation] If the viewFields parameter is not specified and the viewName parameter contains a correctly formatted GUID and refers to a view that does not exist, the protocol server MUST return a SOAP fault with error code 0x82000005.");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesWithKnowledge operation whether the row limit is specified or not.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC24_GetListItemChangesWithKnowledge_RowLimit()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Insert 101 items.
            TestSuiteHelper.AddListItems(listId, 101);

            // Call the GetListItemChangesWithKnowledge when the knowledge is not specified.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult result = null;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                        listId,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);
            DataTable data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // If the response is fine and contains only 101 items,
            // it can be concluded that the server is using the list specified by the list name parameter in the request.
            Site.CaptureRequirementIfAreEqual<int>(
                            101,
                            data.Rows.Count,
                            1082,
                            "[In GetListItemChangesWithKnowledge operation]If the specified listName is a valid GUID and corresponds to the identification of a list on the site, use that list.");

            // If the returned items number equals 101, capture requirement 101.
            Site.CaptureRequirementIfAreEqual<int>(
                            101,
                            data.Rows.Count,
                            1089,
                            "[In GetListItemChangesWithKnowledge operation]Otherwise, all list items specified by the listName parameter that match the criteria passed in MUST be returned, [unless the protocol server encounters an unexpected situation and MUST return a SOAP fault.]");

            // Call the GetListItemChangesWithKnowledge with list title.
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                        listName,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);
            data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // If the response is fine and contain only 101 items,
            // it can be concluded that the server is using the list specified by the list name parameter in the request.
            Site.CaptureRequirementIfAreEqual<int>(
                            101,
                            data.Rows.Count,
                            1083,
                            "[In GetListItemChangesWithKnowledge operation]2. If the specified listName is not a valid GUID [or does not correspond to the identification of a list on the site, ]check if the listName corresponds to the List title of a list on the site and if so, use that list.");

            // If the returned items number equals 101, the following requirement can be captured.
            Site.CaptureRequirementIfAreEqual<int>(
                            101,
                            data.Rows.Count,
                            2162,
                            "[If listName]does not correspond to the identification of a list on the site, check if the listName corresponds to the List title of a list on the site and if so, use that list.");

            // Call GetListItemChangesWithKnowledge with specified list id and row limit.
            string rowLimit = "2";
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                        listId,
                                                        null,
                                                        null,
                                                        null,
                                                        rowLimit,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

            // If the returned rows count does not exceed the specified row limit number,
            // capture the requirements R1108 and R2168.
            data = AdapterHelper.ExtractData(result.listitems.data.Any);
            int rowCount = data.Rows.Count;
            bool isMatchRowLimit = rowCount <= Convert.ToInt32(rowLimit);
            Site.CaptureRequirementIfIsTrue(
                            isMatchRowLimit,
                            1108,
                            "[In GetListItemChangesWithKnowledge]The protocol server MUST NOT return more inserted or updated list items than this parameter specifies.");

            Site.CaptureRequirementIfIsTrue(
                            isMatchRowLimit,
                            2168,
                            "[In GetListItemChangesWithKnowledge][If excluded by the protocol client, all items in the list MUST be returned, subject to the default row limit of the view or the] overriding rowLimit parameter.");

            // This Requirement was not full captured, we only verify the case which is subject to row limit,
            // but does not verify the case which is subject to the view default row limit.
            // The reason is that constructing a view costs too much work.
            Site.CaptureRequirementIfIsTrue(
                          isMatchRowLimit,
                          1113,
                          "[In GetListItemChangesWithKnowledge]If excluded by the protocol client, all items in the list MUST be returned, [subject to the default row limit of the view or the overriding rowLimit parameter.]");

            // Verify requirement R1109.
            // Call GetListItems to get current items.
            GetListItemsResponseGetListItemsResult listItems = null;
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            DataTable listItemsData = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // Remove the first list items in this list.
            string colunmIdName = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            List<string> allItemsId = new List<string>();
            string id = Convert.ToString(listItemsData.Rows[0][colunmIdName]);
            allItemsId.Add(id);

            TestSuiteHelper.RemoveListItems(listId, allItemsId, OnErrorEnum.Continue);

            // Call the GetListItemChangesWithKnowledge with list ID
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                        listId,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);
            data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // Verify if the deleted item is returned.
            bool isItemChangesExist = false;
            foreach (DataRow row in data.Rows)
            {
                if (id == Convert.ToString(row[colunmIdName]))
                {
                    isItemChangesExist = true;
                    break;
                }
            }

            // If the deleted item is not returned and the total number of rows is 100, capture R1109.
            Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: isItemChangesExist[{0}] for requirement #1109",
                        isItemChangesExist);

            Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: Rows count [{0}] for requirement #1109",
                        data.Rows.Count);

            bool isR1109Verified = !isItemChangesExist && data.Rows.Count == 100;

            Site.CaptureRequirementIfIsTrue(
                    isR1109Verified,
                    1109,
                    @"[In GetListItemChangesWithKnowledge]This operation does not return list items that have been deleted on the protocol server.");
        }

        /// <summary>
        /// This test case is used to verify GetListItemChangesWithKnowledge operation with specified knowledge.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC25_GetListItemChangesWithKnowledge_SpecifiedKnowledge()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            // Create a list
            string listId = TestSuiteHelper.CreateList();

            // Insert 1 items
            TestSuiteHelper.AddListItems(listId, 1);

            // Get the syncKnowledge.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult result = null;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null);
            syncKnowledge knowledge = result.listitems.Changes.MadeWithKnowledge.syncKnowledge;
            bool isMoreChangeIncludedWithoutKnowledge = string.Compare("True", result.listitems.Changes.MoreChanges, StringComparison.OrdinalIgnoreCase) == 0;

            // Insert 20 items again
            int totalNumber = 20;
            TestSuiteHelper.AddListItems(listId, totalNumber);

            // Call GetListItemChangesWithKnowledge to get all the 20 items which are make after the knowledge is retrieved.
            GetListItemChangesWithKnowledgeKnowledge syncKnowledge = new GetListItemChangesWithKnowledgeKnowledge();
            syncKnowledge.syncKnowledge = knowledge;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            syncKnowledge,
                                            null);

            // If the RowLimit is not specified, the server will response all changes made to list items after the knowledge data structure.
            DataTable data = AdapterHelper.ExtractData(result.listitems.data.Any);
            this.Site.Assert.AreEqual<int>(
                            totalNumber,
                            data.Rows.Count,
                            "The response will include changes made to list items after the knowledge data structure was retrieved if there are no RowLimit specified.");

            // Insert another 20 items again.
            TestSuiteHelper.AddListItems(listId, totalNumber);

            // Call GetListItemChangesWithKnowledge with previous knowledge and specified RowLimit.
            syncKnowledge.syncKnowledge = result.listitems.Changes.MadeWithKnowledge.syncKnowledge;
            int expectRowLimit = 5;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            expectRowLimit.ToString(),
                                            null,
                                            null,
                                            syncKnowledge,
                                            null);
            data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // If the RowLimit is specified, the return items will restricted by this element.
            this.Site.Assert.AreEqual<int>(
                            expectRowLimit,
                            data.Rows.Count,
                            "The 5 items will be returned if the RowLimit value is 5 when calling GetListItemChangesWithKnowledge.");

            // If whether the RowLimit is specified or not, the server only return the changes made after knowledge
            // then R1114 can be captured.
            this.Site.CaptureRequirementIfAreEqual<int>(
                        expectRowLimit,
                        data.Rows.Count,
                        1114,
                        @"[In GetListItemChangesWithKnowledge]If specified, the response will include changes made to list items after the knowledge data structure was retrieved, subject to the row limit restrictions just described.");

            bool isMoreChange = string.Compare("True", result.listitems.Changes.MoreChanges, StringComparison.OrdinalIgnoreCase) == 0;
            this.Site.Assert.IsTrue(
                            isMoreChange,
                            "If there are more changes, the server will response attribute MoreChanges 'True', but actual value is {0}",
                            result.listitems.Changes.MoreChanges);

            // If the attribute MoreChange is "True", then R1120 and R1131 can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                        isMoreChange,
                        1120,
                        @"[In GetListItemChangesWithKnowledge]The Changes element in the response MUST include the attribute MoreChanges set to ""TRUE"" to indicate that protocol clients can make another call to GetListItemChangesWithKnowledge by using the new knowledge from the most recent response to retrieve subsequent changes.");
            this.Site.CaptureRequirementIfIsTrue(
                        isMoreChange,
                        1131,
                        @"[In GetListItemChangesWithKnowledgeResponse][Attribute of MoreChanges]The Changes element in the response MUST include the attribute MoreChanges set to ""TRUE"" if more changes are known to the protocol server than were returned to the protocol client.");
            this.Site.CaptureRequirementIfIsTrue(
                        !isMoreChangeIncludedWithoutKnowledge && isMoreChange,
                        1133,
                        @"[In GetListItemChangesWithKnowledgeResponse][Attribute of Morechanges]MoreChanges MUST be included in the result only if the knowledge parameter was specified by the protocol client in the request.");

            // Add one field to make schema changed. 
            TestSuiteHelper.AddFieldsToList(
                                listId,
                                new List<string> { TestSuiteHelper.GetUniqueFieldName() },
                                new List<string> { "Text" },
                                new List<string> { null });

            // Retrieve the knowledge again.
            syncKnowledge.syncKnowledge = result.listitems.Changes.MadeWithKnowledge.syncKnowledge;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                            listId,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            null,
                                            syncKnowledge,
                                            null);

            bool isSchemaChange = string.Compare("True", result.listitems.Changes.SchemaChanged, StringComparison.OrdinalIgnoreCase) == 0;
            this.Site.Assert.IsTrue(
                            isSchemaChange,
                            "If there are schema changed, the server will response attribute SchemaChanged 'True', but actual value is {0}",
                            result.listitems.Changes.SchemaChanged);

            this.Site.CaptureRequirementIfIsTrue(
                            isSchemaChange,
                            1134,
                            @"[In GetListItemChangesWithKnowledgeResponse][Attribute of SchemaChanged]The SchemaChanged attribute MUST be set to ""TRUE"" if there have schema changes made to the list.");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesWithKnowledge 
        /// operation when all input parameters are valid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC26_GetListItemChangesWithKnowledge_Succeed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            // Create a List by using specified list name.
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // Add one item into the list.
            int addedItemNumber = 1;
            TestSuiteHelper.AddListItems(listId, addedItemNumber);

            // Invoke the GetListItemChangesSinceToken operation with valid parameters.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListItemChangesresult = null;
            bool isSoapFaultExisted = false;
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { AdapterHelper.FieldAuthorName });
            CamlQueryOptions queryOptions = CreateDefaultCamlQueryOptions();
            try
            {
                getListItemChangesresult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                        uniqueListName,
                                                                        null,
                                                                        null,
                                                                        viewFields,
                                                                        null,
                                                                        queryOptions,
                                                                        null,
                                                                        null,
                                                                        null);
            }
            catch (SoapException)
            {
                isSoapFaultExisted = true;
            }

            string addedItemNumberStrValue = addedItemNumber.ToString();
            if (null == getListItemChangesresult || null == getListItemChangesresult.listitems || null == getListItemChangesresult.listitems.data
                || !addedItemNumberStrValue.Equals(getListItemChangesresult.listitems.data.ItemCount))
            {
                this.Site.Assert.Fail("The response of GetListItemChangesSinceToken operation does not contain any List item change.");
            }

            // If the operation succeeds, then R24881 can be captured.
            this.Site.CaptureRequirementIfIsFalse(
                isSoapFaultExisted,
                24881,
                @"Implementation does support this method [GetListItemChangesWithKnowledge]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");

            #region CaptureRequirement 1083,2162,1412,11461
            // Capture R1083 and R2162 if the GetListItemChangesWithKnowledge succeeds.
            this.Site.CaptureRequirementIfIsFalse(
                isSoapFaultExisted,
                1083,
                @"[In GetListItemChangesWithKnowledge operation]2. If the specified listName is "
                + "not a valid GUID [or does not correspond to the identification of a list on the site, ]"
                + "check if the listName corresponds to the List title of a list on the site and if so, "
                + "use that list.");

            Site.CaptureRequirementIfIsFalse(
                isSoapFaultExisted,
                2162,
                @"[If listName]does not correspond to the identification of a list on the site, check if "
                + "the listName corresponds to the List title of a list on the site and if so, use that list.");

            DataTable itemRowData = AdapterHelper.ExtractData(getListItemChangesresult.listitems.data.Any);
            bool isAllAttributesBeginWithOws = false;
            bool isAllAttributesCorrespond = false;

            GetListResponseGetListResult getListResult = this.listswsAdapter.GetList(uniqueListName);
            if (null == getListResult || null == getListResult.List || null == getListResult.List.Fields
                || null == getListResult.List.Fields.Field)
            {
                this.Site.Assert.IsNotNull(getListResult, "Could not get the valid fields' data from response of GetList operation");
            }

            ValidateOwsPrefixAndAttributeName(
                                        itemRowData,
                                        getListResult,
                                        out isAllAttributesBeginWithOws,
                                        out isAllAttributesCorrespond);

            // If all "attribute" name begin with "OWS_", capture 1412
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isAllAttributesBeginWithOws[{0}] for requirement #R1412",
                    isAllAttributesBeginWithOws,
                    null == getListResult.List.Fields.Field ? "NullOrEmpty" : "Not NullOrEmpty");

            Site.CaptureRequirementIfIsTrue(
                isAllAttributesBeginWithOws,
                1141,
                @"[In GetListItemChangesWithKnowledgeResponse]The names of the attributes containing the list item data in inner z:row elements are prefixed by ""ows_"".");

            if (Common.IsRequirementEnabled(11461, this.Site))
            {
                Site.CaptureRequirementIfIsFalse(
                    getListItemChangesresult.listitems.MaxRecommendedEmbeddedFileSizeSpecified,
                    11461,
                    @"Implementation does not return this attribute[MaxRecommendedEmbeddedFileSize attribute]. [In Appendix B: Product Behavior] <63> Section 3.1.4.23.2.2: This attribute is not returned by wss3 and wss4.");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to verify the names of the attributes containing the list item data in inner z:row elements prefixed by "ows_" in the response of GetListItemChangesWithKnowledgeResponse operation.  
        /// </summary> 
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC27_GetListItemChangesWithKnowledge_VerifyPrefixOfZrow()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            #region Create a new list.
            string listName = TestSuiteHelper.CreateList();
            #endregion

            #region Add a new list item into the new list.
            TestSuiteHelper.AddListItems(listName, 1);
            #endregion

            #region Invoke "GetListItemChangesWithKnowledge" to get last item changes.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListItemChangesWithKnowledgeResult = null;
            getListItemChangesWithKnowledgeResult
                = this.listswsAdapter.GetListItemChangesWithKnowledge(listName, null, null, null, null, null, null, null, null);
            Site.Assert.IsNotNull(getListItemChangesWithKnowledgeResult, "The response of \"GetListItemChangesWithKnowledge\" is null!");
            #endregion

            #region Confirm the Requirement 1141
            bool captureR1141 = true;
            Site.Assert.IsNotNull(getListItemChangesWithKnowledgeResult.listitems, "The \"listitems\" is null in the response of \"GetListItemChangesWithKnowledge\"");
            Site.Assert.IsNotNull(getListItemChangesWithKnowledgeResult.listitems.data, "The \"listitems.data\" is null in the response of \"GetListItemChangesWithKnowledge\"");
            Site.Assert.IsNotNull(getListItemChangesWithKnowledgeResult.listitems.data.Any, "The \"listitems.data.Any\" is null in the response of \"GetListItemChangesWithKnowledge\"");
            foreach (XmlElement row in getListItemChangesWithKnowledgeResult.listitems.data.Any)
            {
                System.Collections.IEnumerator attributeEnumerator = row.Attributes.GetEnumerator();
                while (attributeEnumerator.MoveNext())
                {
                    XmlAttribute curAttribute = (XmlAttribute)attributeEnumerator.Current;
                    string attributeName = curAttribute.Name;
                    string prefix = attributeName.Substring(0, 4);
                    if (prefix != "ows_")
                    {
                        captureR1141 = false;
                        string errorInfo = string.Format("The prefix of attribute {0} is not \"ows_\"", attributeName);
                        Site.Log.Add(LogEntryKind.TestFailed, errorInfo);
                        break;
                    }
                }
            }

            Site.CaptureRequirementIfIsTrue(
                    captureR1141,
                    1141,
                    "[In GetListItemChangesWithKnowledgeResponse]The names of the attributes containing the list item data in inner z:row elements are prefixed by \"ows_\".");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test that the response does not contain ServerTime attribute when there is not item changes in GetListItemChangesWithKnowledge.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC28_GetListItemChangesWithKnowledge_WithoutItemChanges()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Call GetListItemChangesWithKnowledge without knowledge.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult result = null;
            result = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                        listId,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null,
                                                        null);

            // Now there are no changes in the list, if the returned ServerTime equals null,
            // R1136 can be captured.
            Site.CaptureRequirementIfIsNull(
                result.listitems.Changes.ServerTime,
                1136,
                @"[In GetListItemChangesWithKnowledgeResponse][Attribute of ServerTime]This "
                    + "attribute MUST NOT be set if there no changes are returned.");
        }

        /// <summary>
        ///  This test case is used to validate GetListItemChangesWithKnowledge operation when RowLimit parameter is set as 0.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC29_GetListItemChangesWithKnowledge_ZeroRowLimit()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2220, this.Site), @"Test is executed only when R2220Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);

            // Insert 10 items
            TestSuiteHelper.AddListItems(listID, 10);

            // Call GetListItemChangesWithKnowledge with empty row limit and without knowledge
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListKnowledgeResultWithZeroRowLimit = null;
            getListKnowledgeResultWithZeroRowLimit = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                 listName,
                                 null,
                                 null,
                                 null,
                                 "0",
                                 null,
                                 null,
                                 null,
                                 null);

            DataTable data = AdapterHelper.ExtractData(getListKnowledgeResultWithZeroRowLimit.listitems.data.Any);

            // Because insert 10 items, so if the result contains 10 total items,
            // then requirement R22202 will be able to be captured.
            // R22202 is case derived from R2220, capture R22202.
            Site.CaptureRequirementIfAreEqual<int>(
                    10,
                    data.Rows.Count,
                    22202,
                    @"[In GetListItemChangesWithKnowledge] Implementation does retrieve all rows when get list item changes with knowledge excluding this parameter[rowLimit] or specifying a value of 0, unless the knowledge parameter is specified. (The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChangesWithKnowledge operation with viewFields parameter.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC30_GetListItemChangesWithKnowledge_ViewFields()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(24881, this.Site), @"Test is executed only when R24881Enabled is set to true.");

            // Create a List by using specified list name.
            string uniqueListName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(uniqueListName);

            // Add one item into the list.
            int addedItemNumber = 1;
            TestSuiteHelper.AddListItems(listId, addedItemNumber);

            // Get default view GUID.
            GetListAndViewResponseGetListAndViewResult getListAndView = null;
            getListAndView = this.listswsAdapter.GetListAndView(uniqueListName, string.Empty);
            Site.Assert.IsNotNull(getListAndView, "The object \"getListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView, "The object \"getListAndView.ListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View, "The object \"getListAndView.ListAndView.View\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View.Name, "The object \"getListAndView.ListAndView.View.Name\" should not be null !");
            string strViewGUID = getListAndView.ListAndView.View.Name;

            // Invoke the GetListItemChangesSinceToken operation with null viewName parameter.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult nullViewNameAndViewFieldsResult = null;
            nullViewNameAndViewFieldsResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                    uniqueListName,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Invoke the GetListItemChangesSinceToken operation with invalid viewName parameter.
            string invalidGUID = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult invalidViewNameAndNullViewFieldsResult = null;
            invalidViewNameAndNullViewFieldsResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                    uniqueListName,
                                                                    invalidGUID,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // If the server returns the same responses, capture 2276.
            bool isSame = TestSuiteHelper.DeepCompare(nullViewNameAndViewFieldsResult, invalidViewNameAndNullViewFieldsResult);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2276,
                @"If the viewFields parameter is not specified and the viewName parameter does not contain a correctly formatted GUID, the server response will be same if we use invalid and empty viewName parameter.");

            // Create viewFields 
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { fieldName });

            // Invoke the GetListItemChangesSinceToken operation with valid viewName and viewFields parameters.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult validViewNameAndViewFieldsResult = null;
            validViewNameAndViewFieldsResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                    uniqueListName,
                                                                    strViewGUID,
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // Invoke the GetListItemChangesSinceToken operation with null viewName and valid viewFields parameters.
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult nullViewNameAndValidViewFieldsResult = null;
            nullViewNameAndValidViewFieldsResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                    uniqueListName,
                                                                    null,
                                                                    null,
                                                                    viewFields,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null,
                                                                    null);

            // If the server returns the same responses, capture 2271.
            isSame = TestSuiteHelper.DeepCompare(validViewNameAndViewFieldsResult, nullViewNameAndValidViewFieldsResult);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2271,
                @"if the viewFields parameter is specified, the server response will be same if we use valid and empty viewName parameter.");
        }

        #endregion

        #region GetListItemChanges

        /// <summary>
        /// This test case is used to test that the server returns an error when the listName does not correspond to a list in GetListItemChanges operation. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC31_GetListItemChanges_Fail_ListNameNotExists()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2763, this.Site), @"Test is executed only when R2763Enabled is set to true.");

            // Construct a ListName does not correspond to a list.
            string listName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            string errorCode = string.Empty;
            bool caughtSoapException = false;
            try
            {
                // Call GetListItemChanges operation to verify the List has been changed on the server.
                this.listswsAdapter.GetListItemChanges(
                                                       listName,
                                                       null,
                                                       null,
                                                       null);
            }
            catch (SoapException soapException)
            {
                caughtSoapException = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapException);
            }

            Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'GetListItemChanges' that listName does not correspond to a list.");

            #region Capture Requirements R2763

            // If error code "0x82000006" is returned, capture R2763.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                2763,
                @"[In GetListItemChanges operation] Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItemChanges operation when the date time is an empty string.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC32_GetListItemChanges_Fail_SinceIsEmpty()
        {
            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);
            string since = string.Empty;
            string errorCode = string.Empty;

            // Identify whether the SoapException is returned.
            bool isSoapRaultReturned = false;

            try
            {
                // Call GetListItemChanges operation to verify the List has been changed on the server.
                this.listswsAdapter.GetListItemChanges(
                                                        listId,
                                                        null,
                                                        since,
                                                        null);
            }
            catch (SoapException soapException)
            {
                isSoapRaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapException);

                #region Capture Requirements R3017

                // If there throws soap exception, and with no error code, then R3017 should be covered.
                Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: errorCode[{0}] for requirement #R3017",
                isSoapRaultReturned,
                string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    string.IsNullOrEmpty(errorCode),
                    3017,
                    @"[In GetListItemChanges operation] If the specified since date time is an empty string or an invalid date-time string the protocol server MUST return a SOAP fault. 
                There is no error code for this fault.");

                #endregion
            }

            Site.Assert.IsTrue(isSoapRaultReturned, "GetListItemChanges operation throws SoapException when the Since element is empty string.");
        }

        /// <summary>
        /// This test case is used to test the server behavior when the Name attribute in the input FieldRef element matches no field in AddWikiPage operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC33_GetListItemChanges_FieldRefNameNotMatch()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Insert 10 items.
            TestSuiteHelper.AddListItems(listId, 10);

            // Construct the first invalid view field element.
            CamlViewFields fields1 = TestSuiteHelper.CreateViewFields(
                                                    false,
                                                    new List<string> { TestSuiteHelper.GetUniqueFieldName() });

            // Call the GetListItemChanges with the first invalid field name.
            GetListItemChangesResponseGetListItemChangesResult getListItemChangesResult1 = null;
            getListItemChangesResult1 = this.listswsAdapter.GetListItemChanges(
                                                    listId,
                                                    fields1,
                                                    null,
                                                    null);

            // Construct the second invalid view field element.
            CamlViewFields fields2 = TestSuiteHelper.CreateViewFields(
                                                    false,
                                                    new List<string> { TestSuiteHelper.GetUniqueFieldName() });

            // Call the GetListItemChanges with the second invalid field name.
            GetListItemChangesResponseGetListItemChangesResult getListItemChangesResult2 = null;
            getListItemChangesResult2 = this.listswsAdapter.GetListItemChanges(
                                                    listId,
                                                    fields2,
                                                    null,
                                                    null);

            // We should ignore the list item's timestamp value.
            getListItemChangesResult1.listitems.TimeStamp = null;
            getListItemChangesResult2.listitems.TimeStamp = null;
            bool isSame = TestSuiteHelper.DeepCompare(getListItemChangesResult1, getListItemChangesResult2);

            // If the two result are same, then capture R2282.
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2282,
                @"For different FieldRef.Name values that do not match any fields, the server "
                + "response will be same.");
        }

        /// <summary>
        /// The test case is used to verify GetListItemChanges Operation when querying the MetaInfo field and Properties attribute equals to true.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC34_GetListItemChanges_MetaInfo()
        {
            // Create list.
            string listGuid = TestSuiteHelper.CreateList();

            #region Add a list item and add property on the metaInfo field

            // Add a property on the metaInfo field
            string uniquePropertyName = TestSuiteHelper.GetUniqueMetaInfoFieldPropertyName();
            string propertyValue = Guid.NewGuid().ToString("N");
            TestSuiteHelper.AddListItemWithMetaInfoProperty(listGuid, uniquePropertyName, propertyValue);
            #endregion

            // Get the ViewFields whose Properties is true and reference field is MetaInfo.
            // "MetaInfo" is required Field in [MS-LISTSWS]
            string metaInfoFieldName = "MetaInfo";
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { metaInfoFieldName });
            GetListItemsResponseGetListItemsResult getListItemResult = null;
            getListItemResult = this.listswsAdapter.GetListItems(
                            listGuid,
                            null,
                            null,
                            viewFields,
                            null,
                            null,
                            null);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R743
            if (getListItemResult == null || getListItemResult.listitems == null
                || getListItemResult.listitems.data == null || getListItemResult.listitems.data.Any == null)
            {
                this.Site.Assert.Fail("Get list item response error.");
            }

            if (getListItemResult.listitems.data.Any.Length == 0)
            {
                this.Site.Assert.Fail("Get list item data error");
            }

            // A return value represents the schema definition which is read from specified xsd file.
            XmlNode[] anyItems = getListItemResult.listitems.data.Any;

            // "ows_MetaInfo_" is required prefix in [MS-LISTSWS]
            string expectedPrefix = "ows_MetaInfo_";
            DataTable zrowData = AdapterHelper.ExtractData(anyItems);

            // Pick up the items which begin with "ows_MetaInfo_"
            var matchColumn = from DataColumn columnItem in zrowData.Columns
                              where 0 == columnItem.ColumnName.IndexOf(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                              select columnItem;

            // Find out whether there is column item which column name contains added property name in the columns.
            bool isContainAddedPropertyName = matchColumn.Any(founder => (0 < founder.ColumnName.IndexOf(uniquePropertyName, StringComparison.OrdinalIgnoreCase)));

            Site.CaptureRequirementIfIsTrue(
                isContainAddedPropertyName,
                641,
                @"[In GetListItemChanges operation] [In GetListItemChanges element] [In viewField field] If the Properties attribute of the ViewFields structure is set to TRUE and the MetaInfo field is referenced in a FieldRef element, then metadata properties related to the list item MUST be returned with the prefix ""ows_MetaInfo_"".");
        }

        /// <summary>
        /// This test case is used to test the GetListItemChanges operation with a specified time span.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC35_GetListItemChanges_SpecifiedTimespan()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Insert items.
            TestSuiteHelper.AddListItems(listId, 2);

            // Get the time spans. Using this token will exclude the above two changes
            // and only the following subsequent changes can be seen.
            GetListItemChangesResponseGetListItemChangesResult getListItemChangesResult =
                this.listswsAdapter.GetListItemChanges(listId, null, null, null);

            Site.CaptureRequirementIfAreEqual<int>(
                2,
                getListItemChangesResult.listitems.data.Length,
                643,
                @"[In GetListItemChanges operation] [In GetListItemChanges element] [In since field] If this value[since] is null, the protocol server MUST return all list items in the list.");

            string timeSpan = getListItemChangesResult.listitems.TimeStamp;

            // Insert items again
            TestSuiteHelper.AddListItems(listId, 2);

            getListItemChangesResult = this.listswsAdapter.GetListItemChanges(listId, null, timeSpan, null);

            // According to Open Specification, the number of data element is 2.
            Site.Debug.AreEqual<int>(
                    2,
                    getListItemChangesResult.listitems.data.Length,
                    "The number data element in the GetListItemChanges should be 2");

            // According to Open Specification, the first data should contain all the changes after the 
            // specified since change token.
            DataTable data = AdapterHelper.ExtractData(getListItemChangesResult.listitems.data[0].Any);
            bool isAfterSince = true;
            foreach (DataRow row in data.Rows)
            {
                isAfterSince = isAfterSince && (DateTime.Parse(timeSpan, CultureInfo.InvariantCulture) <= DateTime.Parse(row["ows_Created"].ToString(), CultureInfo.InvariantCulture));
            }

            Site.CaptureRequirementIfIsTrue(
                isAfterSince,
                1861,
                @"The first rs:data element contains all the inserted and updated list items that have "
                + "occurred subsequent to the specified since parameter.");

            // According to the Open Specification, the first data should contain all the list items.
            // The total number items should be 4, because the two items are inserted twice.
            Site.CaptureRequirementIfAreEqual<int>(
               4,
               int.Parse(getListItemChangesResult.listitems.data[1].ItemCount),
               1862,
               @"[GetListItemChangesResponse]The second rs:data element contains all of the list items "
               + "currently in the list.");

            // Call GetList operation to get all the list field.
            GetListResponseGetListResult listResult;
            listResult = this.listswsAdapter.GetList(listId);

            bool startWithOwsFirst, startWithOwsSecond;
            bool correspondNamesFirst, correspondNamesSecond;
            DataTable firstData = AdapterHelper.ExtractData(getListItemChangesResult.listitems.data[0].Any);
            DataTable secondData = AdapterHelper.ExtractData(getListItemChangesResult.listitems.data[1].Any);

            ValidateOwsPrefixAndAttributeName(
                    firstData,
                    listResult,
                    out startWithOwsFirst,
                    out correspondNamesFirst);
            ValidateOwsPrefixAndAttributeName(
                    secondData,
                    listResult,
                    out startWithOwsSecond,
                    out correspondNamesSecond);

            // If both the data element's z:row attribute names start with ows_, capture the requirement R1863.
            Site.CaptureRequirementIfIsTrue(
                startWithOwsFirst && startWithOwsSecond,
                1863,
                @"[GetListItemChangesResponse]The names of the attributes containing the list "
                + @"item data in inner z:row elements are prefixed by ""ows_"".");

            // If both the data element's z:row attribute names are consistent with fields returned from GetList operation,
            // capture R2370 and R2256.
            Site.CaptureRequirementIfIsTrue(
                correspondNamesFirst && correspondNamesSecond,
                2370,
                @"[GetListItemChangesResponse]The names of the attributes containing the list "
                + "item data in inner z:row elements correspond to the Name attribute in the Field "
                + "elements of GetList.");

            Site.CaptureRequirementIfIsTrue(
                correspondNamesFirst && correspondNamesSecond,
                2256,
                @"[DataDefinition]The names of the attributes of this element correspond to the "
                + "names of fields in the list.");
        }

        /// <summary>
        /// This test case is used to test that the server returns success when ListName is not valid GUID but exists on the server in GetListItemChanges operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC36_GetListItemChanges_Success_ListNameIsInvalidGuid()
        {
            // Construct a ListName.
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Add 1 item.
            List<string> updateId = TestSuiteHelper.AddListItems(listName, 1);

            // Call GetListItemChanges operation to verify the List has been changed on the server.
            GetListItemChangesResponseGetListItemChangesResult result = null;
            result = this.listswsAdapter.GetListItemChanges(
                                                            listName,
                                                            null,
                                                            null,
                                                            null);

            // As description, the first data element contains all changed rows.
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            DataTable data = AdapterHelper.ExtractData(result.listitems.data[0].Any);

            // If the count row is the same with the value of created list number, and the value of "ID" is equal to 
            // the return value from AddListItems, then GetListItemChanges will use the specified list.
            bool isUseList = data.Rows.Count == 1 && Convert.ToString(data.Rows[0][columnName]) == updateId[0];

            #region Capture Requirement R630

            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: data.Rows.Count[{0}],columnName[{1}] for requirement #R630",
                     data.Rows.Count,
                     0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            Site.CaptureRequirementIfIsTrue(
                isUseList,
                630,
                @"[In GetListItemChanges operation] If the specified listName is not a valid GUID, "
                + "check if the listName corresponds to the list title of a list on the site and if so, "
                + "use that list.");

            #endregion

            #region Capture Requirement R631

            string listId = TestSuiteHelper.CreateList();

            // Add 1 item.
            updateId = TestSuiteHelper.AddListItems(listId, 1);

            result = null;
            result = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            null,
                                                            null,
                                                            null);

            // As description, the first data element contains all changed rows.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            data = AdapterHelper.ExtractData(result.listitems.data[0].Any);

            // If the count row is the same with the value of created list number, and the value of "ID" is equal to 
            // the return value from AddListItems, then GetListItemChanges will use the specified list.
            isUseList = data.Rows.Count == 1 && Convert.ToString(data.Rows[0][columnName]) == updateId[0];

            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: data.Rows.Count[{0}],columnName[{1}] for requirement #R631",
                     data.Rows.Count,
                     0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            Site.CaptureRequirementIfIsTrue(
                isUseList,
                631,
                @"[In GetListItemChanges operation] If the specified listName does not "
                + "correspond to the identification of a list on the site, check if the listName "
                + "corresponds to the list title of a list on the site and if so, use that list.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test that the server returns success when ListName is valid GUID in GetListItemChanges operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC37_GetListItemChanges_Success_ListNameIsValidGuid()
        {
            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Add 1 item.
            List<string> updateId = TestSuiteHelper.AddListItems(listId, 1);

            // Read configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();

            // Construct view field element.
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(false, new List<string> { validFieldName });

            // Call GetListItemChanges operation to verify the List has been changed on the server.
            GetListItemChangesResponseGetListItemChangesResult result = null;
            result = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            fields,
                                                            null,
                                                            null);

            // As description, the first data element contains all changed rows.
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            DataTable data = AdapterHelper.ExtractData(result.listitems.data[0].Any);

            // If the count row is the same with the value of create list number, and the value of "ID" is equal to 
            // the return value from AddListItems, then GetListItemChanges will use the specified list.
            bool isUseList = data.Rows.Count == 1 && Convert.ToString(data.Rows[0][columnName]) == updateId[0];

            #region Capture Requirement R629

            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: data.Rows.Count[{0}],columnName[{1}] for requirement #R629",
                     data.Rows.Count,
                     0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            Site.CaptureRequirementIfIsTrue(
                isUseList,
                629,
                @"[In GetListItemChanges operation] If the specified listName is a valid GUID "
                + "and corresponds to the identification of a list on the site, use that list.");

            #endregion

            #region Capture Requirement R634

            // Construct Contains element.
            CamlContains contains = new CamlContains();
            contains.Contains = new CamlContainsContains();
            contains.Contains.FieldRef = new CamlContainsContainsFieldRef();

            // Set the Contains as an invalid field value.
            contains.Contains.FieldRef.Name = validFieldName;
            contains.Contains.Value = new CamlContainsContainsValue();
            contains.Contains.Value.Type = "String";
            contains.Contains.Value.Text = new string[] { invalidFieldName };

            // Call GetListItemChanges operation with the Contains element which filed value is invalid.
            result = null;
            result = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            fields,
                                                            null,
                                                            contains);

            // As description, the first data element contains all changed rows.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            data = AdapterHelper.ExtractData(result.listitems.data[0].Any);

            Site.Assert.IsNull(data, "If the contains parameter is not match any list item, GetListItemChanges operation should also return without any list items");

            // Construct a Since element not match with any list items.
            string since = System.DateTime.Now.AddYears(100).ToString();

            // Call GetListItemChanges operation with a Since element not match with any list items.
            result = null;
            result = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            fields,
                                                            since,
                                                            null);

            // As description, the first data element contains all changed rows.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            data = AdapterHelper.ExtractData(result.listitems.data[0].Any);

            Site.Assert.IsNull(data, "If the Since parameter does not match with any list items, GetListItemChanges operation should also return without any list items");

            // If the Since parameter does not match with any list items, then captrue R634
            Site.CaptureRequirement(
               634,
               @"[In GetListItemChanges operation] Otherwise, all list items in the list specified by the listName parameter that match the criteria based on the since and contains parameters MUST be returned,");

            #endregion
        }

        /// <summary>
        /// This test case is used to validate GetListItemChanges operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC38_GetListItemChanges_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2480, this.Site), @"Test is executed only when R2480Enabled is set to true.");

            // Initialize an invalid ListName structure.
            bool isSoapFaultExist = false;
            try
            {
                // Call GetListItemChanges method with invalid list name: invalid GUID format and not correspond to a list. 
                string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapter.GetListItemChanges(
                                                        invalidListName,
                                                        null,
                                                        null,
                                                        null);
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2480
                isSoapFaultExist = true;
                bool isVerifyR2480 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2480 = string.IsNullOrEmpty(error) && soapException.Detail.InnerText == "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2480 should be covered.
                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: isSoapFaultExist[{0}],error[{1}] for requirement #R2480",
                        isSoapFaultExist,
                        string.IsNullOrEmpty(error) ? "NullOrEmpty" : error);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2480,
                    2480,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<53> Section 3.1.4.21: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            this.Site.Assert.IsTrue(isSoapFaultExist, "Call GetListItemChanges operation with invalid ListName must generate a Soap Fault.");
        }

        /// <summary>
        /// This test case is used to test GetListItemChanges operation when the contained element is not empty.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC39_GetListItemChanges_WithContain()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Add 10 items.
            List<string> addItems = TestSuiteHelper.AddListItems(listId, 10);

            // Construct view field element.
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(
                                                    false,
                                                    new List<string> { validFieldName });

            // Construct Contains element.
            CamlContains contains = new CamlContains();
            contains.Contains = new CamlContainsContains();
            contains.Contains.FieldRef = new CamlContainsContainsFieldRef();

            // There MUST be a column called ID.
            contains.Contains.FieldRef.Name = "ID";
            contains.Contains.Value = new CamlContainsContainsValue();
            contains.Contains.Value.Type = "Counter";

            // Use the first added item ID.
            string containsValue = addItems[0];
            contains.Contains.Value.Text = new string[] { containsValue };

            GetListItemChangesResponseGetListItemChangesResult getListItemChangesResult = null;
            getListItemChangesResult = this.listswsAdapter.GetListItemChanges(
                                                    listId,
                                                    fields,
                                                    null,
                                                    contains);

            // According to the Open Specification, the number of data element is 2.
            Site.Debug.AreEqual<int>(
                    2,
                    getListItemChangesResult.listitems.data.Length,
                    "The number data element in the GetListItemChanges should be 2");

            // According to the Open Specification, the first data element contains all the changed rows.
            DataTable data = AdapterHelper.ExtractData(getListItemChangesResult.listitems.data[0].Any);
            string columnNameId = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            bool isContainOnlyOne = data.Rows.Count == 1;
            bool isIdEqual = Convert.ToString(data.Rows[0][columnNameId]) == addItems[0];

            // If only one row is returned and this row ID equals the expected ID, capture R645.
            Site.CaptureRequirementIfIsTrue(
                isContainOnlyOne && isIdEqual,
                645,
                @"[In GetListItemChanges operation] [In GetListItemChanges element] [In contains field] The protocol server MUST return only those list items satisfying this criterion.");
        }

        /// <summary>
        /// This test case is used to test GetListItemChanges operation when the listName is an empty string. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC40_GetListItemChanges_Fail_EmptyListName()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(5882, this.Site) || Common.IsRequirementEnabled(30183, this.Site),
                @"Test is executed only when R5882Enabled is set to true or R30183Enabled is set to true.");

            // Construct a ListName does not correspond to a list.
            string listName = string.Empty;
            string errorCode = string.Empty;
            bool isSoapRaultReturned = false;

            try
            {
                // Call GetListItemChanges operation to verify the List has been changed on the server.
                this.listswsAdapter.GetListItemChanges(listName, null, null, null);
            }
            catch (SoapException soapException)
            {
                isSoapRaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapException);
            }

            Site.Assert.IsTrue(isSoapRaultReturned, "A SOAP exception should be thrown by the server when invoke the operation 'GetListItemChanges' that the listName is an empty string .");

            // If a SOAP fault is returned with no error code, capture R30182.
            if (Common.IsRequirementEnabled(5882, this.Site))
            {
                bool isR5882Verified = isSoapRaultReturned && string.IsNullOrEmpty(errorCode);

                Site.CaptureRequirementIfIsTrue(
                    isR5882Verified,
                    5882,
                    @"[In Appendix B: Product Behavior] [In GetListItemChanges operation] Implementation does return a SOAP fault with no error code, if the specified listName is an empty string. ( <52> Section 3.1.4.21: In Windows SharePoint Services 3.0 and SharePoint Foundation 2010 this error code is not returned in the SOAP fault.)");
            }

            // If error code "0x82000006" is returned, capture R30183.
            if (Common.IsRequirementEnabled(30183, this.Site))
            {
                bool isR30183Verified = isSoapRaultReturned && errorCode == "0x82000006";

                Site.CaptureRequirementIfIsTrue(
                    isR30183Verified,
                    30183,
                    @"[In GetListItemChanges operation] Implementation does return a SOAP fault with error code 0x82000006, if the specified listName is an empty string.(SharePoint Foundation 2013 and above follow this behavior.)");
            }
        }

        /// <summary>
        ///  This test case is used to test GetListItemChanges operation when the viewFields is not specified by excluding the element or including an empty element.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC41_GetListItemChanges_NullOrEmptyViewFields()
        {
            // Create a list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Add a field to list and default view
            string addedFieldName = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(listId, new List<string> { addedFieldName }, new List<string> { "Text" }, new List<string> { string.Empty });

            // Add 1 list item with random string in 'addedFieldName' and 'existFieldName' field.
            Dictionary<string, string> fieldNameValuePairs = new Dictionary<string, string>();

            // Add value to 'existFieldName' field.
            string existFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);
            fieldNameValuePairs.Add(existFieldName, fieldValue);

            // Add value to 'addedFieldName' field.
            fieldValue = TestSuiteHelper.GenerateRandomString(5);
            fieldNameValuePairs.Add(addedFieldName, fieldValue);

            UpdateListItemsUpdates listItemUpdates = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { fieldNameValuePairs });

            // Invoke UpdateListItems operation to add list item.
            this.listswsAdapter.UpdateListItems(listName, listItemUpdates);

            // Call GetListItemChanges operation to verify the List has been changed on the server.
            GetListItemChangesResponseGetListItemChangesResult getItemChangesResult = null;

            // 'viewFields' is not specified by excluding the element 
            getItemChangesResult = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            null,
                                                            null,
                                                            null);

            // "ows_" is required prefix in [MS-LISTSWS]
            string expectedPrefix = "ows_";

            bool isFieldInDefaultViewReturned = false;
            bool isFieldNotInDefaultViewReturned = false;

            // Find fields in response
            // According to TD, all fields in the default view MUST be returned from the list item.
            // 'addedFieldName' is added in default view, so it must in the response.
            // 'fieldName' field is not added in default view, so it must not in the response.
            foreach (DataDefinition data in getItemChangesResult.listitems.data)
            {
                XmlNode[] anyItems = data.Any;
                DataTable zrowData = AdapterHelper.ExtractData(anyItems);

                if (!isFieldInDefaultViewReturned)
                {
                    // Check if 'addedFieldName' field in result
                    var matchColumn = from DataColumn columnItem in zrowData.Columns
                                      where columnItem.ColumnName == expectedPrefix + addedFieldName
                                      select columnItem;

                    isFieldInDefaultViewReturned = matchColumn.Any(founder => founder.ColumnName == expectedPrefix + addedFieldName);
                }

                if (!isFieldNotInDefaultViewReturned)
                {
                    // Check if 'fieldName' field in result
                    var matchColumn = from DataColumn columnItem in zrowData.Columns
                                      where columnItem.ColumnName == expectedPrefix + existFieldName
                                      select columnItem;

                    isFieldNotInDefaultViewReturned = matchColumn.Any(founder => founder.ColumnName == expectedPrefix + existFieldName);
                }
            }

            bool isExcludingElementVerified = isFieldInDefaultViewReturned && !isFieldNotInDefaultViewReturned;

            // 'viewFields' is not specified by including an empty element
            CamlViewFields viewFields = new CamlViewFields();

            getItemChangesResult = this.listswsAdapter.GetListItemChanges(
                                                            listId,
                                                            viewFields,
                                                            null,
                                                            null);

            isFieldInDefaultViewReturned = false;
            isFieldNotInDefaultViewReturned = false;

            // Find fields in response
            // According to TD, all fields in the default view MUST be returned from the list item.
            // 'addedFieldName' is added in default view, so it must in the response.
            // 'fieldName' field is not added in default view, so it must not in the response.
            foreach (DataDefinition data in getItemChangesResult.listitems.data)
            {
                XmlNode[] anyItems = data.Any;
                DataTable zrowData = AdapterHelper.ExtractData(anyItems);

                if (!isFieldInDefaultViewReturned)
                {
                    // Check if 'addedFieldName' field in result
                    var matchColumn = from DataColumn columnItem in zrowData.Columns
                                      where columnItem.ColumnName == expectedPrefix + addedFieldName
                                      select columnItem;

                    isFieldInDefaultViewReturned = matchColumn.Any(founder => founder.ColumnName == expectedPrefix + addedFieldName);
                }

                if (!isFieldNotInDefaultViewReturned)
                {
                    // Check if 'fieldName' field in result
                    var matchColumn = from DataColumn columnItem in zrowData.Columns
                                      where columnItem.ColumnName == expectedPrefix + existFieldName
                                      select columnItem;

                    isFieldNotInDefaultViewReturned = matchColumn.Any(founder => founder.ColumnName == expectedPrefix + existFieldName);
                }
            }

            bool isEmptyElementVerified = isFieldInDefaultViewReturned && !isFieldNotInDefaultViewReturned;

            // If both excluding element and empty element requests are verified, it means
            // whether by excluding the element or by including an empty element, 
            // all fields in the default view MUST be returned from the list item.
            // Then R640 can be verified.
            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: isExcludingElementVerified[{0}],isEmptyElementVerified[{1}] for requirement #R640",
                     isExcludingElementVerified,
                     isEmptyElementVerified);

            bool isR640Verified = isExcludingElementVerified && isEmptyElementVerified;

            Site.CaptureRequirementIfIsTrue(
                isR640Verified,
                640,
                @"[In GetListItemChanges operation] [In GetListItemChanges element] [In viewFields field]If this is not specified, whether by excluding the element or by including an empty element, all fields in the default view MUST be returned from the list item.");
        }

        #endregion

        #region GetListItems
        /// <summary>
        /// This test case is used to verify the GetListItemChanges operation when both listName and viewName are empty string.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC42_GetListItems_EmptyListNameAndViewName()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(5894, this.Site) || Common.IsRequirementEnabled(30193, this.Site),
                @"Test is executed only when R5894Enabled is set to true or R30193Enabled is set to true.");

            // Construct a ListName does not correspond to a list.
            string listName = string.Empty;
            string viewName = string.Empty;
            string errorCode = string.Empty;
            bool isSoapRaultReturned = false;

            try
            {
                // Call GetListItemChanges operation to verify the List has been changed on the server.
                this.listswsAdapter.GetListItems(listName, viewName, null, null, null, null, null);
            }
            catch (SoapException soapException)
            {
                isSoapRaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapException);
            }

            Site.Assert.IsTrue(isSoapRaultReturned, "A SOAP exception should be thrown by the server when invoke the operation 'GetListItems' when both listName and viewName are empty string.");

            // If a SOAP fault is returned with no error code, capture R30192.
            if (Common.IsRequirementEnabled(5894, this.Site))
            {
                bool isR5894Verified = isSoapRaultReturned && string.IsNullOrEmpty(errorCode);

                Site.CaptureRequirementIfIsTrue(
                    isR5894Verified,
                    5894,
                    @"[In Appendix B: Product Behavior][In GetListItems operation] Implementation does return a SOAP fault with no error code, if listName and viewName are both empty. (<65> Section 3.1.4.24: In Windows SharePoint Services 3.0 and SharePoint Foundation 2010, this error code is not returned in the SOAP fault.)");
            }

            // If error code "0x82000006" is returned, capture R30193.
            if (Common.IsRequirementEnabled(30193, this.Site))
            {
                bool isR30193Verified = isSoapRaultReturned && errorCode == "0x82000006";

                Site.CaptureRequirementIfIsTrue(
                    isR30193Verified,
                    30193,
                    @"[In GetListItems operation] Implementation does return a SOAP fault with error code 0x82000006, if listName and viewName are both empty.(SharePoint Foundation 2013 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the value of element listName does not correspond to the list. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC43_GetListItems_FilesOnly()
        {
            string listTitle = TestSuiteHelper.GetUniqueListName();
            int docLibraryTemplateId = (int)TemplateType.Document_Library;
            string listId = TestSuiteHelper.CreateList(listTitle, docLibraryTemplateId);

            // Upload the file to the document library
            IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();
            string absoluteFileUrl = sutControlAdapter.UploadFile(listTitle);
            Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        listTitle,
                        absoluteFileUrl);

            // Get Add file item in the document library
            GetListItemsResponseGetListItemsResult getListItemsForFileItem = null;
            getListItemsForFileItem = this.listswsAdapter.GetListItems(
                                                   listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);

            if (null == getListItemsForFileItem || null == getListItemsForFileItem.listitems || null == getListItemsForFileItem.listitems.data
                || !getListItemsForFileItem.listitems.data.ItemCount.Equals("1"))
            {
                this.Site.Assert.Fail("Upload a file to SUT fail.");
            }

            DataTable rowDataOfGetListitem = AdapterHelper.ExtractData(getListItemsForFileItem.listitems.data.Any);

            // Get the current Folder item id, it is required column in [MS-LISTSWS]
            string documentItemId = rowDataOfGetListitem.Rows[0]["ows_ID"].ToString();

            // add a folder.
            string uniqueFolderName = TestSuiteHelper.GetUniqueFolderName();
            UpdateListItemsUpdates listItemUpdates = new UpdateListItemsUpdates();
            listItemUpdates.Batch = new UpdateListItemsUpdatesBatch();
            listItemUpdates.Batch.Method = new UpdateListItemsUpdatesBatchMethod[1];
            listItemUpdates.Batch.Method[0] = new UpdateListItemsUpdatesBatchMethod();
            listItemUpdates.Batch.Method[0].ID = (uint)0;
            listItemUpdates.Batch.Method[0].Cmd = MethodCmdEnum.New;
            listItemUpdates.Batch.Method[0].Field = new UpdateListItemsUpdatesBatchMethodField[2];
            listItemUpdates.Batch.Method[0].Field[0] = new UpdateListItemsUpdatesBatchMethodField();
            listItemUpdates.Batch.Method[0].Field[0].Name = "FSObjType";
            listItemUpdates.Batch.Method[0].Field[0].Value = "1";
            listItemUpdates.Batch.Method[0].Field[1] = new UpdateListItemsUpdatesBatchMethodField();
            listItemUpdates.Batch.Method[0].Field[1].Name = "BaseName";
            listItemUpdates.Batch.Method[0].Field[1].Value = uniqueFolderName;
            UpdateListItemsResponseUpdateListItemsResult updateListResult = null;
            updateListResult = this.listswsAdapter.UpdateListItems(listId, listItemUpdates);

            if (null == updateListResult || null == updateListResult.Results || null == updateListResult.Results
             || updateListResult.Results.Length != 1)
            {
                this.Site.Assert.Fail("Could not create a folder to current list[{0}].", listId);
            }

            if (null == updateListResult.Results[0].Any)
            {
                this.Site.Assert.Fail("Could not get the created folder item[{0}] information.", uniqueFolderName);
            }

            DataTable rowDataOfFolderItems = AdapterHelper.ExtractData(updateListResult.Results[0].Any);

            // Get the current Folder item id, it is required column in [MS-LISTSWS]
            this.Site.Assert.AreEqual(1, rowDataOfFolderItems.Rows.Count, "Could not get the created folder item detail information");
            string folderItemId = rowDataOfFolderItems.Rows[0]["ows_ID"].ToString();

            // If R14021 is enable, Protocol SUT is able to return getListResult.List.EnableFolderCreation attribute.
            // If this attribute is return by SUT and its value is true, then capture R14021
            if (Common.IsRequirementEnabled(14021, this.Site))
            {
                GetListResponseGetListResult getListResult = null;
                getListResult = this.listswsAdapter.GetList(listId);
                this.Site.CaptureRequirementIfIsTrue(
                  bool.TrueString.Equals(getListResult.List.EnableFolderCreation, StringComparison.OrdinalIgnoreCase),
                  14021,
                  @"[ListDefinitionCT.EnableFolderCreation: ] [If the attribute is returned by SUT]Implementation does True if folder creation is enabled on the list.(Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }

            // Set the EnumViewAttributes to "FilesOnly" value.
            CamlQueryOptions queryOption = CreateDefaultCamlQueryOptions();
            queryOption.QueryOptions.ViewAttributes.Scope = EnumViewAttributes.FilesOnly;

            GetListItemsResponseGetListItemsResult getListItems = null;
            getListItems = this.listswsAdapter.GetListItems(
                                                listId,
                                                null,
                                                null,
                                                null,
                                                null,
                                                queryOption,
                                                null);

            if (null == getListItems || null == getListItems.listitems || null == getListItems.listitems.data
                || null == getListItems.listitems.data.Any)
            {
                this.Site.Assert.Fail("Could not get the valid response for GetListItems operation");
            }

            XmlNode[] rowInformation = getListItems.listitems.data.Any;
            DataTable rowDataOfGtListItems = AdapterHelper.ExtractData(rowInformation);

            // Verify whether the response contains added folder item 
            var rowItemsOfFolder = from DataRow folderitem in rowDataOfGtListItems.Rows
                                   where folderitem["ows_Id"].ToString().Equals(folderItemId, StringComparison.OrdinalIgnoreCase)
                                   select folderitem;

            bool isResultContainedFolderItem = rowItemsOfFolder.Count() > 0;
            this.Site.Assert.IsFalse(isResultContainedFolderItem, "The Response of GetListItems operation contain a folder type item");

            // Verify whether the response contains added document. 
            var rowItemsOfDocument = from DataRow documentitem in rowDataOfGtListItems.Rows
                                     where documentitem["ows_Id"].ToString().Equals(documentItemId, StringComparison.OrdinalIgnoreCase)
                                     select documentitem;

            // Because only upload a file, If only one file type item is return and document count equals to 1, capture the R280
            this.Site.CaptureRequirementIfAreEqual(
                 1,
                 rowItemsOfDocument.Count(),
                 280,
                 @"[In EnumViewAttributes] the FilesOnly option MUST return only the documents and not the folders.");
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the
        /// "viewName" element is not a GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC44_GetListItems_InvalidViewNameWithNotGUID()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke GetListItems operation when the value of element "viewName" is not a GUID.
            // Invoke GetListItems operation when the value of element "viewName" is not a GUID.
            // If protocol SUT return a SOAP fault with error code 0x82000001, then will be able to capture R729 and R1912. 
            string strInvalidViewName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool isCaptureR729_R1912 = false;
            string strErrorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetListItems(strList_GUID, strInvalidViewName, null, null, null, null, null);
            }
            catch (SoapException soapEx)
            {
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
                if ("0x82000001".Equals(strErrorCode, StringComparison.OrdinalIgnoreCase))
                {
                    isCaptureR729_R1912 = true;
                }
                else
                {
                    isCaptureR729_R1912 = false;
                    string strCaseFail_Info = string.Format("CaseFail: The actual error code is {0}, but the expected error code is 0x82000001", strErrorCode);
                    Site.Log.Add(LogEntryKind.TestFailed, strCaseFail_Info);
                }
            }
            #endregion

            #region Capture Requirement #729 #1912 when protocol SUT returns SOAP fault with error code 0x82000001 in the response of "GetListItems".

            Site.Log.Add(
              LogEntryKind.Debug,
              "Capture requirements #R729 #R1912: CaptureR729_R1912[{0}]; ErrorCode[{1}]",
              isCaptureR729_R1912,
              string.IsNullOrEmpty(strErrorCode) ? "NullOrEmpty" : strErrorCode);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR729_R1912,
                729,
                @"[In GetListItems operation] If the viewName parameter is not a GUID, the protocol server MUST "
                + "return a SOAP fault with error code 0x82000001.");
            Site.CaptureRequirementIfIsTrue(
                isCaptureR729_R1912,
                1912,
                @"[In GetListItems operation] [If the viewName parameter is not a GUID, the protocol server MUST "
                + "return a SOAP fault with error code 0x82000001.]This indicates that the viewName parameter is "
                + "missing or invalid.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test GetListItems operation when the value of element listName does not correspond to the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC45_GetListItems_ListNameNotCorrespond()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2764, this.Site), @"Test is executed only when R2764Enabled is set to true.");

            #region Invoke GetListItems operation when the value of element "listName" does not correspond to the list.
            // Invoke GetListItems operation when the value of element "listName" does not correspond to the list.
            // If protocol SUT returns a SOAP fault with error code "0x82000006", capture R2764. 
            string strInvalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
            bool isSoapFaultGenerated = false;
            string strErrorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetListItems(strInvalidListName, null, null, null, null, null, null);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultGenerated = true;
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }

            // Capture Requirement #2764 when protocol SUT returns SOAP fault error code "0x82000006" in the response of "GetListItems".
            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isSoapFaultGenerated[{0}], ErrorCode[{1}] for requirement #R2764",
                    isSoapFaultGenerated,
                    string.IsNullOrEmpty(strErrorCode) ? "NullOrEmpty" : strErrorCode);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2764
            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated,
                2764,
                @"[In GetListItems operation] Implementation does return a SOAP fault with error code 0x82000006, if listName is not empty and does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the 
        /// value of "viewName" element is a GUID that does not correspond to 
        /// an existing view in the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC46_GetListItems_NonExistentViewName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke GetListItems operation when the value of element "viewName" does not correspond to an exist view in the list.
            // Invoke GetListItems operation when the value of element "viewName" does not correspond to an existing view in the list.
            // If protocol SUT returns a SOAP fault, capture R730. 
            string strInvalidViewGUID = Guid.NewGuid().ToString();
            bool isSoapFaultGenerated = false;
            string expectErrorCode = "0x82000005";
            string errorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetListItems(strList_GUID, strInvalidViewGUID, null, null, null, null, null);
            }
            catch (SoapException ex)
            {
                isSoapFaultGenerated = true;
                errorCode = TestSuiteHelper.GetErrorCode(ex);
            }
            #endregion

            #region Capture Requirement #730 and #1913 when protocol SUT returns SOAP fault with error code "0x82000005" in the response of "GetListItems".

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated,
                730,
                @"[In GetListItems operation] If the viewName parameter does not correspond to a view that exists and "
                + "each of the optional query, viewFields, queryOptions, and rowLimit parameters are not specified, "
                + "the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                expectErrorCode.Equals(errorCode, StringComparison.OrdinalIgnoreCase),
                1913,
                @"[In GetListItems operation] [If the viewName parameter does not correspond to a view that exists and each of the optional query, viewFields, queryOptions, and rowLimit parameters are not specified, the protocol server MUST return a SOAP fault.] The error code is 0x82000005.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify all field names start with the prefix _ows in GetListItems operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC47_GetListItems_Prefix()
        {
            string listId = TestSuiteHelper.CreateList();

            // Insert 100 items.
            TestSuiteHelper.AddListItems(listId, 100);

            // Call the GetListItems with unspecified view name but row limit specified.
            string rowLimit = "2";
            GetListItemsResponseGetListItemsResult listItemsWithRowLimit = null;
            listItemsWithRowLimit = this.listswsAdapter.GetListItems(
                                                        listId,
                                                        null,
                                                        null,
                                                        null,
                                                        rowLimit,
                                                        null,
                                                        null);
            DataTable data = AdapterHelper.ExtractData(listItemsWithRowLimit.listitems.data.Any);
            int rowCount = data.Rows.Count;
            bool isMatchRowLimit = rowCount <= Convert.ToInt32(rowLimit);

            // If returned data rows do not exceed the specified row limit, capture R1929.
            Site.CaptureRequirementIfIsTrue(
                        isMatchRowLimit,
                        1929,
                        @"[GetListItems]rowLimit : Specifies the maximum number of rows of data to return in the response.");

            GetListResponseGetListResult listResult;
            listResult = this.listswsAdapter.GetList(listId);

            // Verify that the names of the attributes containing the list item data in this response 
            // correspond to the Name attribute in the Field elements of GetList 
            // and are prefixed by "ows_".
            bool startWithOws;
            bool correspondNames;
            ValidateOwsPrefixAndAttributeName(
                    data,
                    listResult,
                    out startWithOws,
                    out correspondNames);

            // If both of startWithOws and correspondNames are true, capture the following requirements.
            Site.CaptureRequirementIfIsTrue(
                startWithOws,
                1937,
                @"[GetListItemsResponse]The names of the attributes containing the list item "
                + @"data in this response are prefixed by ""ows_"".");

            Site.CaptureRequirementIfIsTrue(
                correspondNames,
                2369,
                @"[GetListItemsResponse]The names of the attributes[GetListItemsResult] containing the list item "
                    + "data in this response correspond to the Name attribute in the Field elements of GetList.");

            Site.CaptureRequirementIfIsTrue(
                correspondNames,
                2256,
                @"[DataDefinition]The names of the attributes of this element correspond to "
                + "the names of fields in the list.");
        }

        /// <summary>
        ///  The test case is used to verify GetListItems Operation when querying the MetaInfo field and Properties attribute equals to true.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC48_GetListItems_PrefixForMetaInfoProperty()
        {
            // Create list.
            string listGuid = TestSuiteHelper.CreateList();

            #region Add a list item and add property on the metaInfo field

            // Add a property on the metaInfo field
            string uniquePropertyName = TestSuiteHelper.GetUniqueMetaInfoFieldPropertyName();
            string propertyValue = Guid.NewGuid().ToString("N");
            TestSuiteHelper.AddListItemWithMetaInfoProperty(listGuid, uniquePropertyName, propertyValue);
            #endregion

            // Get the ViewFields whose Properties is true and reference field is MetaInfo.
            // "MetaInfo" is required Field in [MS-LISTSWS]
            string metaInfoFieldName = "MetaInfo";
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { metaInfoFieldName });
            GetListItemsResponseGetListItemsResult getListItemResult = null;
            getListItemResult = this.listswsAdapter.GetListItems(
                            listGuid,
                            null,
                            null,
                            viewFields,
                            null,
                            null,
                            null);

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R743
            if (getListItemResult == null || getListItemResult.listitems == null
                || getListItemResult.listitems.data == null || getListItemResult.listitems.data.Any == null)
            {
                this.Site.Assert.Fail("Get list item response error.");
            }

            if (getListItemResult.listitems.data.Any.Length == 0)
            {
                this.Site.Assert.Fail("Get list item data error");
            }

            // A return value represents the schema definition which is read from specified xsd file.
            XmlNode[] anyItems = getListItemResult.listitems.data.Any;

            // "ows_MetaInfo_" is required prefix in [MS-LISTSWS]
            string expectedPrefix = "ows_MetaInfo_";
            DataTable zrowData = AdapterHelper.ExtractData(anyItems);

            // Pick up the items which begin with "ows_MetaInfo_"
            var matchColumn = from DataColumn columnItem in zrowData.Columns
                              where 0 == columnItem.ColumnName.IndexOf(expectedPrefix, StringComparison.OrdinalIgnoreCase)
                              select columnItem;

            // Pick up the column items contains added property name.
            bool isContainAddedPropertyName = matchColumn.Any(founder => (0 < founder.ColumnName.IndexOf(uniquePropertyName, StringComparison.OrdinalIgnoreCase)));

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R743
            Site.CaptureRequirementIfIsTrue(
               isContainAddedPropertyName,
                743,
                @"[In GetListitems operation] [In GetListitems element] [In viewFields element] If the Properties attribute of the ViewFields structure is set to TRUE and the MetaInfo field is referenced in FieldRef element, then metadata properties related to the list item MUST be returned with the prefix ""ows_MetaInfo_"".");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R46
            Site.CaptureRequirementIfIsTrue(
               isContainAddedPropertyName,
               46,
               @"[In CamlViewFields] [ViewFields.Properties] If True and the MetaInfo field is referenced in a FieldRef element, then metadata properties related to the list item are returned in the protocol server response with the prefix ""ows_MetaInfo_"".");
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation with "rowLimit" parameter.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC49_GetListItems_RowLimit()
        {
            // Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);

            // Invoke GetListAndView operation to get valid GUID value for default view.
            GetListAndViewResponseGetListAndViewResult getListAndView = null;
            getListAndView = this.listswsAdapter.GetListAndView(strList_GUID, string.Empty);
            Site.Assert.IsNotNull(getListAndView, "The object \"getListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView, "The object \"getListAndView.ListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View, "The object \"getListAndView.ListAndView.View\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View.Name, "The object \"getListAndView.ListAndView.View.Name\" should not be null !");
            string strViewGUID = getListAndView.ListAndView.View.Name;

            // Get the rowLimit of default view.
            int defaultRowLimit = getListAndView.ListAndView.View.RowLimit.Value;

            // Invoke UpdateListItems operation to add new list items, the number of new items is greater than the default view rowLimit.
            TestSuiteHelper.AddListItems(strList_GUID, defaultRowLimit + 2);

            // Invoke GetListItems operation excluding the "rowLimit" element
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(strList_GUID, strViewGUID, null, null, null, null, null);

            int actualItemCount = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any).Rows.Count;

            // If the number of returned list items equal to the default rowLimit, capture 2368.
            Site.CaptureRequirementIfAreEqual<int>(
                    defaultRowLimit,
                    actualItemCount,
                    2368,
                    @"[In GetListitems operation] [In GetListitems element] [In rowLimit element] If this is not specified by excluding the element, 
                    the default row limit for the view specified by the viewName parameter is used.");

            // Invoke GetListItems operation with empty "rowLimit" element.
            string rowLimit = string.Empty;
            getListItemsResult = this.listswsAdapter.GetListItems(strList_GUID, strViewGUID, null, null, rowLimit, null, null);

            actualItemCount = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any).Rows.Count;

            // If the number of returned list items equal to the default rowLimit, capture 744.
            Site.CaptureRequirementIfAreEqual<int>(
                    defaultRowLimit,
                    actualItemCount,
                    744,
                    @"[In GetListitems operation] [In GetListitems element] [In rowLimit element] If this is not specified by passing an empty element, 
                    the default row limit for the view specified by the viewName parameter is used.");

            // Invoke GetListItems operation with specified "rowLimit" element, the value of "rowLimit" is greater than the default view rowLimit
            // and less than the total number of the items in list.
            rowLimit = (defaultRowLimit + 1).ToString();
            getListItemsResult = this.listswsAdapter.GetListItems(strList_GUID, strViewGUID, null, null, rowLimit, null, null);

            actualItemCount = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any).Rows.Count;

            // If the number of returned list items less than or equal to the 'rowLimit' parameter, capture R746
            bool isR746Verified = actualItemCount <= int.Parse(rowLimit);
            Site.CaptureRequirementIfIsTrue(
                    isR746Verified,
                    746,
                    @"[In GetListitems operation] [In GetListitems element] [In rowLimit element] 
                    The protocol server MUST NOT return more list items than this parameter specifies.");
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the "listName" element is not a valid GUID but its value corresponds to the title of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC50_GetListItems_SucceedWithNonGuidFormatListName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke GetList operation to get the title of the new generic list.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(strList_GUID);
            Site.Assert.IsNotNull(listDef, "The object \"listDef\" should not be null !");
            Site.Assert.IsNotNull(listDef.Title, "The object \"listDef.Title\" should not be null !");
            string strList_Title = listDef.Title;
            #endregion

            #region Invoke UpdateListItems operation to add three new list items in the new generic list.
            TestSuiteHelper.AddListItems(strList_GUID, 3);
            #endregion

            #region Invoke GetListItems operation with the title of the generic list as the value of element "listName".
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { AdapterHelper.FieldAuthorName });
            CamlQueryOptions queryOptions = CreateDefaultCamlQueryOptions();

            getListItemsResult = this.listswsAdapter.GetListItems(strList_Title, null, null, viewFields, null, queryOptions, null);
            #endregion

            #region Capture Requirements #725 #726, when protocol SUT returns successful response of "GetListItems" operation.
            // If protocol SUT returns the response of GetListItems without any Soap Fault, capture the R725 and R726.
            Site.CaptureRequirementIfIsNotNull(
                getListItemsResult,
                725,
                @"[In GetListItems operation] If the specified listName is not a valid GUID, "
                + "check if the listName corresponds to the list title of a list on the site and if so, use that list");
            Site.CaptureRequirementIfIsNotNull(
                getListItemsResult,
                726,
                @"[In GetListItems operation] If the specified listName does not correspond to the identification "
                + "of a list on the site, check if the listName corresponds to the list title of a list on the site and "
                + "if so, use that list");
            #endregion

            #region Capture Requirement #733 when we get three "z:row" elements in the successful response of "GetListItems" operation.
            bool isCaptureR733 = false;
            Site.Assert.IsNotNull(getListItemsResult, "The object \"getListItemsResult\" is null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems, "The object \"getListItemsResult.listitems\" is null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data, "The object \"getListItemsResult.listitems.data\" is null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data.ItemCount, "The object \"getListItemsResult.listitems.data.ItemCount\" is null !");
            int countZRow = int.Parse(getListItemsResult.listitems.data.ItemCount);
            if (countZRow == 3)
            {
                isCaptureR733 = true;
            }

            Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: getListItemsResult.listitems.data.ItemCount[{0}] for requirement #R733",
                         string.IsNullOrEmpty(getListItemsResult.listitems.data.ItemCount) ? "NullOrEmpty" : getListItemsResult.listitems.data.ItemCount);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR733,
                733,
                @"[In GetListitems operation] [If the specified listName corresponds to the identification or "
                + "the list title of a list on the site, the viewName is a GUID which correspond to a exist view "
                + "or each of the query, viewFields, queryOptions and rowLimit parameters are specified] all list "
                + "items in the list matching the specified criteria MUST be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the
        /// "listName" element is a valid GUID.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC51_GetListItems_SucceedWithValidGuidListName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke UpdateListItems operation to add three new list items in the new generic list.
            TestSuiteHelper.AddListItems(strList_GUID, 3);
            #endregion

            #region Invoke GetListItems operation with the valid GUID of the generic list as the value of element "listName".
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(strList_GUID, null, null, null, null, null, null);
            #endregion

            #region Capture Requirements #724 when protocol SUT returns successful response of "GetListItems" operation.
            // If protocol SUT returns the response of GetListItems without any Soap Fault, capture the R724.
            Site.CaptureRequirementIfIsNotNull(
                getListItemsResult,
                724,
                @"[In GetListItems operation] If the specified listName is a valid GUID and corresponds to "
                + "the identification of a list on the site, use that list.");
            #endregion

            #region Capture Requirement #733 when protocol SUT return three "z:row" elements in the successful response of "GetListItems" operation.
            bool isCaptureR733 = false;
            Site.Assert.IsNotNull(getListItemsResult, "The object \"getListItemsResult\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems, "The object \"getListItemsResult.listitems\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data, "The object \"getListItemsResult.listitems.data\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data.ItemCount, "The object \"getListItemsResult.listitems.data.ItemCount\" should not be null !");
            int countZRow = int.Parse(getListItemsResult.listitems.data.ItemCount);
            if (countZRow == 3)
            {
                isCaptureR733 = true;
            }

            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: getListItemsResult.listitems.data.ItemCount[{0}] for requirement #R733",
                string.IsNullOrEmpty(getListItemsResult.listitems.data.ItemCount) ? "NullOrEmpty" : getListItemsResult.listitems.data.ItemCount);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR733,
                733,
                @"[In GetListitems operation] [If the specified listName corresponds to the identification or "
                + "the list title of a list on the site, the viewName is a GUID which correspond to a exist view "
                + "or each of the query, viewFields, queryOptions and rowLimit parameters are specified] all list "
                + "items in the list matching the specified criteria MUST be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation when the
        /// "viewName" element is a valid GUID which corresponds to an existing view.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC52_GetListItems_SucceedWithValidGuidViewName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke UpdateListItems operation to add three new list items in the new generic list.
            TestSuiteHelper.AddListItems(strList_GUID, 3);
            #endregion

            #region Invoke GetListAndView operation to get valid GUID value for default view.
            GetListAndViewResponseGetListAndViewResult getListAndView = null;
            getListAndView = this.listswsAdapter.GetListAndView(strList_GUID, string.Empty);
            Site.Assert.IsNotNull(getListAndView, "The object \"getListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView, "The object \"getListAndView.ListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View, "The object \"getListAndView.ListAndView.View\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View.Name, "The object \"getListAndView.ListAndView.View.Name\" should not be null !");
            string strViewGUID = getListAndView.ListAndView.View.Name;
            #endregion

            #region Invoke GetListItems operation with the GUID of existed view as the value of element "viewName".
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            getListItemsResult = this.listswsAdapter.GetListItems(strList_GUID, strViewGUID, null, null, null, null, null);
            #endregion

            #region Capture Requirement #733 when protocol SUT returns three "z:row" elements in the successful response of "GetListItems" operation.
            bool isCaptureR733 = false;
            Site.Assert.IsNotNull(getListItemsResult, "The object \"getListItemsResult\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems, "The object \"getListItemsResult.listitems\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data, "The object \"getListItemsResult.listitems.data\" should not be null !");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data.ItemCount, "The object \"getListItemsResult.listitems.data.ItemCount\" should not be null !");
            int countZRow = int.Parse(getListItemsResult.listitems.data.ItemCount);
            if (countZRow == 3)
            {
                isCaptureR733 = true;
            }

            Site.Log.Add(
               LogEntryKind.Debug,
               "The actual value: getListItemsResult.listitems.data.ItemCount[{0}] for requirement #R733",
               string.IsNullOrEmpty(getListItemsResult.listitems.data.ItemCount) ? "NullOrEmpty" : getListItemsResult.listitems.data.ItemCount);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR733,
                733,
                @"[In GetListitems operation] [If the specified listName corresponds to the identification or "
                + "the list title of a list on the site, the viewName is a GUID which correspond to a exist view "
                + "or each of the query, viewFields, queryOptions and rowLimit parameters are specified] all list "
                + "items in the list matching the specified criteria MUST be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetListItems operation with the "ViewFields" element.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC53_GetListItems_ViewFields()
        {
            // Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);

            // Invoke UpdateListItems operation to add one new list item in the new generic list.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            List<Dictionary<string, string>> insertItems = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> insertCmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> insertItem = new Dictionary<string, string>();
            insertItem.Add(validFieldName, fieldValue);
            insertItems.Add(insertItem);
            insertCmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(insertCmds, insertItems, OnErrorEnum.Continue);

            // Call UpdateListItems to insert the item using List GUID.
            UpdateListItemsResponseUpdateListItemsResult addItemResult = null;
            addItemResult = this.listswsAdapter.UpdateListItems(strList_GUID, updates);

            Site.Assert.AreEqual<int>(
                        1,
                        addItemResult.Results.Length,
                        "The number of result element in the UpdateListItems response MUST be equal to the number of Method in the request.");
            DataTable addItemData = AdapterHelper.ExtractData(addItemResult.Results[0].Any);

            // Call GetListItems operation with empty ViewFields element.
            CamlViewFields emptyViewFields = new CamlViewFields();
            emptyViewFields.ViewFields = new CamlViewFieldsViewFields();

            GetListItemsResponseGetListItemsResult getListResult = null;
            getListResult = this.listswsAdapter.GetListItems(strList_GUID, null, null, emptyViewFields, null, null, null);
            DataTable getListData = AdapterHelper.ExtractData(getListResult.listitems.data.Any);

            // If all fields in GetListItems response exists in UpdateListItems response, capture R42.
            bool isR42Verified = true;
            foreach (DataColumn column in getListData.Columns)
            {
                isR42Verified = isR42Verified && addItemData.Columns.Contains(column.ColumnName);
            }

            this.Site.CaptureRequirementIfIsTrue(
                isR42Verified,
                42,
                @"[In CamlViewFields] [ViewFields] An empty element <ViewFields /> specifies that all fields of the list item MUST be returned in the protocol server response.");

            // Get default view GUID.
            GetListAndViewResponseGetListAndViewResult getListAndView = null;
            getListAndView = this.listswsAdapter.GetListAndView(strList_GUID, string.Empty);
            Site.Assert.IsNotNull(getListAndView, "The object \"getListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView, "The object \"getListAndView.ListAndView\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View, "The object \"getListAndView.ListAndView.View\" should not be null !");
            Site.Assert.IsNotNull(getListAndView.ListAndView.View.Name, "The object \"getListAndView.ListAndView.View.Name\" should not be null !");
            string strViewGUID = getListAndView.ListAndView.View.Name;

            // Create query, viewFields, rowLimit, queryOptions parameters.
            GetListItemsQuery query = new GetListItemsQuery();
            query.Query = new CamlQueryRoot();
            query.Query.OrderBy = new OrderByDefinition();
            query.Query.OrderBy.FieldRef = new FieldRefDefinitionOrderBy[1];
            query.Query.OrderBy.FieldRef[0] = new FieldRefDefinitionOrderBy();
            query.Query.OrderBy.FieldRef[0].Name = AdapterHelper.FieldIDName;
            query.Query.OrderBy.FieldRef[0].Ascending = "TRUE";

            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { validFieldName });
            string rowLimit = "1";
            CamlQueryOptions queryOption = CreateDefaultCamlQueryOptions();

            // Call GetListItems with default view GUID and valid query, viewFields, rowLimit, queryOptions parameters.
            GetListItemsResponseGetListItemsResult defaultViewResult = null;
            defaultViewResult = this.listswsAdapter.GetListItems(strList_GUID, strViewGUID, query, viewFields, rowLimit, queryOption, null);

            // Call GetListItems without viewName and with valid query, viewFields, rowLimit, queryOptions parameters.
            GetListItemsResponseGetListItemsResult nullViewResult = null;
            nullViewResult = this.listswsAdapter.GetListItems(strList_GUID, null, query, viewFields, rowLimit, queryOption, null);

            // If the same response returned, capture R2272.
            bool isSame = TestSuiteHelper.DeepCompare(defaultViewResult, nullViewResult);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2272,
                @"If each of the query, viewFields, queryOptions, and rowLimit parameters are specified, the response of protocol server will be same even if viewName parameter are different.");
        }

        /// <summary>
        /// This test case is used to verify GetListItems operation in Windows SharePoint Services 3.0 when 
        /// ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC54_GetListItems_WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2489, this.Site), @"Test is executed only when R2489Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            bool isSoapFault = false;
            try
            {
                // Call GetListItems method to retrieve details about list items in a list that satisfy specified criteria.
                this.listswsAdapter.GetListItems(
                                                    invalidListName,
                                                    string.Empty,
                                                    null,
                                                    null,
                                                    string.Empty,
                                                    null,
                                                    string.Empty);
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2489
                isSoapFault = true;
                bool isVerifyR2489 = false;
                string error = TestSuiteHelper.GetErrorCode(soapException);
                isVerifyR2489 = string.IsNullOrEmpty(error)
                    && soapException.Detail.InnerText.Equals(
                    "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).",
                         StringComparison.OrdinalIgnoreCase);

                // If the value of the error is null and the detailed message of SoapException is equal to the 
                // specific value, then R2489 should be covered.
                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: isSoapFault[{0}],error[{1}] for requirement #R2489",
                        isSoapFault,
                        string.IsNullOrEmpty(error) ? "NullOrEmpty" : error);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2489,
                    2489,
                    @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<64> Section 3.1.4.24: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
            }

            this.Site.Assert.IsTrue(isSoapFault, "Call GetListItems operation with invalid ListName must generate a Soap Fault");
        }

        /// <summary>
        /// This test case is used to test GetListItems operation when Paging element is not empty.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC55_GetListItems_WithPaging()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Insert 10 item.
            List<string> addedItems = TestSuiteHelper.AddListItems(listId, 10);

            // Sorted the list to make sure order by Ascending.
            addedItems = addedItems.OrderBy<string, int>(key => Convert.ToInt32(key)).ToList();

            // Create Query element to make the z:row ordered by OWS_ID Ascending.
            GetListItemsQuery query = new GetListItemsQuery();
            query.Query = new CamlQueryRoot();
            query.Query.OrderBy = new OrderByDefinition();
            query.Query.OrderBy.FieldRef = new FieldRefDefinitionOrderBy[1];
            query.Query.OrderBy.FieldRef[0] = new FieldRefDefinitionOrderBy();
            query.Query.OrderBy.FieldRef[0].Name = AdapterHelper.FieldIDName;
            query.Query.OrderBy.FieldRef[0].Ascending = "TRUE";

            // Set the row limit.
            string rowLimit = "2";

            // Call the GetListItems excluding the paging.
            GetListItemsResponseGetListItemsResult result = null;
            result = this.listswsAdapter.GetListItems(
                                    listId,
                                    null,
                                    query,
                                    null,
                                    rowLimit,
                                    null,
                                    null);

            // Call GetListItems with ListItemCollectionPositionNext which is got by previous call.
            CamlQueryOptions options = new CamlQueryOptions();
            options.QueryOptions = new CamlQueryOptionsQueryOptions();
            options.QueryOptions.Paging = new CamlQueryOptionsQueryOptionsPaging();
            options.QueryOptions.Paging.ListItemCollectionPositionNext = result.listitems.data.ListItemCollectionPositionNext;
            result = this.listswsAdapter.GetListItems(
                                    listId,
                                    null,
                                    query,
                                    null,
                                    null,
                                    options,
                                    null);

            string columnNameId = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            DataTable data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // If the returned first z:row ows_ID equals the next page's first addedItems element, 
            // Then capture requirement R1252.
            bool isNextPage = Convert.ToString(data.Rows[0][columnNameId]) == addedItems[Convert.ToInt32(rowLimit)];
            Site.CaptureRequirementIfIsTrue(
                isNextPage,
                1252,
                @"[Paging]The protocol server MUST return the next page of data based on the "
                + "bookmark supplied.");
        }

        /// <summary>
        /// This test case is used to test GetListItems in case that the Paging element is empty or excluded.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC56_GetListItems_WithoutPaging()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Insert 10 items.
            List<string> addedItems = TestSuiteHelper.AddListItems(listId, 10);

            // Sort the list to make sure order by Ascending.
            addedItems = addedItems.OrderBy<string, int>(key => Convert.ToInt32(key)).ToList();

            // Create Query element to make the z:row ordered by OWS_ID Ascending.
            GetListItemsQuery query = new GetListItemsQuery();
            query.Query = new CamlQueryRoot();
            query.Query.OrderBy = new OrderByDefinition();
            query.Query.OrderBy.FieldRef = new FieldRefDefinitionOrderBy[1];
            query.Query.OrderBy.FieldRef[0] = new FieldRefDefinitionOrderBy();
            query.Query.OrderBy.FieldRef[0].Name = AdapterHelper.FieldIDName;
            query.Query.OrderBy.FieldRef[0].Ascending = "TRUE";

            // Call the GetListItems excluding the paging.
            GetListItemsResponseGetListItemsResult result = null;
            result = this.listswsAdapter.GetListItems(
                                    listId,
                                    null,
                                    query,
                                    null,
                                    null,
                                    null,
                                    null);

            string columnNameId = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            DataTable data = AdapterHelper.ExtractData(result.listitems.data.Any);

            // If the returned z:row count number equals to the all item number
            // and the returned first z:row ows_ID equals the first addedItems element, 
            // it will be sure that it starts from the first index.
            bool isStartFirstIndexExcludePaging = data.Rows.Count == addedItems.Count && Convert.ToString(data.Rows[0][columnNameId]) == addedItems[0];
            Site.Assert.IsTrue(
                            isStartFirstIndexExcludePaging,
                            "When paging element is not specified, the list items returned start from the beginning of the list.");

            // Call GetListItems with empty paging
            CamlQueryOptions options = new CamlQueryOptions();
            options.QueryOptions = new CamlQueryOptionsQueryOptions();
            options.QueryOptions.Paging = new CamlQueryOptionsQueryOptionsPaging();
            result = this.listswsAdapter.GetListItems(
                                    listId,
                                    null,
                                    query,
                                    null,
                                    null,
                                    options,
                                    null);

            // If the returned z:row count number equals to the all item number
            // and the returned first z:row ows_ID equals the first addedItems element, 
            // it will be sure that it starts from the first index.
            data = AdapterHelper.ExtractData(result.listitems.data.Any);
            bool isStartFirstIndexEmptyPaging = data.Rows.Count == addedItems.Count && Convert.ToString(data.Rows[0][columnNameId]) == addedItems[0];
            Site.Assert.IsTrue(
                            isStartFirstIndexEmptyPaging,
                            "When paging element is empty, the list items returned start from the beginning of the list.");

            Site.CaptureRequirementIfIsTrue(
                isStartFirstIndexEmptyPaging && isStartFirstIndexExcludePaging,
                1251,
                @"[Paging]If this element[Paging] is left empty or excluded, the list items returned start from the beginning of the list.");
        }

        #endregion

        #region GetVersionCollection

        /// <summary>
        ///  This test case is used to test GetVersionCollection operation when the client excludes one or more parameters.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC57_GetVersionCollection_ExcludesParameter()
        {
            #region AddList
            string listId = TestSuiteHelper.CreateList();
            #endregion

            // Add one item into the list
            List<string> listIds = TestSuiteHelper.AddListItems(listId, 1);
            string addedListItemId = listIds[0];

            #region GetVersionCollection

            // Call GetVersionCollection.
            bool isSoapFaultExisted = false;
            string errorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetVersionCollection(
                       listId,
                       addedListItemId,
                       null);
            }
            catch (SoapException sopaEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(sopaEx);
            }

            this.Site.Assert.IsTrue(isSoapFaultExisted, "There is no Soap Fault generated for requirement#769,#1946,#777");
            #endregion

            #region CaptureRequirement 769,1946,777
            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(errorCode),
                769,
               @"[In GetVersionCollection operation] If the protocol client excludes one or more "
                 + "parameters, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(errorCode),
                1946,
               @"[In GetVersionCollection operation] [If the protocol client excludes one or more parameters, "
                 + "or if the strListID is null, the protocol server MUST return a SOAP fault.]There is no error code for this fault.");

            Site.CaptureRequirementIfIsTrue(
               string.IsNullOrEmpty(errorCode),
                777,
               @"[In GetVersionCollection operation] otherwise [if the operation fails], the protocol "
                            + "server MUST return a SOAP fault.");
            #endregion
        }

        /// <summary>
        ///  This test case is used to test GetVersionCollection operation when strFieldName does not correspond to any field in the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC58_GetVersionCollection_FieldNameNotCorrespond()
        {
            #region AddList
            string listId = TestSuiteHelper.CreateList();
            #endregion

            // Add one item into the list.
            List<string> listIds = TestSuiteHelper.AddListItems(listId, 1);
            string addedListItemId = listIds[0];

            // Call GetVersionCollection.
            bool isSoapFaultExisted = false;
            string errorCode = string.Empty;
            try
            {
                // This field name will not be included in Current List.
                string notExistedFileName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();
                this.listswsAdapter.GetVersionCollection(
                       listId,
                       addedListItemId,
                       notExistedFileName);
            }
            catch (SoapException sopaEx)
            {
                isSoapFaultExisted = true;
                errorCode = TestSuiteHelper.GetErrorCode(sopaEx);
            }

            this.Site.Assert.IsTrue(isSoapFaultExisted, "There is no Soap Fault generated for requirement#768,#1945,#777");

            Site.Log.Add(
                         LogEntryKind.Debug,
                        "The actual value: errorCode[{0}] for requirement #R768 #r1945 #R777",
                        string.IsNullOrEmpty(errorCode) ? "NullOrempty" : errorCode);

            #region CaptureRequirement 768,1945,777
            Site.CaptureRequirementIfIsTrue(
                string.IsNullOrEmpty(errorCode),
                768,
                        @"[In GetVersionCollection operation] If the strFieldName does not correspond "
                        + "to a field in a specified list item of a specified list on the site, the protocol "
                        + "server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
               string.IsNullOrEmpty(errorCode),
               1945,
                        @"[In GetVersionCollection operation] [If the strFieldName does not correspond "
                        + "to a field in a specified list item of a specified list on the site, the protocol "
                        + "server MUST return a SOAP fault.]There is no error code for this fault.");

            Site.CaptureRequirementIfIsTrue(
               string.IsNullOrEmpty(errorCode),
                777,
               @"[In GetVersionCollection operation] otherwise [if the operation fails], the protocol "
                            + "server MUST return a SOAP fault.");

            #endregion
        }

        /// <summary>
        /// This test case is used to verify GetVersionCollection operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC59_GetVersionCollection_InvalidGUIDAndNotCorrespond_SP3WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2492, this.Site), @"Test is executed only when R2492Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Initialize a valid ListName structure.
            string listname = TestSuiteHelper.CreateList();

            List<string> listString = TestSuiteHelper.AddListItems(listname, 1);

            // Read value from configuration file.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            bool isSoapFaultGenerated = false;
            string errorString = string.Empty;
            try
            {
                // Call GetVersionCollection method to get the version.
                this.listswsAdapter.GetVersionCollection(invalidListName, listString[0], validFieldName);
            }
            catch (SoapException soapException)
            {
                // Verify MS-LISTSWS requirement: MS-LISTSWS_R2492
                isSoapFaultGenerated = true;
                errorString = TestSuiteHelper.GetErrorString(soapException);
            }

            // If the value of the error is null and the detailed message of SoapException is equal to the 
            // specific value, then R2492 should be covered.
            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: isSoapFaultGenerated[{0}],errorString[{1}] for requirement #R2492",
                isSoapFaultGenerated,
                string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultGenerated
                && "Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).".Equals(errorString, StringComparison.OrdinalIgnoreCase),
                2492,
                @"[In Appendix B: Product Behavior] Implementation does not return a SOAP fault with error code 0x82000006. (<67> Section 3.1.4.25: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
        }

        /// <summary>
        /// This test case is used to test GetVersionCollection operation when the list ID is not a valid GUID and also not a valid title.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC60_GetVersionCollection_InvalidListGUID()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2765, this.Site), @"Test is executed only when R2765Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // set list item count 1.
            List<string> ids = TestSuiteHelper.AddListItems(listID, 1);

            string errorCode = string.Empty;
            bool isSoapFaultExist = false;

            try
            {
                // Call GetVersionCollection to query the field version
                this.listswsAdapter.GetVersionCollection(
                                    TestSuiteHelper.GetInvalidGuidAndNocorrespondString(),
                                    ids[0],
                                    fieldName);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isSoapFaultExist = true;
            }

            Site.Assert.IsTrue(
                            isSoapFaultExist,
                            "When the list id is not valid or exist title, the server will return soap fault exception.");

            // If return error code equals 0x82000006, then capture R2765.
            Site.CaptureRequirementIfAreEqual<string>(
                        "0x82000006",
                        errorCode,
                        2765,
                        @"[In GetVersionCollection operation] Implementation does return a SOAP fault with error code 0x82000006 , if strListID does not correspond to a list from either of these checks. (SharePoint Foundation 2010 and above follow this behavior.)");
        }

        /// <summary>
        /// This test case is used to test GetVersionCollection operation when the value of element strListItemID is null or an empty string, 
        /// or when the value of element strListItemID does not correspond to a list item in a specified list on the site. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC61_GetVersionCollection_InvalidParameterWithEmptyListItemID()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Add a new "Required" field in the generic list
            string strNewRequiredField = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(
                                                       strList_GUID,
                                                       new List<string> { strNewRequiredField },
                                                       new List<string> { "Counter" },
                                                       true,
                                                       new List<string> { null });
            #endregion

            #region Add a new list item in the generic list with a valid value for the "Required" field
            Dictionary<string, string> fieldNameValuePairs = new Dictionary<string, string>();
            Random randomIntValue = new Random();
            string fieldValue = randomIntValue.Next().ToString();
            fieldNameValuePairs.Add(strNewRequiredField, fieldValue);
            UpdateListItemsUpdates listItemUpdates_1 = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { fieldNameValuePairs });

            // Invoke UpdateListItems operation to make a change.
            this.listswsAdapter.UpdateListItems(strList_GUID, listItemUpdates_1);
            #endregion

            #region Invoke GetVersionCollection operation, set the value of strListItemID to null or empty string.

            bool isSoapFauleExistedForNullListItemID = false;
            string errorCodeForNullListItemID = string.Empty;
            string strErrorCode = string.Empty;
            string strErrorString = string.Empty;
            try
            {
                this.listswsAdapter.GetVersionCollection(strList_GUID, null, strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                isSoapFauleExistedForNullListItemID = true;
                errorCodeForNullListItemID = TestSuiteHelper.GetErrorCode(soapEx);
            }

            bool isSoapFauleExistedForEmptyListItemID = false;
            string errorCodeForEmptyListItemID = string.Empty;
            try
            {
                this.listswsAdapter.GetVersionCollection(strList_GUID, string.Empty, strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                isSoapFauleExistedForEmptyListItemID = true;
                errorCodeForEmptyListItemID = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture Requirements #3020 #3021

            Site.Log.Add(
              LogEntryKind.Debug,
              "The actual value: errorCodeForEmptyListItemID[{0}],errorCodeForNullListItemID[{1}] for requirement #R3020, #R3021",
              string.IsNullOrEmpty(errorCodeForEmptyListItemID) ? "NullOrEmpty" : strErrorCode,
              string.IsNullOrEmpty(errorCodeForNullListItemID) ? "NullOrEmpty" : strErrorCode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFauleExistedForNullListItemID && isSoapFauleExistedForEmptyListItemID
                && "0x82000001".Equals(errorCodeForNullListItemID, StringComparison.OrdinalIgnoreCase)
                && "0x82000001".Equals(errorCodeForEmptyListItemID, StringComparison.OrdinalIgnoreCase),
                3020,
                @"[In GetVersionCollection operation] If the strListItemID is null or an empty string, "
                + "the protocol server MUST return a SOAP fault with error code 0x82000001.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFauleExistedForNullListItemID && isSoapFauleExistedForEmptyListItemID
                && "0x82000001".Equals(errorCodeForNullListItemID, StringComparison.OrdinalIgnoreCase)
                && "0x82000001".Equals(errorCodeForEmptyListItemID, StringComparison.OrdinalIgnoreCase),
                3021,
                @"[In GetVersionCollection operation] [If the strListItemID is null or an empty string, "
                + "the protocol server MUST return a SOAP fault with error code 0x82000001.] "
                + "This indicates that the parameter strListItemID is missing or invalid.");

            #endregion

            #region Invoke GetVersionCollection operation, set strListItemID to a negative value.

            try
            {
                this.listswsAdapter.GetVersionCollection(strList_GUID, "-12", strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
                strErrorString = TestSuiteHelper.GetErrorString(soapEx);
            }
            #endregion

            #region Capture Requirements #7662 #7663

            if (Common.IsRequirementEnabled(7660, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7662");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R7662
                bool isVerifyR7662 = string.Equals(strErrorCode, "0x81020014");

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR7662,
                    7662,
                    @"Implementation does return a SOAP fault with error code 0x81020014 when strListItemID is a negative value.(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.) ");
            }

            if (Common.IsRequirementEnabled(7661, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7663");

                // Verify MS-LISTSWS requirement: MS-LISTSWS_R7663
                bool isVerifyR7663 = string.Equals(strErrorCode, null) && string.Equals(strErrorString, AdapterHelper.OutOfRange);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR7663,
                    7663,
                    @"[In Appendix B: Product Behavior] Implementation does return a SOAP fault with no error code and the error string ""Value does not fall within the expected range."" when strListItemID is a negative value.(Windows® SharePoint® Services 3.0 follows this behavior)");
            }

            #endregion

            #region Invoke GetVersionCollection operation, set strListItemID to zero.
            bool isCaptureR7664 = false;
            bool isCaptureR7665 = false;
            try
            {
                this.listswsAdapter.GetVersionCollection(strList_GUID, "0", strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
                strErrorString = TestSuiteHelper.GetErrorString(soapEx);
                if (strErrorCode == null)
                {
                    isCaptureR7664 = true;
                    isCaptureR7665 = string.Equals(strErrorString, AdapterHelper.OutOfRange);
                }
                else
                {
                    isCaptureR7664 = false;
                    isCaptureR7665 = false;
                    string strCaseFail_Info = string.Format("CaseFail: The actual error code is {0}, but the expected error code is null!", strErrorCode);
                    Site.Log.Add(LogEntryKind.TestFailed, strCaseFail_Info);
                }
            }
            #endregion

            #region Capture Requirements #7664 #7665

            if (Common.IsRequirementEnabled(7660, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7664");

                Site.CaptureRequirementIfIsTrue(
                    isCaptureR7664,
                    7664,
                    @"Implementation does return a SOAP fault with no error code when strListItemID is zero.(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(7661, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7665");

                Site.CaptureRequirementIfIsTrue(
                    isCaptureR7665,
                    7665,
                    @"[In Appendix B: Product Behavior] Implementation does return a SOAP fault with no error code and the error string ""Value does not fall within the expected range."" when strListItemID is zero. (Windows® SharePoint® Services 3.0 follows this behavior)");
            }

            #endregion

            #region Invoke GetVersionCollection operation, set strListItemID to an invalid positive value.

            bool isCaptureR7666 = false;
            bool isCaptureR7667 = false;
            try
            {
                this.listswsAdapter.GetVersionCollection(strList_GUID, "9", strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
                strErrorString = TestSuiteHelper.GetErrorString(soapEx);
                if (strErrorCode == null)
                {
                    isCaptureR7666 = true;
                    isCaptureR7667 = string.Equals(strErrorString, AdapterHelper.OutOfRange);
                }
                else
                {
                    isCaptureR7666 = false;
                    isCaptureR7667 = false;
                    string strCaseFail_Info = string.Format("CaseFail: The actual error code is {0}, but the expected error code is null!", strErrorCode);
                    Site.Log.Add(LogEntryKind.TestFailed, strCaseFail_Info);
                }
            }

            #endregion

            #region Capture Requirements #7666 #7667

            if (Common.IsRequirementEnabled(7660, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7666");

                Site.CaptureRequirementIfIsTrue(
                    isCaptureR7666,
                    7666,
                    @"Implementation does return a SOAP fault with no error code when strListItemID is a positive value that does not correspond to a list item in a specified list on the site.(Microsoft® SharePoint® Foundation 2010 and above follow this behavior.) ");
            }

            if (Common.IsRequirementEnabled(7661, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-LISTSWS_R7667");

                Site.CaptureRequirementIfIsTrue(
                    isCaptureR7667,
                    7667,
                    @"[In Appendix B: Product Behavior] Implementation does return a SOAP fault with no error code and the error string ""Value does not fall within the expected range."" when strListItemID is a positive value that does not correspond to a list item in a specified list on the site. (Windows® SharePoint® Services 3.0 follows this behavior)");
            }

            #endregion
        }

        /// <summary>
        /// This test case is used to test the negative status of GetVersionCollection operation when the value of element strListID is null.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC62_GetVersionCollection_InvalidParameterWithNullListName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Add a new "Required" field in the generic list
            string strNewRequiredField = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(
                                                       strList_GUID,
                                                       new List<string> { strNewRequiredField },
                                                       new List<string> { "Counter" },
                                                       true,
                                                       new List<string> { null });
            #endregion

            #region Add a new list item in the generic list with a valid value for the "Required" field
            Dictionary<string, string> fieldNameValuePairs = new Dictionary<string, string>();
            Random randomIntValue = new Random();
            string fieldValue = randomIntValue.Next().ToString();
            fieldNameValuePairs.Add(strNewRequiredField, fieldValue);
            UpdateListItemsUpdates listItemUpdates_1 = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { fieldNameValuePairs });
            UpdateListItemsResponseUpdateListItemsResult updateListItemsResult = null;
            updateListItemsResult = this.listswsAdapter.UpdateListItems(strList_GUID, listItemUpdates_1);
            #endregion

            #region Get the list item ID in the response of UpdateListItems
            string strListItemID = null;
            Site.Assert.IsNotNull(updateListItemsResult, "The Object \"updateListItemsResult\" should not be null! ");
            Site.Assert.IsNotNull(updateListItemsResult.Results, "The Object \"updateListItemsResult.Results\" should not be null! ");
            Site.Assert.IsTrue(updateListItemsResult.Results.Length >= 1, "There is no result in the result array!");
            UpdateListItemsResponseUpdateListItemsResultResult[] arrRresults = updateListItemsResult.Results;
            Site.Assert.IsNotNull(arrRresults[0].Any, "The Object \"arrRresults[0].Any\" should not be null! ");
            XmlElement[] itemRows = arrRresults[0].Any;
            System.Data.DataTable dataTable = AdapterHelper.ExtractData(itemRows);
            Site.Assert.IsNotNull(dataTable, "The Object \"dataTable\" should not be null! ");
            Site.Assert.IsNotNull(dataTable.Rows, "The Object \"dataTable.Rows\" should not be null! ");
            Site.Assert.IsTrue(dataTable.Rows.Count >= 1, "There is no list item in the result!");
            strListItemID = dataTable.Rows[0]["ows_ID"].ToString();
            Site.Assert.IsNotNull(strListItemID, "Failed to get the list item ID in the response of UpdateListItems!");
            #endregion

            #region Invoke GetVersionCollection operation set the value of element strListID to null.

            bool isSoapFaultExist = false;
            string strErrorCode = string.Empty;
            try
            {
                this.listswsAdapter.GetVersionCollection(null, strListItemID, strNewRequiredField);
            }
            catch (SoapException soapEx)
            {
                isSoapFaultExist = true;
                strErrorCode = TestSuiteHelper.GetErrorCode(soapEx);
            }
            #endregion

            #region Capture Requirements #3022 #777

            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: isSoapFaultExist[{0}], ErrorCode[{1}] for requirement #R3022, #R777",
                     isSoapFaultExist,
                     string.IsNullOrEmpty(strErrorCode) ? "NullOrEmpty" : strErrorCode);

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultExist,
                3022,
                @"[In GetVersionCollection operation] If the strListID is null, the protocol server MUST return a SOAP fault.");

            Site.CaptureRequirementIfIsTrue(
                isSoapFaultExist,
                777,
                @"[In GetVersionCollection operation] otherwise [if the operation fails], the protocol server MUST return "
                + "a SOAP fault.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the GetVersionCollection when the query field name is specified. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC63_GetVersionCollection_SpecifiedFieldName()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);
            string fieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Add an item with random field value.
            Dictionary<string, string> item = new Dictionary<string, string>();
            item.Add(fieldName, fieldValue);

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { item });

            UpdateListItemsResponseUpdateListItemsResult addItemResult = this.listswsAdapter.UpdateListItems(listID, updates);

            Site.Assert.IsNotNull(addItemResult, "The result of UpdateListItems operation should not be NULL");

            // Get the new item ID
            string columnNameId = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            string itemID = string.Empty;

            DataTable data = AdapterHelper.ExtractData(addItemResult.Results[0].Any);
            if (data != null && data.Rows.Count == 1)
            {
                itemID = Convert.ToString(data.Rows[0][columnNameId]);
            }

            // Call GetVersionCollection to query the field version
            GetVersionCollectionResponseGetVersionCollectionResult result = null;
            result = this.listswsAdapter.GetVersionCollection(
                                listID,
                                itemID,
                                fieldName);

            // If all the versions contain the expect field name and the Value is the value of the field, capture R1959 and R2232
            bool isFieldNameExistAndValueCorrect = true;
            for (int i = 0; i < result.Versions.Length; i++)
            {
                isFieldNameExistAndValueCorrect = isFieldNameExistAndValueCorrect
                                    && result.Versions[i].AnyAttr.FirstOrDefault(attr => attr.Name == fieldName && attr.Value == fieldValue) != null;
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R1959
            Site.CaptureRequirementIfIsTrue(
                isFieldNameExistAndValueCorrect,
                1959,
                @"[GetVersionCollectionResponse]Version.FieldName:  The attribute Name MUST be the name of the field (2) for which the version has been queried.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2232
            Site.CaptureRequirementIfIsTrue(
                isFieldNameExistAndValueCorrect,
                2232,
                @"[GetVersionCollectionResponse][Version.FieldName] the Value MUST be the value of the field (2) for which the version has been queried.");
        }

        /// <summary>
        /// This test case is used to test GetVersionCollection operation when the value of element strListID is not a valid GUID but its value corresponds to the title of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC64_GetVersionCollection_SucceedWithNotGUIDListName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Invoke GetList operation to get the title of the new generic list.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(strList_GUID);
            Site.Assert.IsNotNull(listDef, "The object \"listDef\" should not be null !");
            Site.Assert.IsNotNull(listDef.Title, "The object \"listDef.Title\" should not be null !");
            string strList_Title = listDef.Title;
            #endregion

            #region Add a new "Required" field in the generic list
            string strNewRequiredField = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(
                strList_GUID,
                new List<string> { strNewRequiredField },
                new List<string> { "Counter" },
                true,
                new List<string> { null });
            #endregion

            #region Add a new list item in the generic list with a valid value for the "Required" field
            Dictionary<string, string> fieldNameValuePairs = new Dictionary<string, string>();
            Random randomIntValue = new Random();
            string fieldValue = randomIntValue.Next().ToString();
            fieldNameValuePairs.Add(strNewRequiredField, fieldValue);
            UpdateListItemsUpdates listItemUpdates_1 = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { fieldNameValuePairs });

            UpdateListItemsResponseUpdateListItemsResult updateListItemsResult = null;
            updateListItemsResult = this.listswsAdapter.UpdateListItems(strList_GUID, listItemUpdates_1);
            #endregion

            #region Get the list item ID in the response of UpdateListItems
            string strListItemID = null;
            Site.Assert.IsNotNull(updateListItemsResult, "The Object \"updateListItemsResult\" should not be null! ");
            Site.Assert.IsNotNull(updateListItemsResult.Results, "The Object \"updateListItemsResult.Results\" should not be null! ");
            Site.Assert.IsTrue(updateListItemsResult.Results.Length >= 1, "There is no result in the result array!");
            UpdateListItemsResponseUpdateListItemsResultResult[] arrRresults = updateListItemsResult.Results;
            Site.Assert.IsNotNull(arrRresults[0].Any, "The Object \"arrRresults[0].Any\" should not be null! ");
            XmlElement[] itemRows = arrRresults[0].Any;
            System.Data.DataTable dataTable = AdapterHelper.ExtractData(itemRows);
            Site.Assert.IsNotNull(dataTable, "The Object \"dataTable\" should not be null! ");
            Site.Assert.IsNotNull(dataTable.Rows, "The Object \"dataTable.Rows\" should not be null! ");
            Site.Assert.IsTrue(dataTable.Rows.Count >= 1, "There is no list item in the result!");
            strListItemID = dataTable.Rows[0]["ows_ID"].ToString();
            Site.Assert.IsNotNull(strListItemID, "Failed to get the list item ID in the response of UpdateListItems!");
            #endregion

            #region Invoke GetVersionCollection operation use the list title as the value of element listName.
            GetVersionCollectionResponseGetVersionCollectionResult getVersionCollectionResult = null;
            getVersionCollectionResult = this.listswsAdapter.GetVersionCollection(strList_Title, strListItemID, strNewRequiredField);
            #endregion

            #region Capture Requirements #7571 #763 #764 #771
            // Verify MS-LISTSWS requirement: MS-LISTSWS_R7571.
            // If getting version collection is successful, it means implementation does support this GetListItemChangesSinceToken method. R7571 can be captured.
            if (Common.IsRequirementEnabled(7571, this.Site))
            {
                Site.CaptureRequirementIfIsNotNull(
                    getVersionCollectionResult,
                    7571,
                    @"Implementation does support this method[GetVersionCollection]. (Windows SharePoint Services 3.0 and above follow this behavior.)");
            }

            // If the successful response of GetVersionCollection is returned, capture requirements #763 and #764.
            Site.CaptureRequirementIfIsNotNull(
                getVersionCollectionResult,
                763,
                @"[In GetVersionCollection operation] If the specified strListID is not a valid GUID, check if the strListID "
                + "corresponds to the list title of a list on the site and if so, use that list.");

            Site.CaptureRequirementIfIsNotNull(
              getVersionCollectionResult,
                764,
                @"[In GetVersionCollection operation] If the specified strListID does not correspond to the identification of "
                + "a list on the site, check if the strListID corresponds to the list title of a list on the site and if so, "
                + "use that list.");

            // If protocol SUT returns version collections element in the successful response of GetVersionCollection, 
            // capture requirements #771.
            Site.Assert.IsNotNull(getVersionCollectionResult, "The Object \"getVersionCollectionResult\" should not be null! ");
            Site.Assert.IsNotNull(getVersionCollectionResult.Versions, "The Object \"getVersionCollectionResult.Versions\" should not be null! ");
            bool iscaptureR771 = false;
            if (getVersionCollectionResult.Versions.Length > 0)
            {
                iscaptureR771 = true;
            }

            Site.Log.Add(
                         LogEntryKind.Debug,
                         "The actual value: getVersionCollectionResult.Versions.Length[{0}] for requirement #R771",
                          getVersionCollectionResult.Versions.Length);

            Site.CaptureRequirementIfIsTrue(
                iscaptureR771,
                771,
                @"[In GetVersionCollection operation] If the protocol client passes valid input parameters, the protocol "
                + "server MUST return the version collection.");

            #endregion
        }

        /// <summary>
        /// Test the GetVersionCollection operation when the value of 
        /// element strListID is the valid GUID of the list.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC65_GetVersionCollection_SucceedWithValidGUIDListName()
        {
            #region Invoke AddList operation to create a new generic list.
            string strList_GUID = TestSuiteHelper.CreateList((int)TemplateType.Generic_List);
            #endregion

            #region Add a new "Required" field in the generic list
            string strNewRequiredField = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(
                                                       strList_GUID,
                                                       new List<string> { strNewRequiredField },
                                                       new List<string> { "Counter" },
                                                       true,
                                                       new List<string> { null });
            #endregion

            #region Add a new list item in the generic list with a valid value for the "Required" field
            Dictionary<string, string> fieldNameValuePairs = new Dictionary<string, string>();
            Random randomIntValue = new Random();
            string fieldValue = randomIntValue.Next().ToString();
            fieldNameValuePairs.Add(strNewRequiredField, fieldValue);
            UpdateListItemsUpdates listItemUpdates_1 = TestSuiteHelper.CreateUpdateListItems(
                                                        new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                        new List<Dictionary<string, string>> { fieldNameValuePairs });

            UpdateListItemsResponseUpdateListItemsResult updateListItemsResult = null;
            updateListItemsResult = this.listswsAdapter.UpdateListItems(strList_GUID, listItemUpdates_1);
            #endregion

            #region Get the list item ID in the response of UpdateListItems
            string strListItemID = null;
            Site.Assert.IsNotNull(updateListItemsResult, "The Object \"updateListItemsResult\" is null! ");
            Site.Assert.IsNotNull(updateListItemsResult.Results, "The Object \"updateListItemsResult.Results\" is null! ");
            Site.Assert.IsTrue(updateListItemsResult.Results.Length >= 1, "There is no result in the result array!");
            UpdateListItemsResponseUpdateListItemsResultResult[] arrRresults = updateListItemsResult.Results;
            Site.Assert.IsNotNull(arrRresults[0].Any, "The Object \"arrRresults[0].Any\" is null! ");
            XmlElement[] rowItems = arrRresults[0].Any;
            System.Data.DataTable dataTable = AdapterHelper.ExtractData(rowItems);
            Site.Assert.IsNotNull(dataTable, "The Object \"dataTable\" is null! ");
            Site.Assert.IsNotNull(dataTable.Rows, "The Object \"dataTable.Rows\" is null! ");
            Site.Assert.IsTrue(dataTable.Rows.Count >= 1, "There is no list item in the result!");
            strListItemID = dataTable.Rows[0]["ows_ID"].ToString();
            Site.Assert.IsNotNull(strListItemID, "Failed to get the list item ID in the response of UpdateListItems!");
            #endregion

            #region Invoke GetVersionCollection operation use the list title as the value of element listName.
            GetVersionCollectionResponseGetVersionCollectionResult getVersionCollectionResult = null;
            getVersionCollectionResult = this.listswsAdapter.GetVersionCollection(strList_GUID, strListItemID, strNewRequiredField);
            #endregion

            #region Capture Requirements #762 #771

            // If protocol SUT returns the successful response of GetVersionCollection, capture requirements #762.
            Site.Log.Add(
                         LogEntryKind.Debug,
                         "The actual value: getVersionResult[{0}] for requirement #R762",
                         null == getVersionCollectionResult ? "null" : "Not null");

            Site.CaptureRequirementIfIsNotNull(
                getVersionCollectionResult,
                762,
                @"[In GetVersionCollection operation] If the specified strListID is a valid GUID and corresponds to "
                + "the identification of a list on the site, use that list.");

            // If protocol SUT returns version collections element in the successful response of GetVersionCollection, 
            // capture requirements #771.
            bool isCaptureR771 = false;
            Site.Assert.IsNotNull(getVersionCollectionResult, "The Object \"getVersionCollectionResult\" is null! ");
            Site.Assert.IsNotNull(getVersionCollectionResult.Versions, "The Object \"getVersionCollectionResult.Versions\" is null! ");
            if (getVersionCollectionResult.Versions.Length > 0)
            {
                isCaptureR771 = true;
            }

            Site.Log.Add(
                          LogEntryKind.Debug,
                          "The actual value: getVersionCollectionResult.Versions.Length[{0}] for requirement #R771",
                           getVersionCollectionResult.Versions.Length);

            Site.CaptureRequirementIfIsTrue(
                isCaptureR771,
                771,
                @"[In GetVersionCollection operation] If the protocol client passes valid input parameters, the protocol "
                + "server MUST return the version collection.");

            #endregion
        }

        #endregion

        #region IncludeMandatoryColumns

        /// <summary>
        /// This test case is used to confirm the following behavior of element "IncludeMandatoryColumns" in the complex type "CamlQueryOptions". 
        /// The element "IncludeMandatoryColumns" specifies that required fields and fields used by specified calculated fields be returned in addition to the fields specified by the viewFields parameter, 
        /// if the element "IncludeMandatoryColumns" is set to TRUE. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC66_IncludeMandatoryColumns()
        {
            #region Create a new generic list
            int templateId = (int)TemplateType.Generic_List;
            string listName = TestSuiteHelper.CreateList(templateId);
            #endregion

            #region Add a new "Required" field in the generic list
            string newRequiredField = TestSuiteHelper.GetUniqueFieldName();
            TestSuiteHelper.AddFieldsToList(
                                    listName,
                                    new List<string> { newRequiredField },
                                    new List<string> { "Counter" },
                                    true,
                                    new List<string> { null });
            #endregion

            #region Add a new list item in the generic list with a valid value for the "Required" field
            UpdateListItemsUpdates listItemUpdates_1 = new UpdateListItemsUpdates();
            listItemUpdates_1.Batch = new UpdateListItemsUpdatesBatch();
            listItemUpdates_1.Batch.Method = new UpdateListItemsUpdatesBatchMethod[1];
            listItemUpdates_1.Batch.Method[0] = new UpdateListItemsUpdatesBatchMethod();
            listItemUpdates_1.Batch.Method[0].ID = (uint)0;
            listItemUpdates_1.Batch.Method[0].Cmd = MethodCmdEnum.New;
            listItemUpdates_1.Batch.Method[0].Field = new UpdateListItemsUpdatesBatchMethodField[1];
            listItemUpdates_1.Batch.Method[0].Field[0] = new UpdateListItemsUpdatesBatchMethodField();
            listItemUpdates_1.Batch.Method[0].Field[0].Name = newRequiredField;
            listItemUpdates_1.Batch.Method[0].Field[0].Value = "30";
            this.listswsAdapter.UpdateListItems(listName, listItemUpdates_1);
            #endregion

            #region Invoke the operation "GetListItems" to check the "Required" field.
            GetListItemsResponseGetListItemsResult getListItemsResult = null;
            CamlViewFields viewFields = new CamlViewFields();
            viewFields.ViewFields = new CamlViewFieldsViewFields();
            viewFields.ViewFields.FieldRef = new CamlViewFieldsViewFieldsFieldRef[1];
            viewFields.ViewFields.FieldRef[0] = new CamlViewFieldsViewFieldsFieldRef();
            viewFields.ViewFields.FieldRef[0].Name = "ID";
            CamlQueryOptions queryOption = new CamlQueryOptions();
            queryOption.QueryOptions = new CamlQueryOptionsQueryOptions();
            XmlNode[] row = null;
            System.Collections.IEnumerator attributeEnumerator = null;
            string attribute_NewRequiredField = "ows_" + newRequiredField;

            // Invoke "GetListItems" operation with the value of request element "IncludeMandatoryColumns" is "TRUE".
            queryOption.QueryOptions.IncludeMandatoryColumns = "TRUE";
            getListItemsResult = this.listswsAdapter.GetListItems(listName, null, null, viewFields, null, queryOption, null);
            Site.Assert.IsNotNull(getListItemsResult.listitems, "The \"listitems\" is null in the first invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data, "The \"listitems.data\" is null in the first invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data.Any, "The \"listitems.data.Any\" is null in the first invoke of \"GetListItems\"!");
            row = getListItemsResult.listitems.data.Any;
            Site.Assert.IsNotNull(row[0], "The \"row[0]\" is null in the first invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(row[0].Attributes, "The \"row[0].Attributes\" is null in the first invoke of \"GetListItems\"!");
            attributeEnumerator = row[0].Attributes.GetEnumerator();

            // Check if the "Required" field appears as one attribute of "z:row" element in the response of "GetListItems".
            bool findRequiredField_IncludeMandatoryColumns = false;
            while (attributeEnumerator.MoveNext())
            {
                XmlAttribute curAttribute = (XmlAttribute)attributeEnumerator.Current;
                if (string.Compare(curAttribute.Name, attribute_NewRequiredField, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    findRequiredField_IncludeMandatoryColumns = true;
                    break;
                }
            }

            // Invoke "GetListItems" operation with the value of request element "IncludeMandatoryColumns" is "FALSE".
            queryOption.QueryOptions.IncludeMandatoryColumns = "FALSE";
            getListItemsResult = this.listswsAdapter.GetListItems(listName, null, null, viewFields, null, queryOption, null);
            Site.Assert.IsNotNull(getListItemsResult.listitems, "The \"listitems\" is null in the second invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data, "The \"listitems.data\" is null in the second invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(getListItemsResult.listitems.data.Any, "The \"listitems.data.Any\" is null in the second invoke of \"GetListItems\"!");
            row = getListItemsResult.listitems.data.Any;
            Site.Assert.IsNotNull(row[0], "The \"row[0]\" is null in the second invoke of \"GetListItems\"!");
            Site.Assert.IsNotNull(row[0].Attributes, "The \"row[0].Attributes\" is null in the second invoke of \"GetListItems\"!");
            attributeEnumerator = row[0].Attributes.GetEnumerator();

            // Check if the "Required" field does not appear as one attribute of "z:row" element in the response of "GetListItems".
            bool findRequiredField_NotIncludeMandatoryColumns = false;
            while (attributeEnumerator.MoveNext())
            {
                XmlAttribute curAttribute = (XmlAttribute)attributeEnumerator.Current;
                if (string.Compare(curAttribute.Name, attribute_NewRequiredField, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    findRequiredField_NotIncludeMandatoryColumns = true;
                    break;
                }
            }
            #endregion

            #region Try to confirm the requirement #1256.
            // Capture the requirement #1256 when following 2 conditions are all satisfied:
            //  1. The "Required" field appears as one attribute of "z:row" element in the response of "GetListItems",
            //      when the value of request element "IncludeMandatoryColumns" is "TRUE";
            //  2. The "Required" field does not appear as one attribute of "z:row" element in the response of "GetListItems",
            //      when the value of request element "IncludeMandatoryColumns" is "FALSE";
            bool confirmR1256 = false;
            if (findRequiredField_IncludeMandatoryColumns && !findRequiredField_NotIncludeMandatoryColumns)
            {
                confirmR1256 = true;
            }

            Site.CaptureRequirementIfIsTrue(
                confirmR1256,
                1256,
                "[IncludeMandatoryColumns]Specifies that required fields and fields used by specified calculated fields be returned in addition to the fields specified by the viewFields parameter if set to True.");

            #endregion
        }

        #endregion

        #region UpdateListItemWithKnowledge

        /// <summary>
        /// This test case is used to test UpdateListItemWithKnowledge operation when inserting one item using non-existing field name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC67_UpdateListItemWithKnowledge_FieldNotExist()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Call UpdateListItem operation to add one item using non-exists field name
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item = new Dictionary<string, string>();
            item.Add(fieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item);
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items);
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                listName,
                                                updates,
                                                null,
                                                null);

            this.Site.Assert.AreEqual<int>(
                                1,
                                result.Results.Length,
                                "If only insert one item, then there will be one result element.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2356
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                2356,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] [In Batch element] [In Method element] [In Field element]  The protocol server MUST return an error indicating why the update failed, if the field is not a special field.");

            this.Site.CaptureRequirementIfIsTrue(
                result.Results[0].ErrorCode.IndexOf("0x", StringComparison.OrdinalIgnoreCase) == 0,
                2324,
                @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse][In UpdateListItemsWithKnowledgeResult element] Otherwise, [If an operation does not complete successfully] the ErrorCode MUST be set to a hexadecimal representation of the error encountered.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2367
            Site.CaptureRequirementIfIsNotNull(
                AdapterHelper.GetElementValue(result.Results[0].Any, "ErrorText"),
                2367,
                @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse][In UpdateListItemsWithKnowledgeResult element] Otherwise, [If an operation does not complete successfully] the ErrorText element MUST have a description of the error.");
        }

        /// <summary>
        /// The test case is used to verify UpdateListItemWithKnowledge operation when TRUE and ListVersion is specified with a number.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC68_UpdateListItemWithKnowledge_InvalidListVersionWithSchemaLock()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listID = TestSuiteHelper.CreateList();

            // Call UpdateListItems with lockSchema equals true and version is not the current version
            Dictionary<string, string> listItem = new Dictionary<string, string>();
            listItem.Add(
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site),
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(
                                                                 new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                                 new List<Dictionary<string, string>> { listItem });

            // Call GetListItemChangesSinceToken with null change token to get current list version
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult sinceTokenResult = null;
            sinceTokenResult = this.listswsAdapter.GetListItemChangesSinceToken(
                                             listID,
                                             null,
                                             null,
                                             null,
                                             null,
                                             null,
                                             null,
                                             null);

            // Find the list version number.
            uint currentVersion = (uint)sinceTokenResult.listitems.Changes.List.Version;

            // Set LockSchema is TRUE
            updates.Batch.LockSchema = "TRUE";
            updates.Batch.ListVersionSpecified = true;

            // Plus the current version with 1, then the ListVersion is not valid.
            updates.Batch.ListVersion = currentVersion + 1;
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult updateItemsResult = null;
            updateItemsResult = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                    listID,
                                                    updates,
                                                    null,
                                                    null);

            // If set LockSchema is TRUE and ListVersion is specified with an error number, the SUT will return a an error code.
            Site.Assert.AreEqual<int>(
                                1,
                                updateItemsResult.Results.Length,
                                "When only insert one item, the UpdateListItems return only one result");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2348
            Site.CaptureRequirementIfAreEqual<string>(
                "0x80070666",
                updateItemsResult.Results[0].ErrorCode,
                2348,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] [In Batch element] [LockSchema attribute]  If TRUE and ListVersion is specified with a number, the protocol server MUST return a SOAP fault with error code 0x80070666 if the schema version passed in by the protocol client does not match the list schema version on the protocol server.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItemWithKnowledge operation in the case that the value of OnError attribute is OnErrorReturn.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC69_UpdateListItemWithKnowledge_OnErrorReturn()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Get the configuration value
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct three items, the second one is invalid, the first and the last one are valid
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item which is valid.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item which is invalid.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item which is valid.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                        listName,
                                        updates,
                                        null,
                                        null);

            this.Site.Assert.AreEqual<int>(
                                    3,
                                    result.Results.Length,
                                    "When insert three items, there are must be three result elements.");

            this.Site.Assert.AreEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                "When insert the first item, the return result error code is not equal 0x800704c7.");

            this.Site.Assert.AreNotEqual<string>(
                "0x00000000",
                result.Results[1].ErrorCode,
                "When insert the second item, the return result error code is not equal 0x800704c7.");

            // If third one contains "0x800704c7" then verify R2314.
            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2314
            Site.CaptureRequirementIfAreEqual<string>(
                "0x800704c7",
                result.Results[2].ErrorCode,
                2314,
                @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse][In UpdateListItemsWithKnowledgeResult element] [If the OnError attribute of the Batch element is set to Return] If a Method operation fails, then all subsequent Method operation in the Batch MUST fail with an error of 0x800704c7.");
        }

        /// <summary>
        /// The test case will verify UpdateListItemWithKnowledge operation when owsHiddenversion conflicts.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC70_UpdateListItemWithKnowledge_OwsHiddenVersionConflict()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            string listID = TestSuiteHelper.CreateList();

            // set list item count 1.
            List<string> ids = TestSuiteHelper.AddListItems(listID, 1);

            // Get the owShiddenVersion
            string owsHiddenVersion = TestSuiteHelper.GetOwsHiddenVersion(listID);

            // Construct one update item with invalid owShiddenVersion.
            List<Dictionary<string, string>> updatedItems = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.Update);
            Dictionary<string, string> item = new Dictionary<string, string>();

            // Set the updated item id
            item.Add(AdapterHelper.FieldIDName, ids[0]);

            // Set the invalid owsVersion whose value is current owshiddenversion plus 1.
            string invalidowsVersion = (int.Parse(owsHiddenVersion) + 1).ToString();
            item.Add(AdapterHelper.FieldOwshiddenversionName, invalidowsVersion);
            item.Add(Common.GetConfigurationPropertyValue("ListFieldText", this.Site), TestSuiteHelper.GenerateRandomString(5));
            updatedItems.Add(item);

            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(
                                                                cmds,
                                                                updatedItems);

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                    listID,
                                                    updates,
                                                    null,
                                                    null);

            Site.Assert.AreEqual<int>(
                                1,
                                result.Results.Length,
                                "After updating one item using UpdateListItems with invalid owshiddenversion, the server MUST return one result");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2320
            bool isR2320Verified = int.Parse(result.Results[0].Version) == int.Parse(invalidowsVersion) + 1;

            Site.CaptureRequirementIfIsTrue(
                isR2320Verified,
                2320,
                @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse][In UpdateListItemsWithKnowledgeResult element] The Version attribute MUST be the owshiddenversion contained in the UpdateListItems request plus 1.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2360
            Site.CaptureRequirementIfAreEqual<string>(
                "0x81020015",
                result.Results[0].ErrorCode,
                2360,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] [In Batch element] [In Method element] [In Field element]  If the owshiddenversion specified by the protocol client is different than the current value of the owshiddenversion field's value for the list item on the protocol server, the protocol server MUST return error code 0x81020015.");

            Site.CaptureRequirementIfAreEqual<string>(
              "0x81020015",
              result.Results[0].ErrorCode,
              2353,
              @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] [In Batch element] [In Method element] [Update attribute] [If additional field references not set those columns to the value specified] an error will be returned.");

            bool isListAttributeExist = result.Results[0].List != null;
            bool isVersionAttributeExist = result.Results[0].Version != null;

            Site.Assert.IsTrue(
                        isListAttributeExist,
                        "When returns error code 0x81020015 in the operation UpdateListItems, the list attribute must be exist.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2318
            bool isVersionAndListAttrExist = isVersionAttributeExist && isListAttributeExist;

            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isVersionAndListAttrExist[{0}] for requirement #R2318",
                    string.IsNullOrEmpty(result.Results[0].Version) ? "NullOrEmpty" : result.Results[0].Version);

            Site.CaptureRequirementIfIsTrue(
                isVersionAndListAttrExist,
                2318,
                @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse][In UpdateListItemsWithKnowledgeResult element]The List and Version attributes MUST be returned if a Method operation fails with an error of 0x81020015.");

            // Construct one update item which contains owshiddenversion equal to the value in the server.
            updatedItems.Clear();
            cmds.Clear();
            item.Clear();

            // Set the updated item id.
            item.Add(AdapterHelper.FieldIDName, ids[0]);

            // Set the valid  owsVersion.
            owsHiddenVersion = TestSuiteHelper.GetOwsHiddenVersion(listID);
            item.Add(AdapterHelper.FieldOwshiddenversionName, owsHiddenVersion);

            // Set update the new value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string newFieldValue = TestSuiteHelper.GenerateRandomString(5);
            item.Add(validFieldName, newFieldValue);
            updatedItems.Add(item);
            cmds.Add(MethodCmdEnum.Update);

            // Call UpdateListItems to insert the item using List GUID.
            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, updatedItems, OnErrorEnum.Continue);
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(listID, updates, null, null);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        1,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");
            DataTable updateResultData = AdapterHelper.ExtractData(result.Results[0].Any);
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);

            // If the first successful operation contains one and only one z:row
            // and the field value in the z:row element equals the value in the request,
            // R949 can be captured.
            bool isSuccess = result.Results[0].ErrorCode == "0x00000000"
                                        && updateResultData.Rows.Count == 1
                                        && Convert.ToString(updateResultData.Rows[0][columnName]) == newFieldValue;

            // This requirement is not fully verified, since the error case costs too much to simulate.
            // In this situation, we just verify the successful case.
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                2359,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] [In Batch element] [In Method element] [In Field element]
                    If the owshiddenversion value specified by the protocol client is equal to the owshiddenversion field's value for the list item on the protocol server, 
                    the protocol server MUST update the list item or return an error indicating why the update failed.");
        }

        #endregion

        #region UpdateListItem

        /// <summary>
        /// This test case is used to test UpdateListItem operation when inserting item using non-existing field name.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC71_UpdateListItem_FieldNotExist()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Call UpdateListItem operation to add one item using non-existent field name
            string fieldName = TestSuiteHelper.GetUniqueFieldName();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            cmds.Add(MethodCmdEnum.New);
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item = new Dictionary<string, string>();
            item.Add(fieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item);
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items);
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listName,
                                                                                updates);

            this.Site.Assert.AreEqual<int>(
                                1,
                                result.Results.Length,
                                "If only insert one item, then there will be one result element.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R942
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                942,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [In Field element] The protocol server MUST return an error indicating why the update failed, if the field is not a special field.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2366
            Site.CaptureRequirementIfIsNotNull(
                AdapterHelper.GetElementValue(result.Results[0].Any, "ErrorText"),
                2366,
                @"[UpdateListItemsResponse]Otherwise, [If an operation does not complete successfully] the ErrorText element MUST have a description of the error.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItem operation in the case that the value of Version attribute will be ignored by the server.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC72_UpdateListItem_IgnoreVersion()
        {
            // Create the first list
            string firstListID = TestSuiteHelper.CreateList();

            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(
                    Common.GetConfigurationPropertyValue("ListFieldText", this.Site),
                    Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            UpdateListItemsUpdates updates1 = TestSuiteHelper.CreateUpdateListItems(
                                                                new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                                new List<Dictionary<string, string>> { item1 });

            // Get a random version number.
            string firstVersion = TestSuiteHelper.GenerateVersionNumber();
            updates1.Batch.Version = firstVersion;
            UpdateListItemsResponseUpdateListItemsResult updateItemsResult1;
            updateItemsResult1 = this.listswsAdapter.UpdateListItems(
                                                    firstListID,
                                                    updates1);

            // Create the second list
            string secondListID = TestSuiteHelper.CreateList();

            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site),
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            UpdateListItemsUpdates updates2 = TestSuiteHelper.CreateUpdateListItems(
                                                                new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                                new List<Dictionary<string, string>> { item2 });

            // Get a random string number again.
            string secondVersion = TestSuiteHelper.GenerateVersionNumber();
            updates2.Batch.Version = secondVersion;

            UpdateListItemsResponseUpdateListItemsResult updateItemsResult2;
            updateItemsResult2 = this.listswsAdapter.UpdateListItems(
                                                    secondListID,
                                                    updates2);

            // In the z:row, make sure the attribute value which is specified by validFieldName is equal, 
            // for all the other attribute there is no guarantee. So all the other attributes in the z:row will 
            //  be ignored except the specified one.
            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in updateItemsResult1.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
                }
            }

            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in updateItemsResult2.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
                }
            }

            // Deeply compare updateItemsResult1 and updateItemsResult2, verify R2340
            bool isSame = TestSuiteHelper.DeepCompare(updateItemsResult1, updateItemsResult2);

            Site.CaptureRequirementIfIsTrue(
                isSame,
                2340,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [Version attribute] The protocol server MUST return the same response no matter the value of the Version attribute is 1.0.0.0001 or 2.0.0.0002.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItem operation when the SchemaLock attribute is set to TRUE and the ListVersion is invalid.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC73_UpdateListItem_InvalidListVersionWithSchemaLock()
        {
            // Create list.
            string listName = TestSuiteHelper.GetUniqueListName();
            string listID = TestSuiteHelper.CreateList(listName);

            // Call GetListItemChangesSinceToken with null change token to get current list version
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;
            result = this.listswsAdapter.GetListItemChangesSinceToken(
                                                    listID,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            uint currentVersion = (uint)result.listitems.Changes.List.Version;

            // Call UpdateListItems with lockSchema equals true and version is not the current version
            Dictionary<string, string> item = new Dictionary<string, string>();
            item.Add(
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site),
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site));
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(
                                                                new List<MethodCmdEnum> { MethodCmdEnum.New },
                                                                new List<Dictionary<string, string>> { item });

            // Set LockSchema TRUE
            updates.Batch.LockSchema = "TRUE";
            updates.Batch.ListVersionSpecified = true;

            // Plus the current version with 1, then the ListVersion is not valid.
            updates.Batch.ListVersion = currentVersion + 1;
            bool isVerifyR920 = false;
            UpdateListItemsResponseUpdateListItemsResult updateItemsResult;
            updateItemsResult = this.listswsAdapter.UpdateListItems(
                                                    listID,
                                                    updates);
            if (updateItemsResult != null)
            {
                if (updateItemsResult.Results != null && updateItemsResult.Results.Length == 1)
                {
                    if (updateItemsResult.Results[0].ErrorCode.Equals("0x80070666", StringComparison.OrdinalIgnoreCase))
                    {
                        isVerifyR920 = true;
                    }
                }
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R920
            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: isVerifyR920[{0}] for requirement #R920",
                     null == updateItemsResult.Results[0].ErrorCode ? "NullString" : updateItemsResult.Results[0].ErrorCode);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR920,
                920,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [LockSchema attribute] If TRUE and ListVersion is specified with a number, the protocol server MUST return a SOAP fault with error code 0x80070666 if the schema version passed in by the protocol client does not match the list schema version on the protocol server.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItem operation in the case that the value of OnError attribute is OnErrorReturn.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC74_UpdateListItem_OnErrorReturn()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            TestSuiteHelper.CreateList(listName);

            // Get the configuration value
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct three items, the second one is invalid, the first and the last one are valid
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item which is valid.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item which is invalid.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item which is valid.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listName,
                                                                                updates);

            this.Site.Assert.AreEqual<int>(
                                    3,
                                    result.Results.Length,
                                    "When insert three items, there are must be three result elements.");

            this.Site.Assert.AreEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                "When insert the first item, the result error code is not equal 0x800704c7");

            this.Site.Assert.AreNotEqual<string>(
                "0x00000000",
                result.Results[1].ErrorCode,
                "When insert the second item, the result error code is not equal 0x800704c7");

            // If third one contains "0x800704c7" then verify R961.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x800704c7",
                result.Results[2].ErrorCode,
                961,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In UpdateListItemsResult element] [If the OnError attribute of the Batch element is set to Return] If a Method operation fails, then all subsequent Method operation in the Batch MUST fail with an error of 0x800704c7.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItem operation when owsHiddenversion conflicts.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC75_UpdateListItem_OwsHiddenVersionConflict()
        {
            string listID = TestSuiteHelper.CreateList();

            // set list item count as 1.
            List<string> ids = TestSuiteHelper.AddListItems(listID, 1);

            // Get the owShiddenVersion
            string owsHiddenVersion = TestSuiteHelper.GetOwsHiddenVersion(listID);

            // Construct one update item with invalid owShiddenVersion.
            List<Dictionary<string, string>> updatedItems = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.Update);
            Dictionary<string, string> item = new Dictionary<string, string>();

            // Set the updated item id
            item.Add(AdapterHelper.FieldIDName, ids[0]);

            // Set the invalid owsVersion whose value is current owshiddenversion plus 1.
            string invalidowsVersion = (int.Parse(owsHiddenVersion) + 1).ToString();
            item.Add(AdapterHelper.FieldOwshiddenversionName, invalidowsVersion);
            item.Add(
                Common.GetConfigurationPropertyValue("ListFieldText", this.Site),
                TestSuiteHelper.GenerateRandomString(5));
            updatedItems.Add(item);
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(
                                                                cmds,
                                                                updatedItems);

            // IF the result contains 0x81020015 error code, verify R953.
            // A return value represents the schema definition which is read from specified XSD file.
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                    listID,
                                                                                    updates);
            Site.Assert.AreEqual<int>(
                                1,
                                result.Results.Length,
                                "After updating one item using UpdateListItems with invalid owshiddenversion, the server MUST return one result");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R964
            bool isR964Verified = int.Parse(result.Results[0].Version) == int.Parse(invalidowsVersion) + 1;

            Site.CaptureRequirementIfIsTrue(
                isR964Verified,
                964,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In UpdateListItemsResult element] The Version attribute MUST be the owshiddenversion contained in the UpdateListItems request plus 1.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R953
            Site.CaptureRequirementIfAreEqual<string>(
                "0x81020015",
                result.Results[0].ErrorCode,
                953,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [In Field element] If the owshiddenversion specified by the protocol client is different than the current value of the owshiddenversion field's value for the list item on the protocol server, the protocol server MUST return error code 0x81020015.");

            bool isListAttributeExist = result.Results[0].List != null;
            bool isVersionAttributeExist = result.Results[0].Version != null;

            Site.Assert.IsTrue(
                        isListAttributeExist,
                        "When returns error code 0x81020015 in the operation UpdateListItems, the list attribute must be exist.");

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2114
            bool isVersionAndListAttrExist = isVersionAttributeExist && isListAttributeExist;

            Site.Log.Add(
                    LogEntryKind.Debug,
                    "The actual value: isVersionAndListAttrExist[{0}] for requirement #R2114",
                    string.IsNullOrEmpty(result.Results[0].Version) ? "NullOrEmpty" : result.Results[0].Version);

            Site.CaptureRequirementIfIsTrue(
                isVersionAndListAttrExist,
                2114,
                @"[UpdateListItemsResponse]The List and Version attributes MUST be returned if a Method operation fails with an error of 0x81020015.");
        }

        #endregion

        #region UpdateListItemsWithKnowledge

        /// <summary>
        ///  This test case is used to test "DateInUtc" attribute in UpdateListItemsWithKnowledge operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC76_UpdateListItemsWithKnowledge_DateInUtc()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();

            // Construct one insert item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            // Set "Batch" element with "DateInUtc" attribute as TRUE, then call UpdateListItems method.
            updates.Batch.DateInUtc = "TRUE";
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listId,
                                                                                updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                    1,
                                    result.Results.Length,
                                    "There are at least one result corresponding with the one New method");
            DataTable data = AdapterHelper.ExtractData(result.Results[0].Any);
            string createdDateTime = data.Rows[0]["ows_Created"].ToString();
            bool isUtcFormat = createdDateTime.IndexOf('T') > 0;
            Site.CaptureRequirementIfIsTrue(
                isUtcFormat,
                2345,
                "[In UpdateListItemsWithKnowledge][In Bench element][In DateInUtc attribute] If set "
                + "DaeInUtc to TRUE, then dates passed up and returned are in UTC format.");

            // Set "Batch" element with "DateInUtc" attribute as FALSE, then call UpdateListItems method.
            updates.Batch.DateInUtc = "FALSE";
            result = this.listswsAdapter.UpdateListItems(
                                                        listId,
                                                        updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                    1,
                                    result.Results.Length,
                                    "There are at least one result corresponding with the one New method");
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            createdDateTime = data.Rows[0]["ows_Created"].ToString();
            bool isNotUtcFormat = createdDateTime.IndexOf('T') < 0;
            Site.CaptureRequirementIfIsTrue(
                isNotUtcFormat,
                2365,
                "[In UpdateListItemsWithKnowledge][In Bench element][In DateInUtc attribute] If set "
                    + "DaeInUtc to FALSE, then dates passed up and returned are not in UTC format.");

            // Set "Batch" element with "DateInUtc" attribute as null, then call UpdateListItems method.
            updates.Batch.DateInUtc = null;
            result = this.listswsAdapter.UpdateListItems(
                                                        listId,
                                                        updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                    1,
                                    result.Results.Length,
                                    "There are at least one result corresponding with the one New method");
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            createdDateTime = data.Rows[0]["ows_Created"].ToString();
            isNotUtcFormat = createdDateTime.IndexOf('T') < 0;
            Site.CaptureRequirementIfIsTrue(
                isNotUtcFormat,
                2346,
                "[In UpdateListItemsWithKnowledge][In Bench element][In DateInUtc attribute]"
                    + "[Specifies whether dates passed up and returned are in UTC format. ]This defaults to FALSE");
        }

        /// <summary>
        /// This test case is used to test UpdateListItemsWithKnowledge operation when the specified listName is empty.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC77_UpdateListItemsWithKnowledge_Fail_ListNameIsEmpty()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Read configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string errorCode = string.Empty;

            // Identify whether the SoapException is returned.
            bool isSoapRaultReturned = false;

            // Construct one insert item
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instances
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            Site.Assert.IsNotNull(updates, "Created UpdateListItemsWithKnowledgeUpdates successfully");

            try
            {
                // Call UpdateListItemsWithKnowledge with an empty ListName.
                this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                string.Empty,
                                                                updates,
                                                                null,
                                                                null);
            }
            catch (SoapException soapException)
            {
                isSoapRaultReturned = true;
                errorCode = TestSuiteHelper.GetErrorCode(soapException);

                #region Capture Requirements R3042

                // If there throws soap exception, and with no error code, then R3024 should be covered.
                Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: errorCode[{0}] for requirement #R3024",
                    string.IsNullOrEmpty(errorCode) ? "NullOrEmpty" : errorCode);

                Site.CaptureRequirementIfIsTrue(
                    string.IsNullOrEmpty(errorCode),
                    3024,
                    @"[In UpdateListItemsWithKnowledge] If the specified listName is empty,
                the protocol server MUST return a SOAP fault with no error code.");

                #endregion
            }

            Site.Assert.IsTrue(isSoapRaultReturned, "UpdateListItemsWithKnowledgeUpdates operation throws SoapException when the specified listName is empty string.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItemsWithKnowledge operation when listName does not correspond to a list. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC78_UpdateListItemsWithKnowledge_Fail_ListNameNotExists()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Construct an invalid ListId.
            string invalidListId = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Read configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string errorCode = string.Empty;

            // Construct one insert item
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instances
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            Site.Assert.IsNotNull(updates, "Created UpdateListItemsWithKnowledgeUpdates successfully");

            #region Invoke operations
            try
            {
                // Call UpdateListItemsWithKnowledge with invalid ListId.
                this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                invalidListId,
                                                                updates,
                                                                null,
                                                                null);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
            }
            #endregion

            #region Capture Requirements R1153, R1154

            Site.CaptureRequirementIfAreEqual<string>(
                "0x82000006",
                errorCode,
                1153,
                @"[In UpdateListItemsWithKnowledge]If listName does not correspond to a list from "
                + "either of these checks, the protocol server MUST return a SOAP fault with error "
                + "code 0x82000006.");

            Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    errorCode,
                1154,
                @"[In UpdateListItemsWithKnowledge] [If listName does not correspond to a list from "
                + "either of these checks, the protocol server MUST return a SOAP fault with error "
                + "code 0x82000006.] This indicates that the list does not exist or might have been "
                + "deleted by another user.");

            #endregion
        }

        /// <summary>
        ///  This test case is used to test LockSchema" attribute's default false value in UpdateListItemsWithKnowledge operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC79_UpdateListItemsWithKnowledge_LockSchema()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Get the Field name which has been add into current list.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Call GetList to get the current list version.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(listId);

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();

            // Construct one insert item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);

            // Make sure the LockSchema is not specified.
            // This will make the server treat it as default false value.
            updates.Batch.LockSchema = null;

            // Set an incorrect list version value.
            updates.Batch.ListVersion = (uint)listDef.Version + 1;
            updates.Batch.ListVersionSpecified = true;

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result =
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                    listId,
                                                    updates,
                                                    null,
                                                    null);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                    1,
                                    result.Results.Length,
                                    "There are at least one result corresponding with the one New method");

            // If there is no SoapException and the error code equals 0x00000000,
            // capture requirement R2085.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                2349,
                "[In UpdateListItemsWithKnowledge][In Bench element][In LockSchema attribute] "
                        + "This defaults to FALSE.[ the server will not fail, if the schema version passed in "
                        + "by the protocol client does not match the list schema version on the protocol server. ]");
        }

        /// <summary>
        /// This method is used to test the UpdateListItemsWithKnowledge operation when the Method element's Cmd attribute uses "New" "Update" "Delete" "Move" value.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC80_UpdateListItemsWithKnowledge_Methods()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Read value from configuration file.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct one insert item with ID.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            Dictionary<string, string> newItem = new Dictionary<string, string>();
            newItem.Add(validFieldName, fieldValue);
            newItem.Add("ID", Guid.NewGuid().ToString());
            items.Add(newItem);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsWithKnowledgeUpdates instance.
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult resultWhenIncludeId = null;
            resultWhenIncludeId = this.listswsAdapter.UpdateListItemsWithKnowledge(listId, updates, null, null);

            // Get new item ID from UpdateListItemsWithKnowledge response.
            DataTable data = AdapterHelper.ExtractData(resultWhenIncludeId.Results[0].Any);
            string columnNameId = string.Format("{0}{1}", AdapterHelper.PrefixOws, AdapterHelper.FieldIDName);

            List<string> updatedListItemIds = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                updatedListItemIds.Add(row[columnNameId].ToString());
            }

            // Get item IDs from GetListItems response.
            GetListItemsResponseGetListItemsResult getListItemsResult = this.listswsAdapter.GetListItems(listId, null, null, null, null, null, null);
            data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            List<string> getListItemIds = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                getListItemIds.Add(row[columnNameId].ToString());
            }

            // If new item ID exists in the GetListItems response, it means the new item has been added. R20921 is captured.
            bool isR20921Verified = true;
            foreach (string id in updatedListItemIds)
            {
                if (!getListItemIds.Contains(id))
                {
                    isR20921Verified = false;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isR20921Verified,
                20921,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [New attribute] Adds a new list item to the specified list.");

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one.  
            foreach (UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResultResult result in resultWhenIncludeId.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // Re-construct one insert item without ID.
            items.Clear();
            cmds.Clear();
            newItem.Clear();
            newItem.Add(validFieldName, fieldValue);
            items.Add(newItem);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsWithKnowledgeUpdates instance.
            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult resultWhenExcludeId = null;
            resultWhenExcludeId = this.listswsAdapter.UpdateListItemsWithKnowledge(listId, updates, null, null);

            data = AdapterHelper.ExtractData(resultWhenExcludeId.Results[0].Any);
            foreach (DataRow row in data.Rows)
            {
                updatedListItemIds.Add(row[columnNameId].ToString());
            }

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one.
            foreach (UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResultResult result in resultWhenExcludeId.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // If ignore some field in z:row element and both result are same, capture R2289.
            bool isSame = TestSuiteHelper.DeepCompare(resultWhenIncludeId, resultWhenExcludeId);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2351,
                    "[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] "
                    + "[In Batch element] [In Method element] [New attribute] The server response will be "
                    + "same in both <Field Name=\"ID\"> exists in request or not.");

            // There are 2 list items. Update the first item and delete the last item.
            items.Clear();
            cmds.Clear();

            Dictionary<string, string> updatedItem = new Dictionary<string, string>();
            string newFieldValue = TestSuiteHelper.GenerateRandomString(10);
            updatedItem.Add("ID", updatedListItemIds.First());
            updatedItem.Add(validFieldName, newFieldValue);
            items.Add(updatedItem);
            cmds.Add(MethodCmdEnum.Update);

            Dictionary<string, string> deletedItem = new Dictionary<string, string>();
            deletedItem.Add("ID", updatedListItemIds.Last());
            items.Add(deletedItem);
            cmds.Add(MethodCmdEnum.Delete);

            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);

            // Call UpdateListItemsWithKnowledge operation to update the first item and delete the last item.
            this.listswsAdapter.UpdateListItemsWithKnowledge(listId, updates, null, null);

            // Call GetListItems operation.
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { validFieldName });
            getListItemsResult = this.listswsAdapter.GetListItems(listId, null, null, viewFields, null, null, null);
            data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            getListItemIds.Clear();
            string columnNameField = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);
            bool isR20931Verified = false;

            foreach (DataRow row in data.Rows)
            {
                // Get item IDs from GetListItems response.
                getListItemIds.Add(row[columnNameId].ToString());

                // Check if the updated item has been updated successfully.
                if (row[columnNameId].ToString().Equals(updatedListItemIds.First(), StringComparison.OrdinalIgnoreCase))
                {
                    if (row[columnNameField].ToString().Equals(newFieldValue, StringComparison.Ordinal))
                    {
                        isR20931Verified = true;
                    }
                }
            }

            // If the updated item in GetListItems response has been updated successfully, R20931 is captured.
            Site.CaptureRequirementIfIsTrue(
                isR20931Verified,
                20931,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [Update attribute] Updates fields for a specific list item.");

            // If GetListItems response does not contain the deleted item ID, it means the item has been deleted. Then R20941 is captured.
            Site.CaptureRequirementIfIsFalse(
                getListItemIds.Contains(updatedListItemIds.Last()),
                20941,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [Delete attribute]Deletes a specific list item.");

            // Verify 'Move' method.
            if (Common.IsRequirementEnabled(20962, this.Site))
            {
                // Create 2 document libraries.
                string sourceListName = TestSuiteHelper.GetUniqueListName();
                string sourceList = TestSuiteHelper.CreateList(sourceListName, Convert.ToInt32(TemplateType.Document_Library));

                string destinationListName = TestSuiteHelper.GetUniqueListName();
                string destinationList = TestSuiteHelper.CreateList(destinationListName, Convert.ToInt32(TemplateType.Document_Library));

                // Get the root folder of destination list.
                GetListResponseGetListResult getListResult = this.listswsAdapter.GetList(destinationList);
                string rootFolder = getListResult.List.RootFolder;

                // Upload a document to source list.
                IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();
                string absoluteFileUrl = sutControlAdapter.UploadFile(sourceListName);

                Site.Assert.IsTrue(
                    !string.IsNullOrEmpty(absoluteFileUrl),
                    "Upload file to the list {0} should be successful, the file path is [{1}]",
                    sourceListName,
                    absoluteFileUrl);

                // Get the uploaded file name.
                string uploadFileName = absoluteFileUrl.Substring(absoluteFileUrl.LastIndexOf('/'));

                // Get uploaded item ID.
                getListItemsResult = this.listswsAdapter.GetListItems(sourceList, null, null, null, null, null, null);
                data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

                // Get the value of "ID" and "MoveNewUrl" fields.                
                string fileID = (string)data.Rows[0][columnNameId];
                string moveNewUrl = rootFolder.TrimEnd('/') + "/" + uploadFileName.Trim('/');

                // Construct one insert item with ID.
                items.Clear();
                cmds.Clear();
                Dictionary<string, string> movedItem = new Dictionary<string, string>();
                movedItem.Add("ID", fileID);
                movedItem.Add("FileRef", absoluteFileUrl);
                movedItem.Add("MoveNewUrl", moveNewUrl);
                items.Add(movedItem);
                cmds.Add(MethodCmdEnum.Move);

                // Construct the UpdateListItemsWithKnowledgeUpdates instance.
                UpdateListItemsWithKnowledgeUpdates moveUpdates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);

                // Call UpdateListItemsWithKnowledge method to move file.
                UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult moveResult = null;
                moveResult = this.listswsAdapter.UpdateListItemsWithKnowledge(sourceList, moveUpdates, null, null);

                // Error code "0x00000000" indicates the operation completes successfully. Then R20962 captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x00000000",
                    moveResult.Results[0].ErrorCode,
                    20962,
                    @"Implementation does support this behavior[UpdateListItemsWithKnowledge.Move]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to test the attribute "OnError" with value "Continue" or default value in UpdateListItemsWithKnowledge operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC81_UpdateListItemsWithKnowledge_OnErrorContinue()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Get the random generate unique field name.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listId,
                                                                updates,
                                                                null,
                                                                null);

            // Capture R2315 if there are three result elements,
            // because we have there method in the request.
            Site.CaptureRequirementIfAreEqual<int>(
                            3,
                            result.Results.Length,
                            2315,
                            "[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                                + "[In UpdateListItemsWithKnowledgeResult element] If the OnError attribute of the Batch "
                                + "element is set to Continue, then each Method corresponds to a Result.");

            // Call GetListItems to get current items.
            GetListItemsResponseGetListItemsResult listItems = null;
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            DataTable data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // If there is only one item exist, the third insert will be aborted due to 
            // the failure of the second insert operation.
            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 1507,
                                 @"[OnErrorEnum]Continue: Continues running subsequent methods after an error is encountered.");

            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 2316,
                                "[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                                    + "[In UpdateListItemsWithKnowledgeResult element] [If the OnError attribute of the Batch "
                                    + "element is set to Continue,] The outcome of a Method operation MUST not affect the other Method operations.");

            if (Common.IsRequirementEnabled(4010, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<int>(
                                     2,
                                     data.Rows.Count,
                                     4010,
                                     @"[In UpdateListItemsWithKnowledge][In Bench element][In OnError attribute] Implementation does execute the subsequent methods in the batch [if ""OnError=Continue""] when encountering an error."
                                     + @"(The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }

            // Remove all the list items in this list.
            string colunmIdName = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            List<string> allItemsId = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                string id = Convert.ToString(row[colunmIdName]);
                allItemsId.Add(id);
            }

            TestSuiteHelper.RemoveListItems(listId, allItemsId, OnErrorEnum.Continue);

            // Construct UpdateListItemsUpdates instance not specified OnError, using the default value.
            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items);

            // Call UpdateListItems using the default OnError value.
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(listId, updates, null, null);

            // Call the GetListItems to get the current list items.
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // The UpdateListItmes does not send OnError, but if there are still 2 items,
            // the third insert won’t be aborted due to the failure of the second insert operation.
            // The behavior is the same on Continue.
            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 2343,
                                @"[In UpdateListItemsWithKnowledge][In Bench element][In OnError attribute] "
                                    + "The default is to continue executing subsequent items.");

            // Remove all the list items in this list.
            allItemsId = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                string id = Convert.ToString(row[colunmIdName]);
                allItemsId.Add(id);
            }

            TestSuiteHelper.RemoveListItems(listId, allItemsId, OnErrorEnum.Continue);

            // Try to update an item.
            Dictionary<string, string> itemOfUpdated = new Dictionary<string, string>();
            itemOfUpdated.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));

            // Set the list item ID equal to 1 and Current List does not contain any list item. The list item which "ID" is equal to 1 does not exist in Current List.
            itemOfUpdated.Add("ID", "1");
            items = new List<Dictionary<string, string>>();
            items.Add(itemOfUpdated);
            cmds = new List<MethodCmdEnum>();
            cmds.Add(MethodCmdEnum.Update);

            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult resultOfUpdateListItemWithKnowledge = null;
            resultOfUpdateListItemWithKnowledge = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listId,
                                                                updates,
                                                                null,
                                                                null);

            // Here is one method in request, so there is one matched result element in response.
            if (null == resultOfUpdateListItemWithKnowledge || null == resultOfUpdateListItemWithKnowledge.Results
                || 1 != resultOfUpdateListItemWithKnowledge.Results.Length)
            {
                this.Site.Assert.Fail("Could not get the expected result information from response of UpdateListItemsWithKnowledge operation");
            }

            // Here is one list item in request, so that there is one z:row that should be returned.
            if (null == resultOfUpdateListItemWithKnowledge.Results[0].Any || 1 != resultOfUpdateListItemWithKnowledge.Results[0].Any.Length)
            {
                this.Site.Assert.Fail("Could not get the updated list item detail from response of UpdateListItemsWithKnowledge operation");
            }

            // If the SUT protocol presents the error code "0x81020016" in the response of resultOfUpdateListItemWithKnowledge operation,
            // capture R2328.
            Site.CaptureRequirementIfAreEqual<string>(
                         "0x81020016",
                         resultOfUpdateListItemWithKnowledge.Results[0].ErrorCode,
                           2328,
                           "[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                        + "[In UpdateListItemsWithKnowledgeResult element]  This [error code 0x81020016] "
                        + "specifies that the list item referred to in the request does not exist.");
        }

        /// <summary>
        ///  This test case is used to test the server behavior when the value of OnError attribute in the input parameters is “Return” in UpdateListItemsWithKnowledge operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC82_UpdateListItemsWithKnowledge_OnErrorReturn()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Get the random generate unique field name.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listId,
                                                                updates,
                                                                null,
                                                                null);

            // Capture R2313 if there are three result elements,
            // because we have three methods in the request.
            Site.CaptureRequirementIfAreEqual<int>(
                            3,
                            result.Results.Length,
                            2313,
                            @"[In UpdateListItemsWithKnowledgeResult][In UpdateListItemsWithKnowledgeResponse]"
                            + "[In UpdateListItemsWithKnowledgeResult element] If the OnError attribute of the Batch "
                            + "element is set to Return, then there MUST be a Result element for each Method operation.");

            // If there are any error, this means that
            // the server cannot set the column with the specified value. Then capture requirement R2352.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0x00000000",
                result.Results[1].ErrorCode,
                2352,
                "[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element]"
                    + "[In Batch element] [In Method element] [New attribute] [If Additional field "
                    + "references DO NOT set those columns to the value specified,] an error will be returned.");

            if (result.Results[1].ErrorCode == "0x81020014")
            {
                // If the error equals 0x81020014, the second item insert will fail. Then capture requirement 2120.
                // There is no mandatory rule in this document to specify the server MUST return 0x81020014 in this case,
                // So there is no force rule to capture this requirement.
                Site.CaptureRequirement(
                    2327,
                    @"[In UpdateListItemsWithKnowledgeResult][In "
                        + "UpdateListItemsWithKnowledgeResponse][In "
                        + "UpdateListItemsWithKnowledgeResult element] This [0x81020014] "
                        + "specifies a generic error has been encountered, such as an invalid value "
                        + "being specified for a Field.");
            }

            // Call GetListItems to get current items.
            GetListItemsResponseGetListItemsResult listItems = null;
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            DataTable data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // If there is only one item existing, the third insert will be aborted due to 
            // the failure of the second insert operation.
            Site.CaptureRequirementIfAreEqual<int>(
                                 1,
                                 data.Rows.Count,
                                 1506,
                                "[OnErrorEnum]Return: [If the request includes the simple type with value \"Return\", "
                                    + "the protocol server] Stops any more methods from running after the first error is encountered.");

            if (Common.IsRequirementEnabled(4008, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<int>(
                                     1,
                                     data.Rows.Count,
                                     4008,
                                     @"[In UpdateListItemsWithKnowledge][In Bench element][In OnError attribute] Implementation does result in the entire operation"
                                     + @"being aborted [if ""OnError=Return""],[or whether subsequent methods in the batch should be executed.] when encountering an error."
                                     + @"(The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to test UpdateListItemsWithKnowledge operation in the case that the attribute "PreCalc" will be ignored by the server.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC83_UpdateListItemsWithKnowledge_PreCalcIgnore()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct one insert item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instances and set the PreCalc attribute to true.
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Return);
            updates.Batch.PreCalc = "TRUE";
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult resultWhenPreCalcTrue = null;
            resultWhenPreCalcTrue = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listId,
                                                                updates,
                                                                null,
                                                                null);

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one. 
            foreach (UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResultResult result in resultWhenPreCalcTrue.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // Reset the PreCalc to some invalid value.
            updates.Batch.PreCalc = "NotValidBoolean";
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult resultWhenPreCalcInvalid = null;
            resultWhenPreCalcInvalid = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                listId,
                                                                updates,
                                                                null,
                                                                null);

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one. 
            foreach (UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResultResult result in resultWhenPreCalcInvalid.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // If ignore some fields in z:row element and both results are the same then capture R2344.
            bool isSame = TestSuiteHelper.DeepCompare(resultWhenPreCalcTrue, resultWhenPreCalcInvalid);
            Site.CaptureRequirementIfIsTrue(
                        isSame,
                        2344,
                        "[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] "
                            + "[In updates element] [In Batch element] [PreCalc attribute] The response of server will "
                            + "be same for different values of PreCalc.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItemsWithKnowledge operation when the listName is a valid GUID. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC84_UpdateListItemsWithKnowledge_Success()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            // Read configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Create a list
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Construct two insert items, one is valid and the other one is invalid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(2);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(2);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, fieldValue);
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instances
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            Site.Assert.IsNotNull(updates, "Created UpdateListItemsWithKnowledgeUpdates successfully");

            // Call UpdateListItemsWithKnowledge with valid Guid.
            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult updateListItemsWithKnowledgeResult = null;
            bool isSoapExceptionExist = false;
            try
            {
                updateListItemsWithKnowledgeResult = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                                    listId,
                                                                    updates,
                                                                    null,
                                                                    null);
            }
            catch (SoapException)
            {
                isSoapExceptionExist = true;
            }

            this.Site.Assert.IsNotNull(updateListItemsWithKnowledgeResult, "UpdateListItemWithKnowledge operation with valid GUID is successful.");

            // If there is no any SoapException exist, then capture R25151.
            this.Site.CaptureRequirementIfIsFalse(
                    isSoapExceptionExist,
                    25151,
                    @"Implementation does support this method[UpdateListItemsWithKnowledge]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");

            #region Capture Requirement R1155

            // Verify if the FieldName is valid, the ErrorCode is equal to "0x00000000".
            bool isValisFieldName = updateListItemsWithKnowledgeResult.Results[0].ErrorCode.Equals("0x00000000");
            Site.Assert.IsTrue(isValisFieldName, "the ErrorCode is equal to 0x00000000.");

            // Verify if the FiledName is invalid, the ErrorCode is not equal to "0x00000000"".
            bool isInValidFieldName = !updateListItemsWithKnowledgeResult.Results[1].ErrorCode.Equals("0x00000000");
            Site.Assert.IsTrue(isInValidFieldName, "the ErrorCode is equal not to 0x00000000.");

            // Verify R1155.
            Site.CaptureRequirement(
                1155,
                @"[In UpdateListItemsWithKnowledge]Otherwise, the protocol server MUST process "
                + "the batched operations on the list and return success or failure conditions per "
                + "operation.");

            #endregion

            #region Capture Requirement R1151, R2358

            // Call GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult operation by using the List Guid 
            // to verify whether the item exists in the specified list.
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(false, new List<string> { validFieldName });
            GetListItemChangesWithKnowledgeResponseGetListItemChangesWithKnowledgeResult getListItemWithKnowledgeResult = null;
            getListItemWithKnowledgeResult = this.listswsAdapter.GetListItemChangesWithKnowledge(
                                                                                                  listId,
                                                                                                  null,
                                                                                                  null,
                                                                                                  fields,
                                                                                                  null,
                                                                                                  null,
                                                                                                  null,
                                                                                                  null,
                                                                                                  null);

            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);
            DataTable data = AdapterHelper.ExtractData(getListItemWithKnowledgeResult.listitems.data.Any);
            bool isUseList = data.Rows.Count == 1 && Convert.ToString(data.Rows[0][columnName]) == fieldValue;

            Site.Log.Add(
                     LogEntryKind.Debug,
                     "The actual value: data.Rows.Count[{0}],columnName[{1}] for requirement #R1151",
                     data.Rows.Count,
                     0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            // If the specified list contains one item and the field value in the z:row element equals to value in
            // the UpdateListItemsWithKnowledge request, then R1151 should be covered.
            Site.CaptureRequirementIfIsTrue(
                isUseList,
                1151,
                @"[In UpdateListItemsWithKnowledge]If the specified listName is a valid GUID and corresponds to 
                the identification of a list on the site, use that list.");

            // If the protocol client does not specify the owshiddenversion field reference in 
            // the UpdateListItemsWithKnow request, then capture R2358.
            Site.CaptureRequirement(
               2358,
               @"[In UpdateListItemsWithKnowledge operation] [In "
               + "UpdateListItemsWithKnowledge element] [In updates element] [In Batch "
               + "element] [In Method element] [In Field element] If the protocol client does "
               + "not specify the owshiddenversion field reference in the "
               + "UpdateListItemsWithKnow request, then the protocol server MUST overwrite "
               + "any changes in the list item or return an error.");

            #endregion

            #region Capture Requirements R2326

            string fieldCounter = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);

            // Set the value of Counter as "1".
            string fieldCounterValue = "1";
            cmds = new List<MethodCmdEnum>(1);
            items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(fieldCounter, fieldCounterValue);
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            // Call UpdateListItemsWithKnowledge with valid Guid. Add a counter.
            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            updateListItemsWithKnowledgeResult
           = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                               listId,
                                                               updates,
                                                               null,
                                                               null);

            // Get the value of the "ID" from the server.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            data = AdapterHelper.ExtractData(updateListItemsWithKnowledgeResult.Results[0].Any);
            string id = Convert.ToString(data.Rows[0][columnName]);

            // As description, if update the list item, must add the ID value for the specified list item.
            cmds = new List<MethodCmdEnum>(1);
            items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item4 = new Dictionary<string, string>();
            item4.Add("ID", id);

            // Set an invalid value for the counter, Guid is not integer. 
            item4.Add(fieldCounter, Guid.NewGuid().ToString());
            items.Add(item4);
            cmds.Add(MethodCmdEnum.Update);
            updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            updateListItemsWithKnowledgeResult
           = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                               listId,
                                                               updates,
                                                               null,
                                                               null);

            // If the specified item is not be changed, then R2326 should be covered.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, fieldCounter);
            data = AdapterHelper.ExtractData(updateListItemsWithKnowledgeResult.Results[0].Any);
            bool isReturnInZRowElement = Convert.ToString(data.Rows[0][columnName]) == fieldCounterValue;

            Site.Log.Add(
            LogEntryKind.Debug,
            "The actual value: columnName[{0}] for requirement #R2326",
            0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            Site.CaptureRequirementIfIsTrue(
                isReturnInZRowElement,
                2326,
                @"[In UpdateListItemsWithKnowledgeResult][In "
                + "UpdateListItemsWithKnowledgeResponse][In "
                + "UpdateListItemsWithKnowledgeResult element] On failure, the z:row element "
                + "MUST contain the list item data that exists on the server if the list item "
                + "specified in the request exists.");

            #endregion
        }

        /// <summary>
        /// This test case is used to test the UpdateListItemsWithKnowledge operation when listName is not a GUID but a valid list title. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC85_UpdateListItemsWithKnowledge_ValidListTitle()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(25151, this.Site), @"Test is executed only when R25151Enabled is set to true.");

            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Get the configure value.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct two insert items, one is valid and the other one is invalid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(2);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(2);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, fieldValue);
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsWithKnowledgeUpdates updates = TestSuiteHelper.CreateUpdateListWithKnowledgeItems(cmds, items, OnErrorEnum.Continue);

            UpdateListItemsWithKnowledgeResponseUpdateListItemsWithKnowledgeResult result = null;
            result = this.listswsAdapter.UpdateListItemsWithKnowledge(
                                                listName,
                                                updates,
                                                null,
                                                null);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        2,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");

            // Expect the first result is successful but the second one is failed.
            bool isFirstResultSuccess = result.Results[0].ErrorCode == "0x00000000";
            bool isSecondResultFail = result.Results[1].ErrorCode != "0x00000000";

            // If the first operation result equals to 0x00000000, capture 
            // Requirement R2323.
            Site.CaptureRequirementIfIsTrue(
                isFirstResultSuccess,
                2323,
                @"[In UpdateListItemsWithKnowledgeResult][In "
                    + "UpdateListItemsWithKnowledgeResponse][In "
                    + "UpdateListItemsWithKnowledgeResult element] If an operation completes "
                    + "successfully, the value of the ErrorCode MUST be set to 0x00000000.");

            DataTable resultData1 = AdapterHelper.ExtractData(result.Results[0].Any);
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);

            // If the first successful operation contains one and only one z:row
            // and the field value in the z:row element equals the value in the request,
            // the requirement R2325 will be able to be captured.
            bool isContainUpdateData = resultData1.Rows.Count == 1
                                            && Convert.ToString(resultData1.Rows[0][columnName]) == fieldValue;
            Site.CaptureRequirementIfIsTrue(
                isContainUpdateData,
                2325,
                @"[In UpdateListItemsWithKnowledgeResult][In "
                    + "UpdateListItemsWithKnowledgeResponse][In "
                    + "UpdateListItemsWithKnowledgeResult element]If the operation succeeded, "
                    + "the z:row element MUST contain the updated data for the list item that is "
                    + "operated on.");

            // Call GetListItems operation by using the List GUID to verify 
            // the operation of inserting item just happens to the specified list.
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(false, new List<string> { validFieldName });
            GetListItemsResponseGetListItemsResult itemResult = null;
            itemResult = this.listswsAdapter.GetListItems(
                                                listId,
                                                null,
                                                null,
                                                fields,
                                                null,
                                                null,
                                                null);

            // If the list specified by the GUID only contains one item
            // and the field value in the z:row element equals the value in the UpdateListItems request,
            // we use the proper list in the UpdateListItems operation.
            // R1152 and R2160 can be captured in this condition.
            DataTable data = AdapterHelper.ExtractData(itemResult.listitems.data.Any);
            bool isUseProperList = data.Rows.Count == 1
                                    && Convert.ToString(data.Rows[0][columnName]) == fieldValue;
            Site.CaptureRequirementIfIsTrue(
                isUseProperList,
                1152,
                @"[In UpdateListItemsWithKnowledge]If the specified listName is not a valid "
                    + "GUID [or does not correspond to the identification of a list on the site,] "
                    + "check if the listName corresponds to the list title of a list on the site and, "
                    + "if so, use that list.");
            Site.CaptureRequirementIfIsTrue(
                isUseProperList,
                2160,
                @"[If the specified listName ]does not correspond to the identification of a list "
                    + "on the site, check if the listName corresponds to the list title of a list on the "
                    + "site and, if so, use that list.");

            // If we make sure using the proper list
            // and the operation result contains either successful or failed result,
            // R1155 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isUseProperList && isFirstResultSuccess && isSecondResultFail,
                1155,
                @"[In UpdateListItemsWithKnowledge]Otherwise, the protocol server MUST process "
                + "the batched operations on the list and return success or failure conditions per "
                + "operation.");

            // If the field value of the list item equals the value in request, R2355 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                fieldValue,
                Convert.ToString(data.Rows[0][columnName]),
                2355,
                @"[In UpdateListItemsWithKnowledge operation] [In UpdateListItemsWithKnowledge element] [In updates element] "
                + "[In Batch element] [In Method element] [In Field element] The protocol server "
                + "MUST set the field's value equal to the value of the Field element in the protocol client request, "
                + "if the field is not a special field.");
        }

        #endregion

        #region UpdateListItems

        /// <summary>
        ///  This test case is used to test "DateInUtc" attribute in UpdateListItems operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC86_UpdateListItems_DateInUtc()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct one insert item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            // Set "Batch" element with "DateInUtc" attribute is TRUE, and then call UpdateListItems method.
            updates.Batch.DateInUtc = "TRUE";
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listId,
                                                                                updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                  1,
                                  result.Results.Length,
                                  "There are at least one result corresponding with the one New method");
            DataTable data = AdapterHelper.ExtractData(result.Results[0].Any);
            string createdDateTime = data.Rows[0]["ows_Created"].ToString();
            bool isUtcFormat = createdDateTime.IndexOf('T') > 0;
            Site.CaptureRequirementIfIsTrue(
                isUtcFormat,
                2363,
                "[In UpdateListItems][In Bench element][In DateInUtc attribute] If set "
                    + "DaeInUtc to TRUE, then dates passed up and returned are in UTC format.");

            // Set "Batch" element with "DateInUtc" attribute is FALSE, and then call UpdateListItems method.
            updates.Batch.DateInUtc = "FALSE";
            result = this.listswsAdapter.UpdateListItems(
                                                        listId,
                                                        updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                  1,
                                  result.Results.Length,
                                  "There are at least one result corresponding with the one New method");
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            createdDateTime = data.Rows[0]["ows_Created"].ToString();
            bool isNotUtcFormat = createdDateTime.IndexOf('T') < 0;
            Site.CaptureRequirementIfIsTrue(
                isNotUtcFormat,
                2364,
                "[In UpdateListItems][In Bench element][In DateInUtc attribute] If set "
                    + "DaeInUtc to FALSE, then dates passed up and returned are not in UTC format.");

            // Set "Batch" element with "DateInUtc" attribute is null, and then call UpdateListItems method.
            updates.Batch.DateInUtc = null;
            result = this.listswsAdapter.UpdateListItems(
                                                        listId,
                                                        updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                  1,
                                  result.Results.Length,
                                  "There are at least one result corresponding with the one New method");
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            createdDateTime = data.Rows[0]["ows_Created"].ToString();
            isNotUtcFormat = createdDateTime.IndexOf('T') < 0;
            Site.CaptureRequirementIfIsTrue(
                isNotUtcFormat,
                2307,
                "[In UpdateListItems][In Bench element][In DateInUtc attribute]"
                    + "[Specifies whether dates passed up and returned are in UTC format. ]This defaults to FALSE");
        }

        /// <summary>
        /// This test case is used to verify UpdateListItem operation in Windows SharePoint Services 3.0 when ListName parameter cannot be found in current existing lists.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC87_UpdateListItems_InvalidGUIDAndNotCorrespond_SP3WSS3()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2510, this.Site), @"Test is executed only when R2510Enabled is set to true.");

            // Initialize an invalid ListName structure.
            string invalidListName = TestSuiteHelper.GetInvalidGuidAndNocorrespondString();

            // Read value from configuration file.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);

            // Construct one insert item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instances and set the Pre Cal to true
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            Site.Assert.IsNotNull(updates, "UpdateListItemsUpdates success");

            bool issoapFaultGenerated = false;
            string errorcode = string.Empty;
            string errorString = string.Empty;
            string expectedErrorString = @"Guid should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).";
            try
            {
                // Call UpdateListItems method to make inserts, updates, and deletions to specified list items in a list.
                this.listswsAdapter.UpdateListItems(invalidListName, updates);
            }
            catch (SoapException soapException)
            {
                issoapFaultGenerated = true;
                errorcode = TestSuiteHelper.GetErrorCode(soapException);
                errorString = TestSuiteHelper.GetErrorString(soapException);
            }

            // Verify MS-LISTSWS requirement: MS-LISTSWS_R2510
            // If the value of the error is null and the detailed message of SoapException is equal to the 
            // specific value, then R2510 should be covered.
            this.Site.Log.Add(
                     LogEntryKind.Debug,
                    "The actual value: isSoapFault[{0}],ErrorCode[{1}],errorString[{2}] for requirement #R2510",
                    issoapFaultGenerated,
                    string.IsNullOrEmpty(errorcode) ? "NullOrEmpty" : errorcode,
                    string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);
            this.Site.Assert.IsTrue(issoapFaultGenerated, "There should be a soap fault generated when calling UpdateListItems operation with invalid listname.");
            this.Site.Assert.IsTrue(string.IsNullOrEmpty(errorcode), "There should be a soap fault generated when calling UpdateListItems operation with invalid listname.");

            Site.CaptureRequirementIfIsTrue(
                expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase),
                2510,
                @"[In Appendix B: Product Behavior] Implementation does return a SOAP fault with no error code. (<80> Section 3.1.4.31: Windows SharePoint Services 3.0 return the following SOAP fault with no error code: ""GUID should contain 32 digits with 4 dashes (xxxxxxxx-xxxx-xxxx-xxxx-xxxx-xxxxxxxxxxx)"".)");
        }

        /// <summary>
        /// This test case is used to test UpdateListItems operation when listName is in a not valid GUID and the list title does not exist.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC88_UpdateListItems_InvalidListTitle()
        {
            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct one new item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            UpdateListItemsResponseUpdateListItemsResult result = null;
            bool isSoapFault = false;
            string errorCode = null;
            try
            {
                result = this.listswsAdapter.UpdateListItems(
                                                    TestSuiteHelper.GetUniqueListName(),
                                                    updates);
            }
            catch (SoapException exp)
            {
                errorCode = TestSuiteHelper.GetErrorCode(exp);
                isSoapFault = true;
            }

            Site.Assert.IsTrue(
                    isSoapFault,
                    "When call the operation UpdateListItems, the list name is not exist GUID. The server will return a soap fault.");

            // If error code equals 0x82000006, capture R2770.
            if (Common.IsRequirementEnabled(2770, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x82000006",
                    errorCode,
                    2770,
                    @"[In UpdateListItems operation] Implementation does return a SOAP fault with error code 0x82000006, if listName does not correspond to a list from either of these checks. (SharePoint Foundation 2010 follow this behavior.)");
            }

            // If SUT is SharePointFoundation2013 or SharePointServer2013, the error code equals 0x82000006.
            string sutVersion = Common.GetConfigurationPropertyValue("SutVersion", this.Site);
            if (sutVersion == "SharePointFoundation2013" || sutVersion == "SharePointServer2013")
            {
                Site.Assert.AreEqual<string>(
                    "0x82000006",
                    errorCode,
                    @"{0} return a SOAP fault with error code 0x82000006, when the specified listName is a valid GUID and does not correspond to the identification of a list on the site.",
                    sutVersion);
            }

            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Construct one update item.
            items = new List<Dictionary<string, string>>();
            cmds = new List<MethodCmdEnum>();
            item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);

            // This list is empty, so just generate an ID randomly. Then the items should not exist.
            item1.Add(AdapterHelper.FieldIDName, TestSuiteHelper.GenerateRandomNumber(1, 9));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.Update);
            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            // Call UpdateListItems to use an non-exist list item id.
            result = this.listswsAdapter.UpdateListItems(
                                                listId,
                                                updates);

            this.Site.Assert.AreEqual<int>(
                                1,
                                result.Results.Length,
                                "After call UpdateListItems with one item updating in request, there must be one result element in the response");

            this.Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020016",
                     result.Results[0].ErrorCode,
                     975,
                    "[In UpdateListItems operation] [In UpdateListItemsResponse element]"
                        + "[In UpdateListItemsResult element] This [error code 0x81020016] specifies"
                        + "that the list item referred to in the request does not exist.");

            this.Site.CaptureRequirementIfAreEqual<string>(
                    "0x81020016",
                     result.Results[0].ErrorCode,
                     932,
                    "[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [Update attribute] [If additional field references not set those columns to the value specified] an error will be returned.");

            this.Site.Log.Add(
                             LogEntryKind.Debug,
                             "The actual value: ErrorCode[{0}] for requirement #R2117",
                             result.Results[0].ErrorCode);

            this.Site.CaptureRequirementIfIsTrue(
                      result.Results[0].ErrorCode.IndexOf("0x", StringComparison.OrdinalIgnoreCase) == 0,
                      2117,
                      @"[UpdateListItemsResponse]Otherwise, [If an operation does not complete successfully] the ErrorCode MUST be set to a hexadecimal representation of the error encountered.");
        }

        /// <summary>
        ///  This test case is used to test "LockSchema" attribute's default false value in UpdateListItems operation.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC89_UpdateListItems_LockSchema()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Call GetList to get the current list version.
            ListDefinitionSchema listDef = TestSuiteHelper.GetListDefinition(listId);

            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct three items: the second one is invalid, the first and the last one is valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Construct one insert item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            // Make sure the LockSchema is not specified.
            // This will make the server treat it as default false value.
            updates.Batch.LockSchema = null;

            // Set an incorrect list version value.
            updates.Batch.ListVersion = (uint)listDef.Version + 1;
            updates.Batch.ListVersionSpecified = true;

            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listId,
                                                                                updates);

            // There must be one result element.
            Site.Assert.AreEqual<int>(
                                  1,
                                  result.Results.Length,
                                  "There are at least one result corresponding with the one New method");

            // If there is no SoapException and the error code equals 0x00000000,
            // capture requirement R2085.
            Site.CaptureRequirementIfAreEqual<string>(
                "0x00000000",
                result.Results[0].ErrorCode,
                2085,
                "[In UpdateListItems][In Bench element][In LockSchema attribute] "
                 + "This defaults to FALSE.[ the server will not fail, if the schema version passed in "
                 + "by the protocol client does not match the list schema version on the protocol server. ]");
        }

        /// <summary>
        /// This method is used to test the UpdateListItems operation when the Method element's Cmd attribute uses "New" "Update" "Delete" "Move" value.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC90_UpdateListItems_Methods()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Read value from configuration file.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct one insert item with ID.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            Dictionary<string, string> newItem = new Dictionary<string, string>();
            newItem.Add(validFieldName, fieldValue);
            newItem.Add("ID", Guid.NewGuid().ToString());
            items.Add(newItem);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            UpdateListItemsResponseUpdateListItemsResult resultWhenIncludeId = null;
            resultWhenIncludeId = this.listswsAdapter.UpdateListItems(listId, updates);

            // Get new item ID from UpdateListItems response.
            DataTable data = AdapterHelper.ExtractData(resultWhenIncludeId.Results[0].Any);
            string columnNameId = string.Format("{0}{1}", AdapterHelper.PrefixOws, AdapterHelper.FieldIDName);

            List<string> updatedListItemIds = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                updatedListItemIds.Add(row[columnNameId].ToString());
            }

            // Get item IDs from GetListItems response.
            GetListItemsResponseGetListItemsResult getListItemsResult = this.listswsAdapter.GetListItems(listId, null, null, null, null, null, null);
            data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            List<string> getListItemIds = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                getListItemIds.Add(row[columnNameId].ToString());
            }

            // If new item ID exists in the GetListItems response, it means the new item has been added. R2092 is captured.
            bool isR2092Verified = true;
            foreach (string id in updatedListItemIds)
            {
                if (!getListItemIds.Contains(id))
                {
                    isR2092Verified = false;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isR2092Verified,
                2092,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [New attribute] Adds a new list item to the specified list.");

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one.  
            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in resultWhenIncludeId.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // Re-construct one insert item without ID.
            items.Clear();
            cmds.Clear();
            newItem.Clear();
            newItem.Add(validFieldName, fieldValue);
            items.Add(newItem);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            UpdateListItemsResponseUpdateListItemsResult resultWhenExcludeId = null;
            resultWhenExcludeId = this.listswsAdapter.UpdateListItems(listId, updates);

            data = AdapterHelper.ExtractData(resultWhenExcludeId.Results[0].Any);
            foreach (DataRow row in data.Rows)
            {
                updatedListItemIds.Add(row[columnNameId].ToString());
            }

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, and for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one.
            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in resultWhenExcludeId.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // If ignore some fields in z:row element and both results are the same, capture R2289.
            bool isSame = TestSuiteHelper.DeepCompare(resultWhenIncludeId, resultWhenExcludeId);
            Site.CaptureRequirementIfIsTrue(
                isSame,
                2289,
                  "[In UpdateListItems operation] [In UpdateListItems element] [In updates element] "
                    + "[In Batch element] [In Method element] [New attribute] The server response will be "
                    + "same in both <Field Name=\"ID\"> exists in request or not.");

            // There are 2 list items. Update the first item and delete the last item.
            items.Clear();
            cmds.Clear();

            Dictionary<string, string> updatedItem = new Dictionary<string, string>();
            string newFieldValue = TestSuiteHelper.GenerateRandomString(10);
            updatedItem.Add("ID", updatedListItemIds.First());
            updatedItem.Add(validFieldName, newFieldValue);
            items.Add(updatedItem);
            cmds.Add(MethodCmdEnum.Update);

            Dictionary<string, string> deletedItem = new Dictionary<string, string>();
            deletedItem.Add("ID", updatedListItemIds.Last());
            items.Add(deletedItem);
            cmds.Add(MethodCmdEnum.Delete);

            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

            // Call UpdateListItems operation to update the first item and delete the last item.
            this.listswsAdapter.UpdateListItems(listId, updates);

            // Call GetListItems operation.
            CamlViewFields viewFields = TestSuiteHelper.CreateViewFields(true, new List<string> { validFieldName });
            getListItemsResult = this.listswsAdapter.GetListItems(listId, null, null, viewFields, null, null, null);
            data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

            getListItemIds.Clear();
            string columnNameField = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);
            bool isR2093Verified = false;

            foreach (DataRow row in data.Rows)
            {
                // Get item IDs from GetListItems response.
                getListItemIds.Add(row[columnNameId].ToString());

                // Check if the updated item has been updated successfully.
                if (row[columnNameId].ToString().Equals(updatedListItemIds.First(), StringComparison.OrdinalIgnoreCase))
                {
                    if (row[columnNameField].ToString().Equals(newFieldValue, StringComparison.Ordinal))
                    {
                        isR2093Verified = true;
                    }
                }
            }

            // If the updated item in GetListItems response has been updated successfully, R2093 is captured.
            Site.CaptureRequirementIfIsTrue(
                isR2093Verified,
                2093,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [Update attribute] Updates fields for a specific list item.");

            // If GetListItems response does not contain the deleted item ID, it means the item has been deleted. Then R2094 is captured.
            Site.CaptureRequirementIfIsFalse(
                getListItemIds.Contains(updatedListItemIds.Last()),
                2094,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] [In Batch element] [In Method element] [Delete attribute] Deletes a specific list item.");

            // Verify 'Move' method.
            if (Common.IsRequirementEnabled(20961, this.Site))
            {
                // Create 2 document libraries.
                string sourceListName = TestSuiteHelper.GetUniqueListName();
                string sourceList = TestSuiteHelper.CreateList(sourceListName, Convert.ToInt32(TemplateType.Document_Library));

                string destinationListName = TestSuiteHelper.GetUniqueListName();
                string destinationList = TestSuiteHelper.CreateList(destinationListName, Convert.ToInt32(TemplateType.Document_Library));

                // Get the root folder of destination list.
                GetListResponseGetListResult getListResult = this.listswsAdapter.GetList(destinationList);
                string rootFolder = getListResult.List.RootFolder;

                // Upload a document to source list.
                IMS_LISTSWSSUTControlAdapter sutControlAdapter = this.Site.GetAdapter<IMS_LISTSWSSUTControlAdapter>();
                string absoluteFileUrl = sutControlAdapter.UploadFile(sourceListName);

                Site.Assert.IsTrue(
                        !string.IsNullOrEmpty(absoluteFileUrl),
                        "Upload file to the list {0} should be successful, the file path is [{1}]",
                        sourceListName,
                        absoluteFileUrl);

                // Get the uploaded file name.
                string uploadFileName = absoluteFileUrl.Substring(absoluteFileUrl.LastIndexOf('/'));

                // Get uploaded item ID.
                getListItemsResult = this.listswsAdapter.GetListItems(sourceList, null, null, null, null, null, null);
                data = AdapterHelper.ExtractData(getListItemsResult.listitems.data.Any);

                // Get the value of "ID" and "MoveNewUrl" fields.                
                string fileID = (string)data.Rows[0][columnNameId];
                string moveNewUrl = rootFolder.TrimEnd('/') + "/" + uploadFileName.Trim('/');

                // Construct one insert item with ID.
                items.Clear();
                cmds.Clear();
                Dictionary<string, string> movedItem = new Dictionary<string, string>();
                movedItem.Add("ID", fileID);
                movedItem.Add("FileRef", absoluteFileUrl);
                movedItem.Add("MoveNewUrl", moveNewUrl);
                items.Add(movedItem);
                cmds.Add(MethodCmdEnum.Move);

                // Construct the UpdateListItemsUpdates instance.
                UpdateListItemsUpdates moveUpdates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);

                // Call UpdateListItems method to move file.
                UpdateListItemsResponseUpdateListItemsResult moveResult = null;
                moveResult = this.listswsAdapter.UpdateListItems(sourceList, moveUpdates);

                // Error code "0x00000000" indicates the operation completes successfully. Then R20961 captured.
                Site.CaptureRequirementIfAreEqual<string>(
                    "0x00000000",
                    moveResult.Results[0].ErrorCode,
                    20961,
                    @"Implementation does support this behavior[UpdateListItems.Move]. (Microsoft SharePoint Foundation 2010 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This test case is used to verify UpdateListItems operation in the case that the value of the attribute "OnError" is "Continue" or default.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC91_UpdateListItems_OnErrorContinue()
        {
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Get the random generate unique field name.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listId,
                                                                                updates);

            // Capture R2112 if there are three result elements.
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                result.Results.Count(),
                2112,
                "[UpdateListItemsResponse]If the OnError attribute of the Batch element is set to Continue, "
                    + "then each Method corresponds to a Result.");

            // Call the GetListItems to get the current list items.
            GetListItemsResponseGetListItemsResult listItems = null;
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            DataTable data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // If there are two items existing, the third insert won’t be aborted due to 
            // failure of the second insert operation.
            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 1507,
                                 @"[OnErrorEnum]Continue: Continues running subsequent methods after an error is encountered.");

            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 2113,
                                 "[UpdateListItemsResponse] [If the OnError attribute of the Batch element is set to Continue,] "
                                    + "The outcome of a Method operation MUST not affect the other Method operations.");

            if (Common.IsRequirementEnabled(4009, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<int>(
                                     2,
                                     data.Rows.Count,
                                     4009,
                                     @"[In UpdateListItems][In Bench element][In OnError attribute] Implementation does execute the subsequent methods in the batch [if ""OnError=Continue""] when encountering an error."
                                     + @"(The 2007 Microsoft® Office system/Windows® SharePoint® Services 3.0 and above follow this behavior.)");
            }

            // Remove all the list items in this list.
            string colunmIdName = AdapterHelper.PrefixOws + AdapterHelper.FieldIDName;
            List<string> allItemsId = new List<string>();
            foreach (DataRow row in data.Rows)
            {
                string id = Convert.ToString(row[colunmIdName]);
                allItemsId.Add(id);
            }

            TestSuiteHelper.RemoveListItems(listId, allItemsId, OnErrorEnum.Continue);

            // Construct UpdateListItemsUpdates instance using the default OnError value but not the specified value.
            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items);

            // Call UpdateListItems using the default OnError value.
            result = this.listswsAdapter.UpdateListItems(listId, updates);

            // Call the GetListItems to get the current list items.
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // UpdateListItmes does not send OnError, but if there are still 2 items,
            // the third insert won’t be aborted due to failure of the second insert operation.
            // The behavior is the same on Continue.
            Site.CaptureRequirementIfAreEqual<int>(
                                 2,
                                 data.Rows.Count,
                                 2079,
                                 @"[In UpdateListItems][In Bench element][In OnError attribute] The default is "
                                    + "to continue executing subsequent items.");
        }

        /// <summary>
        ///  This test case is used to verify UpdateListItems operation in the case that the value of attribute "OnError" is "Return.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC92_UpdateListItems_OnErrorReturn()
        {
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Get the random generate unique field name.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();

            // Construct three items: the second one is invalid, the first and the last one are valid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(3);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(3);

            // Add the first item.
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Add the second item.
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Add the third item.
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);
            UpdateListItemsResponseUpdateListItemsResult result = this.listswsAdapter.UpdateListItems(
                                                                                listId,
                                                                                updates);

            // Capture R960 if there are three result elements.
            Site.CaptureRequirementIfAreEqual<int>(
                3,
                result.Results.Count(),
                960,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In UpdateListItemsResult element] "
                    + @"If the OnError attribute of the Batch element is set to Return, "
                    + @"then there MUST be a Result element for each Method operation.");

            // If there are any error, it means that
            // the server cannot set the column with the specified value, then capture requirement R2290.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "0x00000000",
                result.Results[1].ErrorCode,
                2290,
               "[In UpdateListItems operation] [In UpdateListItems element] [In updates element] "
                 + "[In Batch element] [In Method element] [New attribute] [If Additional field "
                 + "references DO NOT set those columns to the value specified,] an error will be returned.");

            if (result.Results[1].ErrorCode == "0x81020014")
            {
                // If the error equals 0x81020014, the second item insert will fail. Then capture requirement 2120.
                // There is no mandatory rule in this document to specify the server MUST return 0x81020014 in this case,
                // So there is no force rule to capture this requirement.
                Site.CaptureRequirement(
                    2120,
                    "[In UpdateListItems operation] [In UpdateListItemsResponse element] "
                        + "[In UpdateListItemsResult element] This [0x81020014] specifies a generic "
                        + "error has been encountered, such as an invalid value being specified for a Field.");
            }

            GetListItemsResponseGetListItemsResult listItems = null;
            listItems = this.listswsAdapter.GetListItems(
                                                    listId,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null,
                                                    null);
            DataTable data = AdapterHelper.ExtractData(listItems.listitems.data.Any);

            // If there is only one item existing, the third insert won’t be aborted due to 
            // failure of the second insert operation.
            Site.CaptureRequirementIfAreEqual<int>(
                                 1,
                                 data.Rows.Count,
                                 1506,
                                "[OnErrorEnum]Return: [If the request includes the simple type with value \"Return\", "
                               + "the protocol server] Stops any more methods from running after the first error is encountered.");

            if (Common.IsRequirementEnabled(4007, this.Site))
            {
                Site.CaptureRequirementIfAreEqual<int>(
                                     1,
                                     data.Rows.Count,
                                     4007,
                                     @"[In UpdateListItems][In Bench element][In OnError attribute] Implementation does result in the entire operation"
                                     + @"being aborted [if ""OnError=Return""],[or whether subsequent methods in the batch should be executed.] when encountering an error."
                                     + @"(The 2007 Microsoft® Office system/Microsoft® Office 2010 suite and above follow this behavior.)");
            }
        }

        /// <summary>
        /// This Test case is used to test UpdateListItems operation when listName is a valid GUID and whether OwsHiddenVersion is specified or not. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC93_UpdateListItems_OwsHiddenVersion()
        {
            string listId = TestSuiteHelper.CreateList();

            // Get the configure value.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct two insert items, one is valid and the other one is invalid.
            List<Dictionary<string, string>> insertItems = new List<Dictionary<string, string>>(2);
            List<MethodCmdEnum> insertCmds = new List<MethodCmdEnum>(2);
            Dictionary<string, string> insertItem1 = new Dictionary<string, string>();
            insertItem1.Add(validFieldName, fieldValue);
            insertItems.Add(insertItem1);
            insertCmds.Add(MethodCmdEnum.New);
            Dictionary<string, string> insertItem2 = new Dictionary<string, string>();
            insertItem2.Add(invalidFieldName, fieldValue);
            insertItems.Add(insertItem2);
            insertCmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(insertCmds, insertItems, OnErrorEnum.Continue);

            // Call UpdateListItems to insert the item using List GUID.
            UpdateListItemsResponseUpdateListItemsResult result = null;
            result = this.listswsAdapter.UpdateListItems(
                                                listId,
                                                updates);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        2,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");

            // Expect the first result is successful but the second one is failed.
            bool isFirstResultSuccess = result.Results[0].ErrorCode == "0x00000000";
            bool isSecondResultFail = result.Results[1].ErrorCode != "0x00000000";

            // Call GetListItems operation by using the List GUID to verify 
            // the operation of inserting item just happen on the specified list.
            GetListItemsResponseGetListItemsResult itemResult = null;
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(false, new List<string> { validFieldName });
            itemResult = this.listswsAdapter.GetListItems(
                                                listId,
                                                null,
                                                null,
                                                fields,
                                                null,
                                                null,
                                                null);

            // If the list specified by the GUID only contains one item
            // and the field value in the z:row element equals the value in the UpdateListItems request,
            // we use the proper list in the UpdateListItems operation.
            // R906 can be captured in this condition.
            DataTable data = AdapterHelper.ExtractData(itemResult.listitems.data.Any);
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);
            bool isUseProperList = data.Rows.Count == 1
                                 && Convert.ToString(data.Rows[0][columnName]) == fieldValue;
            Site.CaptureRequirementIfIsTrue(
                isUseProperList,
                906,
                @"[In UpdateListItems operation] If the specified listName is a valid GUID and "
                + "corresponds to the identification of a list on the site, use that list.");

            // If we make sure use the proper list
            // and operation result contain either successful and failed result
            // then R910 can be captured.
            Site.CaptureRequirementIfIsTrue(
               isUseProperList && isFirstResultSuccess && isSecondResultFail,
               910,
               @"[In UpdateListItems operation] Otherwise [If the specified listName "
                   + "corresponds to the identification or the list title of a list on the site],  "
                   + "the protocol server MUST process the batched operations on the list "
                   + "and return success or failure conditions per operation.");

            // Extract the insert item id and owshiddenversion.
            string id = Convert.ToString(data.Rows[0]["ows_ID"]);
            string owshiddenversion = Convert.ToString(data.Rows[0]["ows_owshiddenversion"]);

            // Construct one update item which contains owshiddenversion equal to the value in the server.
            List<Dictionary<string, string>> updateItems = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> updateCmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> updateItem1 = new Dictionary<string, string>();

            // Set the updated item id.
            updateItem1.Add("ID", id);

            // Set the valid  owsVersion.
            updateItem1.Add("owshiddenversion", owshiddenversion);

            // Set update the new value.
            string newFieldValue = TestSuiteHelper.GenerateRandomString(5);
            updateItem1.Add(validFieldName, newFieldValue);
            updateItems.Add(updateItem1);
            updateCmds.Add(MethodCmdEnum.Update);

            // Call UpdateListItems to insert the item using List GUID.
            updates = TestSuiteHelper.CreateUpdateListItems(updateCmds, updateItems, OnErrorEnum.Continue);
            result = this.listswsAdapter.UpdateListItems(
                                                listId,
                                                updates);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        1,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");
            DataTable updateResultData = AdapterHelper.ExtractData(result.Results[0].Any);

            // If the first successful operation contains one and only one z:row
            // and the field value in the z:row element equals the value in the request,
            // R949 can be captured.
            bool isSuccess = result.Results[0].ErrorCode == "0x00000000"
                                        && updateResultData.Rows.Count == 1
                                        && Convert.ToString(updateResultData.Rows[0][columnName]) == newFieldValue;

            // This requirement is not fully verified, since the error case costs too much to simulate.
            // In this situation, we just verify the successful case.
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                949,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates "
                + "element] [In Batch element] [In Method element] [In Field element] If the "
                + "owshiddenversion value specified by the protocol client is equal to the "
                + "owshiddenversion field's value for the list item on the protocol server, the "
                + "protocol server MUST update the list item or return an error indicating why "
                + "the update failed.");

            // Re-construct one update item which does not contain owshiddenversion.
            updateItems = new List<Dictionary<string, string>>(1);
            updateCmds = new List<MethodCmdEnum>(1);
            updateItem1 = new Dictionary<string, string>();

            // Set the updated item ID.
            updateItem1.Add("ID", id);

            // Set update the new value again.
            newFieldValue = TestSuiteHelper.GenerateRandomString(5);
            updateItem1.Add(validFieldName, newFieldValue);
            updateItems.Add(updateItem1);
            updateCmds.Add(MethodCmdEnum.Update);

            // Call UpdateListItems to insert the item using List GUID.
            updates = TestSuiteHelper.CreateUpdateListItems(updateCmds, updateItems, OnErrorEnum.Continue);
            result = this.listswsAdapter.UpdateListItems(
                                                listId,
                                                updates);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        1,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");
            updateResultData = AdapterHelper.ExtractData(result.Results[0].Any);

            // If the first successful operation contains one and only one z:row
            // and the field value in the z:row element equals the value in the request,
            // R945 can be capture.
            isSuccess = result.Results[0].ErrorCode == "0x00000000"
                                        && updateResultData.Rows.Count == 1
                                        && Convert.ToString(updateResultData.Rows[0][columnName]) == newFieldValue;
            Site.CaptureRequirementIfIsTrue(
                isSuccess,
                945,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates "
                + "element] [In Batch element] [In Method element] [In Field element]If the "
                + "protocol client does not specify the owshiddenversion field reference in the "
                + "UpdateListItems request, then the protocol server MUST overwrite any "
                + "changes in the list item or return an error.");
        }

        /// <summary>
        /// This Test case is used to test UpdateListItems operation in case that the attribute "PreCalc" will be ignored by the server.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC94_UpdateListItems_PreCalcIgnore()
        {
            // Create a list.
            string listId = TestSuiteHelper.CreateList();

            // Read value from configuration file.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);

            // Construct one insert item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(1);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(1);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, TestSuiteHelper.GenerateRandomString(5));
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance and set the PreCalc attribute to true.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Return);
            updates.Batch.PreCalc = "TRUE";
            UpdateListItemsResponseUpdateListItemsResult resultWhenPreCalcTrue = null;
            resultWhenPreCalcTrue = this.listswsAdapter.UpdateListItems(
                                                                listId,
                                                                updates);

            // In the z:row, verify the attribute value which is specified by validFieldName is equal,
            // and for all the other attribute there is no guarantee. So all the other attributes in the z:row 
            // will be ignored except the specified one. 
            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in resultWhenPreCalcTrue.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // Reset the PreCalc to some invalid value.
            updates.Batch.PreCalc = "NotValidBoolean";
            UpdateListItemsResponseUpdateListItemsResult resultWhenPreCalcInvalid = null;
            resultWhenPreCalcInvalid = this.listswsAdapter.UpdateListItems(
                                                                listId,
                                                                updates);

            // In the z:row, we can only make sure the attribute value which 
            // is specified by validFieldName is equal, for all the other attribute there is no 
            // guarantee. So all the other attributes in the z:row will be ignored except the 
            // specified one. 
            foreach (UpdateListItemsResponseUpdateListItemsResultResult result in resultWhenPreCalcInvalid.Results)
            {
                foreach (XmlElement row in result.Any)
                {
                    IgnoreAttributeExcept(
                                    row,
                                    validFieldName);
                }
            }

            // If ignore some field in z:row element and both result are same, capture R2294.
            bool isSame = TestSuiteHelper.DeepCompare(resultWhenPreCalcTrue, resultWhenPreCalcInvalid);
            Site.CaptureRequirementIfIsTrue(
                  isSame,
                  2294,
                  "[In UpdateListItems operation] [In UpdateListItems element] [In updates element]"
                       + "[In Batch element] [PreCalc attribute] The response of server will be same for"
                       + "different values of PreCalc.");
        }

        /// <summary>
        /// This test case is used to test UpdateListItems operation when the listName is not a GUID but a valid list title. 
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC95_UpdateListItems_ValidListTitle()
        {
            string listName = TestSuiteHelper.GetUniqueListName();
            string listId = TestSuiteHelper.CreateList(listName);

            // Get the configure value.
            string invalidFieldName = TestSuiteHelper.GetUniqueFieldName();
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);

            // Construct two insert items: one is valid and the other one is invalid.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>(2);
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>(2);
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);
            Dictionary<string, string> item2 = new Dictionary<string, string>();
            item2.Add(invalidFieldName, fieldValue);
            items.Add(item2);
            cmds.Add(MethodCmdEnum.New);

            // Construct the UpdateListItemsUpdates instance.
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            // Call UpdateListItems using specified list name but not list GUID.
            UpdateListItemsResponseUpdateListItemsResult result = null;
            result = this.listswsAdapter.UpdateListItems(
                                                listName,
                                                updates);

            // The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.
            Site.Assert.AreEqual<int>(
                        2,
                        result.Results.Length,
                        "The number of result element in the  UpdateListItems response MUST be equal to the number of Method in the request.");

            // Expect the first result is successful but the second one is failed.
            bool isFirstResultSuccess = result.Results[0].ErrorCode == "0x00000000";
            bool isSecondResultFail = result.Results[1].ErrorCode != "0x00000000";

            // If the first operation result equal to 0x00000000, capture 
            // Requirement R967
            Site.CaptureRequirementIfIsTrue(
                isFirstResultSuccess,
                967,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In "
                    + "UpdateListItemsResult element] If an operation completes successfully, "
                    + "the value of the ErrorCode MUST be set to 0x00000000.");

            DataTable resultData1 = AdapterHelper.ExtractData(result.Results[0].Any);
            string columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, validFieldName);

            // If the first successful operation contains one and only one z:row
            // and the field value in the z:row element equals the value in the request,
            // the requirement R968 will be able to be captured.
            bool isContainUpdateData = resultData1.Rows.Count == 1
                                            && Convert.ToString(resultData1.Rows[0][columnName]) == fieldValue;
            Site.CaptureRequirementIfIsTrue(
                isContainUpdateData,
                968,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In "
                    + "UpdateListItemsResult element] If the operation succeeded, the z:row "
                    + "element MUST contain the updated data for the list item that is operated on.");

            // Call GetListItems operation by using the List GUID to verify 
            // the operation of inserting item just happens to the specified list.
            CamlViewFields fields = TestSuiteHelper.CreateViewFields(false, new List<string> { validFieldName });
            GetListItemsResponseGetListItemsResult itemResult = null;
            itemResult = this.listswsAdapter.GetListItems(
                                                listId,
                                                null,
                                                null,
                                                fields,
                                                null,
                                                null,
                                                null);

            // If the list specified by the GUID only contains one item
            // and the field value in the z:row element equals the value in the UpdateListItems request,
            // we will use the proper list in the UpdateListItems operation.
            // R907 and R908 can be captured in this condition.
            DataTable data = AdapterHelper.ExtractData(itemResult.listitems.data.Any);
            bool isUseProperList = data.Rows.Count == 1
                                 && Convert.ToString(data.Rows[0][columnName]) == fieldValue;
            Site.CaptureRequirementIfIsTrue(
                isUseProperList,
                907,
                @"[In UpdateListItems operation] If the specified listName is not a valid GUID, "
                    + "check if the listName corresponds to the list title of a list on the site and if "
                    + "so, use that list.");
            Site.CaptureRequirementIfIsTrue(
                isUseProperList,
                908,
                @"[In UpdateListItems operation] If the specified listName does not correspond "
                    + "to the identification of a list on the site, check if the listName corresponds to "
                    + "the list title of a list on the site and if so, use that list.");

            // If we make sure using the proper list
            // and the operation result contains either successful or failed result,
            // R910 can be captured.
            Site.CaptureRequirementIfIsTrue(
               isUseProperList && isFirstResultSuccess && isSecondResultFail,
               910,
               @"[In UpdateListItems operation] Otherwise [If the specified listName "
                   + "corresponds to the identification or the list title of a list on the site],  "
                   + "the protocol server MUST process the batched operations on the list "
                   + "and return success or failure conditions per operation.");

            // If the field value of the list item equals the value in request, R941 can be captured.
            Site.CaptureRequirementIfAreEqual<string>(
                fieldValue,
                Convert.ToString(data.Rows[0][columnName]),
                941,
                @"[In UpdateListItems operation] [In UpdateListItems element] [In updates element] "
                + "[In Batch element] [In Method element] [In Field element]The protocol server "
                + "MUST set the field's value equal to the value of the Field element in the protocol client request, "
                + "if the field is not a special field.");

            #region Capture Requirements R972

            string fieldCounter = Common.GetConfigurationPropertyValue("ListFieldCounter", this.Site);

            // Set the value of Counter as "1".
            string fieldCounterValue = "1";
            cmds = new List<MethodCmdEnum>(1);
            items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item3 = new Dictionary<string, string>();
            item3.Add(fieldCounter, fieldCounterValue);
            items.Add(item3);
            cmds.Add(MethodCmdEnum.New);

            // Call UpdateListItems with valid Guid. Add a counter.
            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            result = this.listswsAdapter.UpdateListItems(listId, updates);

            // Get the value of the "ID" from the server.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, "ID");
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            string id = Convert.ToString(data.Rows[0][columnName]);

            // As description, if update the list item, must add the ID value for the specified list item.
            cmds = new List<MethodCmdEnum>(1);
            items = new List<Dictionary<string, string>>(1);
            Dictionary<string, string> item4 = new Dictionary<string, string>();
            item4.Add("ID", id);

            // Set an invalid value for the counter, Guid is not integer. 
            item4.Add(fieldCounter, Guid.NewGuid().ToString());
            items.Add(item4);
            cmds.Add(MethodCmdEnum.Update);
            updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            result = this.listswsAdapter.UpdateListItems(listId, updates);

            // If the specified item is not be changed, then R972 should be covered.
            columnName = string.Format("{0}{1}", AdapterHelper.PrefixOws, fieldCounter);
            data = AdapterHelper.ExtractData(result.Results[0].Any);
            bool isReturnInZRowElement = Convert.ToString(data.Rows[0][columnName]) == fieldCounterValue;

            Site.Log.Add(
                LogEntryKind.Debug,
                "The actual value: columnName[{0}] for requirement #R972",
                0 == data.Rows.Count ? "NoRows" : data.Rows[0][columnName]);

            Site.CaptureRequirementIfIsTrue(
                isReturnInZRowElement,
                972,
                @"[In UpdateListItems operation] [In UpdateListItemsResponse element] [In UpdateListItemsResult element] "
                + "On failure, the z:row element MUST contain the list item data that exists on the server if the list item specified in the request exists.");

            #endregion
        }

        /// <summary>
        /// This Test case is used to test UpdateListItems operation when listName is an empty string.
        /// </summary>
        [TestCategory("MSLISTSWS"), TestMethod()]
        public void MSLISTSWS_S03_TC96_UpdateListItems_EmptyListName()
        {
            Site.Assume.IsTrue(
                Common.IsRequirementEnabled(30231, this.Site) || Common.IsRequirementEnabled(30232, this.Site),
                @"Test is executed only when R30231Enabled is set to true or R30232Enabled is set to true.");

            // Get the configure value.
            string validFieldName = Common.GetConfigurationPropertyValue("ListFieldText", this.Site);
            string fieldValue = TestSuiteHelper.GenerateRandomString(5);
            bool caughtSoapException = false;

            // Construct one new item.
            List<Dictionary<string, string>> items = new List<Dictionary<string, string>>();
            List<MethodCmdEnum> cmds = new List<MethodCmdEnum>();
            Dictionary<string, string> item1 = new Dictionary<string, string>();
            item1.Add(validFieldName, fieldValue);
            items.Add(item1);
            cmds.Add(MethodCmdEnum.New);
            UpdateListItemsUpdates updates = TestSuiteHelper.CreateUpdateListItems(cmds, items, OnErrorEnum.Continue);

            // Capture R30231 and R30232.
            try
            {
                // Call GetAttachmentCollection with empty listName.
                this.listswsAdapter.UpdateListItems(string.Empty, updates);
                Site.Assert.Fail("UpdateListItems operation must return a SOAP fault if the listName parameter is an empty string.");
            }
            catch (SoapException ex)
            {
                caughtSoapException = true;

                string errorString = TestSuiteHelper.GetErrorString(ex);

                Site.Log.Add(
                        LogEntryKind.Debug,
                        "The actual value: errorString[{0}]",
                        string.IsNullOrEmpty(errorString) ? "NullOrEmpty" : errorString);

                if (Common.IsRequirementEnabled(30231, this.Site))
                {
                    string expectedErrorString = "Object reference not set to an instance of an object.";
                    bool isR30231Verified = expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase);

                    Site.CaptureRequirementIfIsTrue(
                        isR30231Verified,
                        30231,
                        @"[In UpdateListItems operation] Implementation does return the SOAP fault: ""Object reference not set to an instance of an object."", if listName is an empty string. "
                        + @"(<81> Section 3.1.4.31:  Windows SharePoint Services 3.0 returns the SOAP fault: ""Object reference not set to an instance of an object"".)");
                }

                if (Common.IsRequirementEnabled(30232, this.Site))
                {
                    string expectedErrorString = "Cannot access a closed Stream.";
                    bool isR30232Verified = expectedErrorString.Equals(errorString, StringComparison.OrdinalIgnoreCase);

                    Site.CaptureRequirementIfIsTrue(
                        isR30232Verified,
                        30232,
                        @"[In UpdateListItems operation] Implementation does return the SOAP fault: ""Cannot access a closed Stream."", if listName is an empty string. (SharePoint Foundation 2010 and above follow this behavior.)");
                }
            }
            finally
            {
                Site.Assert.IsTrue(caughtSoapException, "A SOAP exception should be thrown by the server when invoke the operation 'UpdateListItems' when listName is an empty string .");
            }
        }

        #endregion

        #endregion

        #region Override properties

        /// <summary>
        /// This method will run before test case executes.
        /// </summary>
        [TestInitialize]
        public void TestCaseInitialize()
        {
            this.listswsAdapter = this.Site.GetAdapter<IMS_LISTSWSAdapter>();

            Common.CheckCommonProperties(this.Site, true);

            // Initialize the TestSuiteHelper.
            TestSuiteHelper.Initialize(this.Site, this.listswsAdapter);
        }

        /// <summary>
        /// This method will run after test case executes.
        /// </summary>
        [TestCleanup]
        public void TestCaseCleanup()
        {
            TestSuiteHelper.CleanUp();
        }

        #endregion

        #region Private helper methods

        /// <summary>
        /// Use this method to verify following requirements in multiple operations.
        /// The names of the attributes containing the list item data in this response correspond 
        /// to the Name attributes in the Field elements of GetList and are prefixed by "ows_". 
        /// </summary>
        /// <param name="data">The xml Response of the operation that is test</param>
        /// <param name="list">The xml Response of GetList operation</param>
        /// <param name="startWithOws">Whether the names have prefix "ows_".</param>
        /// <param name="correspondNames">Whether the names correspond with the fields in the list.</param>
        private static void ValidateOwsPrefixAndAttributeName(DataTable data, GetListResponseGetListResult list, out bool startWithOws, out bool correspondNames)
        {
            startWithOws = true;
            correspondNames = true;

            foreach (DataColumn column in data.Columns)
            {
                startWithOws = startWithOws && column.ColumnName.StartsWith(AdapterHelper.PrefixOws, StringComparison.OrdinalIgnoreCase);
                if (!startWithOws)
                {
                    break;
                }
            }

            if (startWithOws)
            {
                foreach (DataColumn column in data.Columns)
                {
                    FieldDefinition fd = list.List.Fields.Field.FirstOrDefault<FieldDefinition>(field => AdapterHelper.PrefixOws + field.Name == column.ColumnName);
                    if (fd == null)
                    {
                        correspondNames = false;
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// This method is used to set attribute value as NULL from an xml element, except the specified attribute names.
        /// </summary>
        /// <param name="row">The specified xml element instance.</param>
        /// <param name="attrNames">The attribute names whose value will not be set to NULL.</param>
        private static void IgnoreAttributeExcept(XmlElement row, params string[] attrNames)
        {
            // Only interested in the element that is called z:row.
            if (row.LocalName != "row")
            {
                return;
            }

            int i = 0;
            while (i < row.Attributes.Count)
            {
                bool isFind = false;
                XmlAttribute attr = row.Attributes[i];
                foreach (string fieldName in attrNames)
                {
                    if (attr.LocalName == string.Format("{0}{1}", AdapterHelper.PrefixOws, fieldName))
                    {
                        i++;
                        isFind = true;
                        break;
                    }
                }

                if (!isFind)
                {
                    row.Attributes.Remove(attr);
                }
            }
        }

        /// <summary>
        /// This method is used to construct a CamlQueryOptions instance using the following 
        /// specified value :RecurrenceOrderBy – True, 
        ///                 ViewAttributes – Scope= RecursiveAll, 
        ///                 DateInUtc –True,
        ///                 IncludePermissions – True, 
        ///                 IncludeAttachmentUrls – True, 
        ///                 IncludeAttachmentVersion – True, 
        ///                 RecurrencePatternXMLVersion – v3, 
        ///                 ExpandUserField – True, 
        ///                 MeetingInstanceID – -1.
        /// </summary>
        /// <returns>This method returns the CamlQueryOptions instance using the default specified value.</returns>
        private static CamlQueryOptions CreateDefaultCamlQueryOptions()
        {
            CamlQueryOptions options = new CamlQueryOptions();
            options.QueryOptions = new CamlQueryOptionsQueryOptions();
            options.QueryOptions.RecurrenceOrderBy = "True";
            options.QueryOptions.ViewAttributes = new CamlQueryOptionsQueryOptionsViewAttributes();
            options.QueryOptions.ViewAttributes.ScopeSpecified = true;
            options.QueryOptions.ViewAttributes.Scope = EnumViewAttributes.RecursiveAll;
            options.QueryOptions.DateInUtc = "True";
            options.QueryOptions.IncludePermissions = "True";
            options.QueryOptions.IncludeAttachmentUrls = "True";
            options.QueryOptions.IncludeAttachmentVersion = "True";
            options.QueryOptions.RecurrencePatternXMLVersionSpecified = true;
            options.QueryOptions.RecurrencePatternXMLVersion = EnumRecurrencePatternXMLVersion.v3;
            options.QueryOptions.ExpandUserField = "True";
            options.QueryOptions.MeetingInstanceID = "-1";
            return options;
        }

        /// <summary>
        /// Get the change token for the specified list.
        /// </summary>
        /// <param name="listName">Specified the list name.</param>
        /// <returns>Returns the change token.</returns>
        private string GetChangeToken(string listName)
        {
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult result = null;
            try
            {
                result = this.listswsAdapter.GetListItemChangesSinceToken(
                                                            listName,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null,
                                                            null);
            }
            catch (SoapException ex)
            {
                Site.Debug.Fail("Get latest change token for the list {} failed, the reason is {1}", listName, ex.Detail.InnerText);
            }

            return result.listitems.Changes.LastChangeToken;
        }

        #endregion
    }
}