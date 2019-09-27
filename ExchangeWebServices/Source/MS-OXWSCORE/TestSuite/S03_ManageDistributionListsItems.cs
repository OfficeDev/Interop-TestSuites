namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, deletion and mark of distribution list items on the server.
    /// </summary>
    [TestClass]
    public class S03_ManageDistributionListsItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for distribution list with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC01_CreateGetDeleteDistributionListsItemSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for distribution list with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC02_CopyDistributionListsItemSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create the distribution list type item.
            DistributionListType item = new DistributionListType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Copy the distribution list type item.
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied distribution list type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied distribution list type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion 

            #region Step 3: Get the first created distribution list type item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);            

            // One distribution list type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One distribution list type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2022");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2022
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.DistList",
                ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0].ItemClass,
                2022,
                @"[In t:ItemType Complex Type] This value is ""IPM.DistList"" for distribution list item.");
            #endregion 

            #region Step 4: Get the second copied distribution list type item success.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One distribution list type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One distribution list type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for distribution list with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC03_MoveDistributionListsItemSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create the distribution list type item.
            DistributionListType item = new DistributionListType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the distribution list type item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();
            
            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved distribution list type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved distribution list type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the created distribution list type item failed.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            Site.Assert.AreEqual<int>(
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemResponse.ResponseMessages.Items.GetLength(0));

            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                getItemResponse.ResponseMessages.Items[0].ResponseClass,
                string.Format(
                    "Get distribution list type item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion 

            #region Step 4: Get the moved distribution list type item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One distribution list type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One distribution list type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for distribution list with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC04_UpdateDistributionListsItemSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for distribution lists with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC05_MarkAllDistributionListsItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            DistributionListType[] items = new DistributionListType[] { new DistributionListType(), new DistributionListType() };
            this.TestSteps_VerifyMarkAllItemsAsRead(items);
        }

        /// <summary>
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for distribution list.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC06_VerifyExtendPropertyType()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.contacts, item);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for distribution list.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC07_CreateDistributionListsItemFailed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create the distribution list type item with invalid item class.
            DistributionListType[] createdItems = new DistributionListType[]
            { 
                new DistributionListType() 
                { 
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem),

                    // Set an invalid ItemClass to contact item.
                    ItemClass = TestSuiteHelper.InvalidItemClass
                } 
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.contacts, createdItems);

            #endregion

            // Get ResponseCode from CreateItem operation response.
            ResponseCodeType responseCode = createItemResponse.ResponseMessages.Items[0].ResponseCode;

            // Verify MS-OXWSCDATA_R619.
            this.VerifyErrorObjectTypeChanged(responseCode);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for distribution list.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC08_UpdateDistributionListsItemFailed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get and copy the multiple distribution lists with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC09_OperateMultipleDistributionListsItemsSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType[] items = new DistributionListType[] { new DistributionListType(), new DistributionListType() };
            this.TestSteps_VerifyOperateMultipleItems(items);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC10_GetDistributionListsItemWithItemResponseShapeType()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exist or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC11_GetDistributionListsItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC12_GetDistributionListsItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC13_GetDistributionListsItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC14_GetDistributionListsItemWithDefaultShapeNamesTypeEnum()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC15_GetDistributionListsItemWithBodyTypeResponseTypeEnum()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate if invalid ItemClass values are set for DistributionLists items in the CreateItem request,
        /// an ErrorObjectTypeChanged response code will be returned in the CreateItem response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC16_CreateDistributionListsItemWithInvalidItemClass()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create the DistributionLists item with ItemClass set to IPM.Appointment.
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            DistributionListType item = new DistributionListType();
            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Appointment";
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a DistributionLists item with ItemClass IPM.Appointment.");
            #endregion

            #region Step 2: Create the DistributionLists item with ItemClass set to IPM.Post.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Post";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a DistributionLists item with ItemClass IPM.Post.");
            #endregion

            #region Step 3: Create the DistributionLists item with ItemClass set to IPM.Task.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 3);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Task";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a DistributionLists item with ItemClass IPM.Task.");
            #endregion

            #region Step 4: Create the DistributionLists item with ItemClass set to IPM.Contact.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 4);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Contact";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a DistributionLists item with ItemClass IPM.Contact.");
            #endregion

            #region Step 5: Create the DistributionLists item with ItemClass set to random string.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 5);
            createItemRequest.Items.Items[0].ItemClass = Common.GenerateResourceName(this.Site, "ItemClass");
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a DistributionLists item with ItemClass is set to a random string.");
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2023");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2023
            this.Site.CaptureRequirement(
                2023,
                @"[In t:ItemType Complex Type] If invalid values are set for these items in the CreateItem request, an ErrorObjectTypeChanged ([MS-OXWSCDATA] section 2.2.5.24) response code will be returned in the CreateItem response.");
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which IncludeMimeContent element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S03_TC17_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(23091, this.Site), "The MimeContent element is not applicable for ContactType, TaskType and DistributionListType item when retrieving MIME content in E2010SP3 version below.");

            DistributionListType item = new DistributionListType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean(item);
        }

        #endregion
    }
}