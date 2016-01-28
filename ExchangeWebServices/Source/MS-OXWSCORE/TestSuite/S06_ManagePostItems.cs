namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, deletion and mark of post items on the server.
    /// </summary>
    [TestClass]
    public class S06_ManagePostItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for post item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC01_CreateGetDeletePostItemSuccessfully()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for post item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC02_CopyPostItemSuccessfully()
        {
            #region Step 1: Create the post item.
            PostItemType item = new PostItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Copy the post item.
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied post item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied post item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion 

            #region Step 3: Get the first created post item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One post item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One post item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2020");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2020
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Post",
                ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0].ItemClass,
                2020,
                @"[In t:ItemType Complex Type] This value is ""IPM.Post"" for post item.");
            #endregion

            #region Step 4: Get the second copied post item success.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One post item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One post item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for post item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC03_MovePostItemSuccessfully()
        {
            #region Step 1: Create the post item.
            PostItemType item = new PostItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the post item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved post item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved post item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion 

            #region Step 3: Get the created post item failed.
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
                    "Get post item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion 

            #region Step 4: Get the moved post item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One post item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One post item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for post item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC04_UpdatePostItemSuccessfully()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for post items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC05_MarkAllPostItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            PostItemType[] items = new PostItemType[] { new PostItemType(), new PostItemType() };
            this.TestSteps_VerifyMarkAllItemsAsRead(items);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for post item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC06_CreatePostItemFailed()
        {
            #region Step 1: Create the post item with invalid item class.
            PostItemType[] items = new PostItemType[]
            { 
                new PostItemType() 
                { 
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem),

                    // Set an invalid ItemClass to post item.
                    ItemClass = TestSuiteHelper.InvalidItemClass
                }
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, items);
            #endregion

            // Get ResponseCode from CreateItem operation response.
            ResponseCodeType responseCode = createItemResponse.ResponseMessages.Items[0].ResponseCode;

            // Verify MS-OXWSCDATA_R619.
            this.VerifyErrorObjectTypeChanged(responseCode);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for post item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC07_UpdatePostItemFailed()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the relationship among the child elements of PathToExtendedFieldType complex type with the responses returned by CreateItem operation for post item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC08_VerifyExtendPropertyType()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.drafts, item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get and copy the multiple post items with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC09_OperateMultiplePostItemsSuccessfully()
        {
            PostItemType[] items = new PostItemType[] { new PostItemType(), new PostItemType() };
            this.TestSteps_VerifyOperateMultipleItems(items);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC10_GetPostItemWithItemResponseShapeType()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which IncludeMimeContent element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC11_GetPostItemWithIncludeMimeContent()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC12_GetPostItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC13_GetPostItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC14_GetPostItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC15_GetPostItemWithDefaultShapeNamesTypeEnum()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC16_GetPostItemWithBodyTypeResponseTypeEnum()
        {
            PostItemType item = new PostItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate if invalid ItemClass values are set for Post items in the CreateItem request,
        /// an ErrorObjectTypeChanged response code will be returned in the CreateItem response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S06_TC17_CreatePostItemWithInvalidItemClass()
        {
            #region Step 1: Create the Post item with ItemClass set to IPM.Appointment.
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            PostItemType item = new PostItemType();
            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Appointment";
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Post item with ItemClass IPM.Appointment.");
            #endregion

            #region Step 2: Create the Post item with ItemClass set to IPM.Contact.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Contact";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Post item with ItemClass IPM.Contact.");
            #endregion

            #region Step 3: Create the Post item with ItemClass set to IPM.Task.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 3);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Task";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Post item with ItemClass IPM.Task.");
            #endregion

            #region Step 4: Create the Post item with ItemClass set to IPM.DistList.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 4);
            createItemRequest.Items.Items[0].ItemClass = "IPM.DistList";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Post item with ItemClass IPM.DistList.");
            #endregion

            #region Step 5: Create the Post item with ItemClass set to random string.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 5);
            createItemRequest.Items.Items[0].ItemClass = Common.GenerateResourceName(this.Site, "ItemClass");
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Post item with ItemClass is set to a random string.");
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2023");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2023
            this.Site.CaptureRequirement(
                2023,
                @"[In t:ItemType Complex Type] If invalid values are set for these items in the CreateItem request, an ErrorObjectTypeChanged ([MS-OXWSCDATA] section 2.2.5.24) response code will be returned in the CreateItem response.");
        }
        #endregion
    }
}