namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, deletion and mark of contact items on the server.
    /// </summary>
    [TestClass]
    public class S02_ManageContactItems : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, GetItem and DeleteItem operations for contact item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC01_CreateGetDeleteContactItemSuccessfully()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful responses returned by CreateItem, CopyItem and GetItem operations for contact item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC02_CopyContactItemSuccessfully()
        {
            #region Step 1: Create the contact item.
            ContactItemType item = new ContactItemType();            
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2:Copy the contact item
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the first created contact item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            ContactItemType[] getItems = Common.GetItemsFromInfoResponse<ContactItemType>(getItemResponse);

            // One contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                getItemIds.GetLength(0),
                "One contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                1,
                getItemIds.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2018");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2018
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Contact",
                ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0].ItemClass,
                2018,
                @"[In t:ItemType Complex Type] This value is ""IPM.Contact"" for contact item.");
            #endregion

            #region Step 4: Get the second copied contact item success.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for contact item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC03_MoveContactItemSuccessfully()
        {
            #region Step 1: Create the contact item.
            ContactItemType item = new ContactItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the contact item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the created contact item failed.
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
                    "Get contact item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion

            #region Step 4: Get the moved contact item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for contact item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC04_UpdateContactItemSuccessfully()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for contact items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC05_MarkAllContactItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            ContactItemType[] items = new ContactItemType[] { new ContactItemType(), new ContactItemType() };
            this.TestSteps_VerifyMarkAllItemsAsRead<ContactItemType>(items);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for contact item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC06_UpdateContactItemFailed()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for contact item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC07_CreateContactItemFailed()
        {
            #region Step 1: Create the contact item with invalid item class.
            ContactItemType[] createdItems = new ContactItemType[]
            { 
                new ContactItemType() 
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
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for contact item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC08_VerifyExtendPropertyType()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.contacts, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.contacts, item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get and copy the multiple contact items with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC09_OperateMultipleContactItemsSuccessfully()
        {
            ContactItemType[] items = new ContactItemType[] { new ContactItemType(), new ContactItemType() };
            this.TestSteps_VerifyOperateMultipleItems(items);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC10_GetContactItemWithItemResponseShapeType()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC11_GetContactItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC12_GetContactItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC13_GetContactItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC14_GetContactItemWithDefaultShapeNamesTypeEnum()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC15_GetContactItemWithBodyTypeResponseTypeEnum()
        {
            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate if invalid ItemClass values are set for contact items in the CreateItem request,
        /// an ErrorObjectTypeChanged response code will be returned in the CreateItem response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S02_TC16_CreateContactItemWithInvalidItemClass()
        {
            #region Step 1: Create the contact item with ItemClass set to IPM.Appointment.
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            ContactItemType item = new ContactItemType();
            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Appointment";
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a contact item with ItemClass IPM.Appointment.");
            #endregion

            #region Step 2: Create the contact item with ItemClass set to IPM.Post.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Post";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a contact item with ItemClass IPM.Post.");
            #endregion

            #region Step 3: Create the contact item with ItemClass set to IPM.Task.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 3);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Task";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a contact item with ItemClass IPM.Task.");
            #endregion

            #region Step 4: Create the contact item with ItemClass set to IPM.DistList.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 4);
            createItemRequest.Items.Items[0].ItemClass = "IPM.DistList";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a contact item with ItemClass IPM.DistList.");
            #endregion

            #region Step 5: Create the contact item with ItemClass set to random string.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 5);
            createItemRequest.Items.Items[0].ItemClass = Common.GenerateResourceName(this.Site, "ItemClass");
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a contact item with ItemClass is set to a random string.");
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
        public void MSOXWSCORE_S02_TC17_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(23091, this.Site), "E2010SP3 version below do not support the MimeContent element for ContactType, TaskType and DistributionListType item when retrieving MIME content.");

            ContactItemType item = new ContactItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean(item);
        }

        #endregion
    }
}