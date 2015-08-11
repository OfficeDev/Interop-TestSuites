namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using System.Collections.Generic;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, sending, deletion and mark of base items on the server.
    /// </summary>
    [TestClass]
    public class S01_ManageBaseItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for base item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC01_CreateGetDeleteItemSuccessfully()
        {
            #region Step 1: Create the item.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);

            // Add Complete FlagType to verify R1045
            if (Common.IsRequirementEnabled(1271, this.Site))
            {
                createdItems[0].Flag = new FlagType();
                createdItems[0].Flag.FlagStatus = FlagStatusType.Complete;
                createdItems[0].Flag.CompleteDateSpecified = true;
                createdItems[0].Flag.CompleteDate = DateTime.UtcNow;
            }

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 createdItemIds.GetLength(0),
                 "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdItemIds.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R292");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R292
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                createItemResponse,
                292,
                @"[In m:CreateItemResponseType Complex Type] The CreateItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            ItemInfoResponseMessageType createItemResponseMessage = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1583");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1583
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                createItemResponseMessage,
                "MS-OXWSCDATA",
                1583,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""CreateItemResponseMessage"" is ""m:ItemInfoResponseMessageType""(section 2.2.4.37) type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1037");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1037
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                createItemResponseMessage,
                "MS-OXWSCDATA",
                1037,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""CreateItemResponseMessage"" with type ""m:ItemInfoResponseMessageType(section 2.2.4.37)"" specifies the response message for the CreateItem operation ([MS-OXWSCORE] section 3.1.4.2).");
            #endregion

            #region Step 2: Get the item.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R380");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R380
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                getItemResponse,
                380,
                @"[In m:GetItemResponseType Complex Type] The GetItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            ItemInfoResponseMessageType getItemResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1585");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1585
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                getItemResponseMessage,
                "MS-OXWSCDATA",
                1585,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""GetItemResponseMessage"" is ""m:ItemInfoResponseMessageType""(section 2.2.4.37) type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1039");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1039
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                getItemResponseMessage,
                "MS-OXWSCDATA",
                1039,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""GetItemResponseMessage"" with type ""m:ItemInfoResponseMessageType"" specifies the response message for the GetItem operation ([MS-OXWSCORE] section 3.1.4.4).");

            // Verify R1045.
            if (Common.IsRequirementEnabled(1271, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1045");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1045
                this.Site.CaptureRequirementIfAreEqual<string>(
                    createdItems[0].Flag.CompleteDate.Date.ToString(),
                    getItemResponseMessage.Items.Items[0].Flag.CompleteDate.Date.ToString(),
                    1045,
                    @"[In t:FlagType Complex Type] CompleteDate: An element of type dateTime that represents the completion date.");
            }
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R299");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R299
            // If the Items element is null in request, the CreateItem operation will fail.
            // The CreateItem operation executed successfully, so this requirement can be verified directly.
            Site.CaptureRequirement(
                299,
                @"[In m:CreateItemType Complex Type] [The element ""Items""] Specifies the collection of items to be created.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R387");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R387
            // If the Items element is null in request, the GetItem operation will fail.
            // The GetItem operation executed successfully, so this requirement can be verified directly.
            Site.CaptureRequirement(
                387,
                @"[In m:GetItemType Complex Type] [The element ""ItemIds""] Specifies the collection of items that a GetItem operation is to get.");

            #region Step 3: Delete the item.
            DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R335");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R335
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse,
                335,
                @"[In m:DeleteItemResponseType Complex Type] The DeleteItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1584");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1584
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse.ResponseMessages.Items[0],
                "MS-OXWSCDATA",
                1584,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""DeleteItemResponseMessage"" is ""m:ResponseMessageType""(section 2.2.4.57) type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1038");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1038
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse.ResponseMessages.Items[0],
                "MS-OXWSCDATA",
                1038,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""DeleteItemResponseMessage"" with type ""m:ResponseMessageType(section 2.2.4.57)"" specifies the response message for the DeleteItem operation ([MS-OXWSCORE] section 3.1.4.3).");

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 4:Get the deleted item
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(getItemIds);

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
                    "Get deleted item should be failed! Expected response code: {0}, actual response code: {1}",
                    ResponseCodeType.ErrorItemNotFound,
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R341");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R341
            // The ItemIds are specified in DeleteItem request in step 3,
            // If the deleted item cannot be gotten in step 4,
            // This requirement can be verified.
            Site.CaptureRequirement(
                341,
                @"[In m:DeleteItemType Complex Type] [The element ""ItemIds""] Specifies the collection of items to be deleted.");
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for base item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC02_CopyItemSuccessfully()
        {
            #region Step 1: Create the item.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Copy the item.
            CopyItemType copyItemRequest = new CopyItemType();
            CopyItemResponseType copyItemResponse = new CopyItemResponseType();

            // Configure ItemIds.
            copyItemRequest.ItemIds = createdItemIds;

            // Configure copying item to draft folder.
            DistinguishedFolderIdType distinguishedFolderIdForCopyItem = new DistinguishedFolderIdType();
            distinguishedFolderIdForCopyItem.Id = DistinguishedFolderIdNameType.drafts;
            copyItemRequest.ToFolderId = new TargetFolderIdType();
            copyItemRequest.ToFolderId.Item = distinguishedFolderIdForCopyItem;

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                copyItemRequest.ReturnNewItemIds = true;
                copyItemRequest.ReturnNewItemIdsSpecified = true;
            }

            copyItemResponse = this.COREAdapter.CopyItem(copyItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 copiedItemIds.GetLength(0),
                 "One copied item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R253");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R253
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                copyItemResponse,
                253,
                @"[In m:CopyItemResponseType Complex Type] The CopyItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1601");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1601
            // The schema is validated, so this requirement can be captured.
            this.Site.CaptureRequirement(
                1601,
                @"[In m:BaseMoveCopyItemType Complex Type] [The element ""ItemIds""] Specifies an array of elements of type BaseItemIdType that specifies a set of items to be copied.");

            ItemInfoResponseMessageType copyItemResponseMessage = copyItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1602");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1602
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                copyItemResponseMessage,
                "MS-OXWSCDATA",
                1602,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""CopyItemResponseMessage"" is ""m:ItemInfoResponseMessageType"" type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1057");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1057
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                copyItemResponseMessage,
                "MS-OXWSCDATA",
                1057,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""CopyItemResponseMessage"" with type ""m:ItemInfoResponseMessageType"" specifies the response message for the CopyItem operation ([MS-OXWSCORE] section 3.1.4.1).");
            #endregion

            #region Step 3: Get the first created item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion

            #region Step 4: Get the second copied item success.

            // The Item properties returned.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1600");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1600
            // If the copied item was got successfully, R1600 can be captured.
            this.Site.CaptureRequirement(
                1600,
                @"[In m:BaseMoveCopyItemType Complex Type] [The element ""ToFolderId""] Specifies an instance of the TargetFolderIdType complex type that specifies the folder to which the items specified by the ItemIds property are to be copied.");

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1604");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1604
                // The copied item was got successfully and the returned item Id is different with the created item Id, so this requirement can be captured.
                bool isVerifiedR1604 = this.IsSchemaValidated && copiedItemIds != createdItemIds;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR1604,
                    1604,
                    @"[In m:BaseMoveCopyItemType Complex Type] [ReturnNewItemIds is] True, indicates the ItemId element is to be returned for new items [when item is copied].");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for base item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC03_MoveItemSuccessfully()
        {
            #region Step 1: Create the item.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the item.

            MoveItemType moveItemRequest = new MoveItemType();
            MoveItemResponseType moveItemResponse = new MoveItemResponseType();

            // Configure ItemIds.
            moveItemRequest.ItemIds = createdItemIds;

            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Configure moving item to inbox folder.
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.inbox;
            moveItemRequest.ToFolderId = new TargetFolderIdType();
            moveItemRequest.ToFolderId.Item = distinguishedFolderId;

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                moveItemRequest.ReturnNewItemIds = true;
                moveItemRequest.ReturnNewItemIdsSpecified = true;
            }

            moveItemResponse = this.COREAdapter.MoveItem(moveItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 movedItemIds.GetLength(0),
                 "One moved item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1230");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1230
                // The ReturnNewItemId was set to true in the MoveItem request and MoveItem operation was executed successfully, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1230,
                    @"[In Appendix C: Product Behavior] Implementation does introduce the ReturnNewItemIds element. (<32> Section 2.2.4.3: The ReturnNewItemIds element was introduced in Exchange 2010 SP2 (Exchange 2013 and above follow this behavior).)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R419");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R419
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                moveItemResponse,
                419,
                @"[In m:MoveItemResponseType Complex Type] The MoveItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R46");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R46
            // If the schema is validated, this requirement can be captured.
            this.Site.CaptureRequirement(
                46,
                @"[In m:BaseMoveCopyItemType Complex Type] [The element ""ItemIds""] Specifies an array of elements of type BaseItemIdType that specifies a set of items to be moved.");

            ItemInfoResponseMessageType moveItemResponseMessage = moveItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1601");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1601
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                moveItemResponseMessage,
                "MS-OXWSCDATA",
                1601,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""MoveItemResponseMessage"" is ""m:ItemInfoResponseMessageType"" type (section 2.2.4.37).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1056");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1056
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                moveItemResponseMessage,
                "MS-OXWSCDATA",
                1056,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""MoveItemResponseMessage"" with type ""m:ItemInfoResponseMessageType(section 2.2.4.37)"" specifies the response message for the MoveItem operation ([MS-OXWSCORE] section 3.1.4.7).");
            #endregion

            #region Step 3: Get the created item failed.
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
                    "Get item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion

            #region Step 4: Get the moved item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1602");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1602
                // The moved item was got successfully and the returned item Id is different with the created item Id, so this requirement can be captured.
                bool isVerifiedR1602 = this.IsSchemaValidated && movedItemIds != createdItemIds;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR1602,
                    1602,
                    @"[In m:BaseMoveCopyItemType Complex Type] [ReturnNewItemIds is] True, indicates the ItemId element is to be returned for new items [when item is moved].");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R45");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R45
            // If the moved item was got successfully, R45 can be captured.
            this.Site.CaptureRequirement(
                45,
                @"[In m:BaseMoveCopyItemType Complex Type] [The element ""ToFolderId""] Specifies an instance of the TargetFolderIdType complex type that specifies the folder to which the items specified by the ItemIds property are to be moved.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for base item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC04_UpdateItemSuccessfully()
        {
            #region Step 1: Create the item.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Update the item, using AppendToItemField element.
            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            AppendToItemFieldType append = new AppendToItemFieldType();
            append.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemBody
            };
            append.Item1 = new ItemType()
            {
                Body = new BodyType()
                {
                    BodyType1 = BodyTypeType.Text,
                    Value = TestSuiteHelper.BodyForBaseItem
                }
            };
            itemChanges[0].Updates[0] = append;

            // Call UpdateItem to update the body of the created item, by using ItemId in CreateItem response.
            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);

            ItemIdType[] updatedItemIds = createdItemIds;

            // One updated item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 updatedItemIds.GetLength(0),
                 "One updated item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 updatedItemIds.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R508");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R508
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponse,
                508,
                @"[In m:UpdateItemResponseType Complex Type] The UpdateItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            UpdateItemResponseMessageType updateItemResponseMessage = updateItemResponse.ResponseMessages.Items[0] as UpdateItemResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R58");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R58
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage,
                58,
                @"[In m:UpdateItemResponseMessageType Complex Type] The UpdateItemResponseMessageType complex type extends the ItemInfoResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.37).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1586");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1586
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage,
                "MS-OXWSCDATA",
                1586,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""UpdateItemResponseMessage"" is ""m:UpdateItemResponseMessageType"" type ([MS-OXWSCORE] section 2.2.4.6).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1040");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1040
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage,
                "MS-OXWSCDATA",
                1040,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""UpdateItemResponseMessage"" with type ""m:UpdateItemResponseMessageType"" specifies the response message for the UpdateItem operation ([MS-OXWSCORE] section 3.1.4.9).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1305");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1305
            // The schema is validated and the conflict result is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage.ConflictResults,
                1305,
                @"[In m:UpdateItemResponseMessageType Complex Type] The type of ConflictResults is t:ConflictResultsType (section 2.2.4.7).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1306");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1306
            // The schema is validated, so this requirement can be captured.
            this.Site.CaptureRequirement(
                1306,
                @"[In t:ConflictResultsType Complex Type] The type of Count is xs:int [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R65");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R65
            // The schema is validated, so this requirement can be captured.
            this.Site.CaptureRequirement(
                65,
                @"[In t:ConflictResultsType Complex Type] [The element ""Count""] Specifies an integer value that indicates the number of conflicts in an UpdateItem operation response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R61");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R61
            // The schema is validated and the conflict result is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage.ConflictResults,
                61,
                @"[In m:UpdateItemResponseMessageType Complex Type] [The element ""ConflictResults""] Specifies the number of conflicts in the result of a single call.");

            #endregion

            #region Step 3: Get the item to check the updates.
            // Call GetItem to get the updated item, by using updatedItemIds in UpdateItem response.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(updatedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            ItemInfoResponseMessageType getItemResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            Site.Assert.AreEqual<BodyTypeType>(
                append.Item1.Body.BodyType1,
                getItemResponseMessage.Items.Items[0].Body.BodyType1,
                string.Format(
                "The value of BodyType1 should be {0}, actual {1}.",
                append.Item1.Body.BodyType1,
                getItemResponseMessage.Items.Items[0].Body.BodyType1));

            Site.Assert.AreEqual<string>(
                append.Item1.Body.Value,
                getItemResponseMessage.Items.Items[0].Body.Value,
                string.Format(
                "The value of Body should be {0}, actual {1}.",
                append.Item1.Body.Value,
                getItemResponseMessage.Items.Items[0].Body.Value));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R555");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R555
            // The value of BodyType1 from response is equal to the value of BodyType1 from the request,
            // and the value of Body from response is equal to the value of Body from request,
            // so this requirement can be captured.
            this.Site.CaptureRequirement(
                555,
                @"[In t:NonEmptyArrayOfItemChangeDescriptionsType Complex Type] [The element ""AppendToItemField""] Specifies data to append to a single property of an item during an UpdateItem operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R561");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R561
            // The value of BodyType1 from response is equal to the value of BodyType1 from the request,
            // and the value of Body from response is equal to the value of Body from request,
            // so this requirement can be captured.
            this.Site.CaptureRequirement(
                561,
                @"[In t:NonEmptyArrayOfItemChangesType Complex Type] [The element ""ItemChange""] Specifies an item identifier and the updates to apply to the item.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R514");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R514
            // Because the value of SavedItemFolderId cannot be compared with the parent folder id from response.
            // So if the updated item can be gotten successfully, this requirement can be captured.
            this.Site.CaptureRequirement(
                514,
                @"[In m:UpdateItemType Complex Type] [The element ""SavedItemFolderId""] Specifies the target folder for saved items.");
            #endregion

            #region Step 4: Update the item, using SetItemField element.
            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();
            itemChanges[0].Item = getItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemSubject
            };
            setItem.Item1 = new ItemType()
            {
                Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForUpdateItem)
            };
            itemChanges[0].Updates[0] = setItem;

            // Call UpdateItem to update the subject of the created item, by using ItemId in CreateItem response.
            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);

            updatedItemIds = Common.GetItemIdsFromInfoResponse(updateItemResponse);

            // One updated item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 updatedItemIds.GetLength(0),
                 "One updated item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 updatedItemIds.GetLength(0));

            #endregion

            #region Step 5: Get the item to check the updates.

            // Call GetItem to get the updated item in the Inbox folder, by using updatedItemIds in UpdateItem response.
            getItemResponse = this.CallGetItemOperation(updatedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItemIds.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            getItemResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R556");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R556
            this.Site.CaptureRequirementIfAreEqual<string>(
                setItem.Item1.Subject,
                getItemResponseMessage.Items.Items[0].Subject,
                556,
                @"[In t:NonEmptyArrayOfItemChangeDescriptionsType Complex Type] [The element ""SetItemField""] Specifies an update to a single property of an item in an UpdateItem operation.");
            #endregion

            #region Step 6: Update the item, using DeleteItemField element.
            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();
            itemChanges[0].Item = getItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            DeleteItemFieldType delField = new DeleteItemFieldType();
            delField.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemBody
            };
            itemChanges[0].Updates[0] = delField;

            // Call UpdateItem to delete the body value of the created item, by using ItemId in CreateItem response.
            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);

            updatedItemIds = Common.GetItemIdsFromInfoResponse(updateItemResponse);

            // One updated item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 updatedItemIds.GetLength(0),
                 "One updated item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 updatedItemIds.GetLength(0));

            #endregion

            #region Step 7: Get the item to check the updates
            // Call GetItem to get the updated item in the Inbox folder, by using updatedItemIds in UpdateItem response.
            getItemResponse = this.CallGetItemOperation(updatedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R557");
        
            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R557
            // The value of Body is null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNull(
                getItemResponseMessage.Items.Items[0].Body.Value,
                557,
                @"[In t:NonEmptyArrayOfItemChangeDescriptionsType Complex Type] [The element ""DeleteItemField"" with type ""t:DeleteItemFieldType""] Specifies an operation to delete a given property from an item during an UpdateItem operation.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for base items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC05_MarkAllItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            #region Step 1: Create two items.
            ItemType[] createdItems = new ItemType[] { new ItemType(), new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                1);
            createdItems[1].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                2);

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 2, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // Two created items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    createdItemIds.GetLength(0),
                    "Two created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    createdItemIds.GetLength(0));
            #endregion

            #region Step 2: Get two items.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));

            #endregion

            #region Step 3: Mark all items as unread, and suppress the receive receipts.
            BaseFolderIdType[] folderIds = new BaseFolderIdType[1];
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.drafts;
            folderIds[0] = distinguishedFolderId;

            // Mark all items in drafts folder as unread, and suppress the receive receipts.
            MarkAllItemsAsReadResponseType markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(false, true, folderIds);

            // Check the operation response.
            Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1054011, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1054011");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1054011
                // The MarkAllItemsAsReadResponseMessage is not null and the schema is validated, so this requirement can be captured.
                this.Site.CaptureRequirementIfIsNotNull(
                    markAllItemsAsReadResponse.ResponseMessages.Items[0],
                    "MS-OXWSCDATA",
                    1054011,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""MarkAllItemsAsReadResponseMessage"" with type ""m:ResponseMessageType"", which specifies the response message for the MarkAllItemsAsRead operation.(Exchange 2013 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1212");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1212
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                markAllItemsAsReadResponse,
                1212,
                @"[In m:MarkAllItemsAsReadResponseType Complex Type] The MarkAllItemsAsReadResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");
            #endregion

            #region Step 4: Get two items and check the updates.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));

            #endregion

            #region Step 5: Mark all items as read, and suppress the receive receipts.
            // Mark all items in drafts folder as read, and suppress the receive receipts.
            markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(true, true, folderIds);

            // Check the operation response.
            Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);

            #endregion

            #region Step 6:Get two items and check the updates
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));

            #endregion

            #region Step 7: Mark all items as unread, and don't suppress the receive receipts.
            // Mark all items in drafts folder as unread, and don't suppress the receive receipts
            markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(false, false, folderIds);

            // Check the operation response.
            Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);

            #endregion

            #region Step 8: Get two items and check the updates.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));

            #endregion

            #region Step 9: Mark all items as read, and don't suppress the receive receipts.
            // Mark all items in drafts folder as read, and don't suppress the receive receipts.
            markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(true, false, folderIds);

            // Check the operation response.
            Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);

            #endregion

            #region Step 10: Get two items and check the updates.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC06_UpdateItemFailed()
        {
            ItemType item = new ItemType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC07_CreateItemFailed()
        {
            #region Step 1: Create the item with invalid item class.
            ItemType[] createdItems = new ItemType[] 
            { 
                new ItemType() 
                { 
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem),

                    // Set an invalid ItemClass to core item.
                    ItemClass = TestSuiteHelper.ItemClassNotNote
                } 
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            #endregion

            // Get ResponseCode from CreateItem operation response.
            ResponseCodeType responseCode = createItemResponse.ResponseMessages.Items[0].ResponseCode;

            // Verify MS-OXWSCDATA_R619.
            this.VerifyErrorObjectTypeChanged(responseCode);
        }

        /// <summary>
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC08_VerifyExtendPropertyType()
        {
            ItemType item = new ItemType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.drafts, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.drafts, item);
        }

        /// <summary>
        /// This test case is intended to validate all required and optional child element of ItemType with the successful response returned by CreateItem and GetItem operations for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC09_VerifyItemWithWithAllElement()
        {
            #region Step 1: Create the item.
            ItemType[] items = new ItemType[1];
            items[0] = this.CreateFullPropertiesItem();

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.inbox, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 createdItemIds.GetLength(0),
                 "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdItemIds.GetLength(0));
            #endregion

            #region Step 2: Get the item.
            GetItemType getItem = new GetItemType();
            GetItemResponseType getItemResponse = new GetItemResponseType();

            // The Item properties returned.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            List<PathToUnindexedFieldType> pathToUnindexedFields = new List<PathToUnindexedFieldType>();
            if (Common.IsRequirementEnabled(1354, this.Site))
            {
                PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemPreview;
                pathToUnindexedFields.Add(pathToUnindexedField);
            }

            if (Common.IsRequirementEnabled(1729, this.Site))
            {
                PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemGroupingAction;
                pathToUnindexedFields.Add(pathToUnindexedField);
            }

            if (Common.IsRequirementEnabled(1731, this.Site))
            {
                PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemTextBody;
                pathToUnindexedFields.Add(pathToUnindexedField);
            }

            if (pathToUnindexedFields.Count > 0)
            {
                getItem.ItemShape.AdditionalProperties = pathToUnindexedFields.ToArray();
            }

            // The item to get.
            getItem.ItemIds = createdItemIds;

            getItemResponse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 getItems.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItems.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);
            ItemIdId itemIdId = this.ITEMIDAdapter.ParseItemId(itemIds[0]);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R63");

            // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R63
            Site.CaptureRequirementIfAreEqual<IdStorageType>(
               IdStorageType.MailboxItemMailboxGuidBased,
               itemIdId.StorageType,
               "MS-OXWSITEMID",
               63,
               @"[In Id Storage Type (byte)] Its [Id Storage Type's] value maps to the following enumeration value.
                    /// <summary>
                    /// Indicates which type of storage is used for the item/folder represented by this Id.
                    /// </summary>
                    internal enum IdStorageType : byte
                    {  
                [        /// <summary>
                        /// The Id represents an item or folder in a mailbox and 
                        /// it contains a primary SMTP address. 
                        /// </summary>
                        MailboxItemSmtpAddressBased = 0,
                        
                        /// <summary>
                        /// The Id represents a folder in a PublicFolder store.
                        /// </summary>
                        PublicFolder = 1,
                        
                        /// <summary>
                        /// The Id represents an item in a PublicFolder store.
                        /// </summary>
                        PublicFolderItem = 2,
                ]      
                        /// <summary>
                        /// The Id represents an item or folder in a mailbox and contains a mailbox GUID.
                        /// </summary>
                        MailboxItemMailboxGuidBased = 3,
                [
                        /// <summary>
                        /// The Id represents a conversation in a mailbox and contains a mailbox GUID.
                        /// </summary>
                        ConversationIdMailboxGuidBased = 4,
                        
                        /// <summary>
                        /// The Id represents (by objectGuid) an object in the Active Directory.
                        /// </summary>
                        ActiveDirectoryObject = 5,]
                }");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R156");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R156
            this.Site.CaptureRequirementIfAreEqual<string>(
                createdItemIds[0].Id,
                getItems[0].ItemId.Id,
                156,
                @"[In t:ItemIdType Complex Type] [The attribute ""Id""] Specifies an item identifier.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R298");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R298
            // The created item was got successfully, so R298 can be captured.
            this.Site.CaptureRequirement(
                298,
                @"[In m:CreateItemType Complex Type] [The element ""SavedItemFolderId""] Specifies the folder in which new items are saved.");
            #endregion

            #region Capture Code

            // Verify the ItemIdType.
            this.VerifyItemIdType(getItems[0].ItemId);

            // Verify the FolderIdType.
            this.VerifyFolderIdType(getItems[0].ParentFolderId);

            // Verify the ItemClassType.
            this.VerifyItemClassType(getItems[0].ItemClass, items[0].ItemClass);

            // Verify the Subject.
            this.VerifySubject(getItems[0].Subject, items[0].Subject);

            // Verify the SensitivityChoicesType.
            this.VerifySensitivityChoicesType(getItems[0].Sensitivity, items[0].Sensitivity);

            // Verify the BodyType.
            this.VerifyBodyType(getItems[0].Body, items[0].Body);

            // Verify the ArrayOfStringsType.
            this.VerifyArrayOfStringsType(getItems[0].Categories, items[0].Categories);

            // Verify the ImportanceChoicesType.
            this.VerifyImportanceChoicesType(getItems[0].ImportanceSpecified, getItems[0].Importance, items[0].Importance);

            // Verify the InReplyTo.
            this.VerifyInReplyTo(getItems[0].InReplyTo, items[0].InReplyTo);

            // Verify the NonEmptyArrayOfResponseObjectsType.
            this.VerifyNonEmptyArrayOfResponseObjectsType(getItems[0].ResponseObjects);

            // Verify the ReminderDueBy.
            this.VerifyReminderDueBy(getItems[0].ReminderDueBySpecified, getItems[0].ReminderDueBy, items[0].ReminderDueBy);

            // Verify the ReminderIsSet.
            this.VerifyReminderIsSet(getItems[0].ReminderIsSetSpecified);

            // Verify the DisplyTo.
            this.VerifyDisplayTo(getItems[0].DisplayTo);

            // Verify the DisplyCc.
            this.VerifyDisplayCc(getItems[0].DisplayCc);

            // Verify the Culture.
            this.VerifyCulture(getItems[0].Culture, items[0].Culture);

            // Verify the LastModifiedName.
            this.VerifyLastModifiedName(getItems[0].LastModifiedName, Common.GetConfigurationPropertyValue("User1Name", this.Site));

            // Verify the EffectiveRightsType.
            this.VerifyEffectiveRightsType(getItems[0].EffectiveRights);

            // Verify the FlagType.
            this.VerifyFlagType(getItems[0].Flag);

            // Verify the Preview.
            this.VerifyPreview(getItems[0].Preview, getItems[0].Body.Value);

            // Verify the GroupingAction.
            this.VerifyGroupingAction(getItems[0].GroupingAction);

            // Verify the TextBody.
            this.VerifyTextBody(getItems[0].TextBody);

            // Verify the InstanceKey.
            this.VerifyInstanceKey(getItems[0].InstanceKey);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1315");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1315
            Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeReceivedSpecified,
                1315,
                @"[In t:ItemType Complex Type] The type of DateTimeReceived is xs:dateTime [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1316");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1316
            Site.CaptureRequirementIfIsTrue(
                getItems[0].SizeSpecified,
                1316,
                @"[In t:ItemType Complex Type] The type of Size is xs:int [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1320");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1320
            Site.CaptureRequirementIfIsTrue(
                getItems[0].IsSubmittedSpecified,
                1320,
                @"[In t:ItemType Complex Type] The type of IsSubmitted is xs:boolean [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1607");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1607
            Site.CaptureRequirementIfIsFalse(
                getItems[0].IsSubmitted,
                1607,
                @"[In t:ItemType Complex Type] otherwise [IsSubmitted is] false, indicates [an item has not been submitted to the Outbox folder].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1321");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1321
            Site.CaptureRequirementIfIsTrue(
                getItems[0].IsDraftSpecified,
                1321,
                @"[In t:ItemType Complex Type] The type of IsDraft is xs:boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1322");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1322
            Site.CaptureRequirementIfIsTrue(
                getItems[0].IsFromMeSpecified,
                1322,
                @"[In t:ItemType Complex Type] The type of IsFromMe is xs:boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1611");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1611
            Site.CaptureRequirementIfIsFalse(
                getItems[0].IsFromMe,
                1611,
                @"[In t:ItemType Complex Type] otherwise [IsFromMe is] false, indicates [a user does not sent an item to himself or herself].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1323");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1323
            Site.CaptureRequirementIfIsTrue(
                getItems[0].IsResendSpecified,
                1323,
                @"[In t:ItemType Complex Type] The type of IsResend is xs:boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1613");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1613
            Site.CaptureRequirementIfIsFalse(
                getItems[0].IsResend,
                1613,
                @"[In t:ItemType Complex Type] otherwise [IsResend is] false, indicates [an item has not previously been sent].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1324");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1324
            Site.CaptureRequirementIfIsTrue(
                getItems[0].IsUnmodifiedSpecified,
                1324,
                @"[In t:ItemType Complex Type] The type of IsUnmodified is xs:boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1326");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1326
            Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeSentSpecified,
                1326,
                @"[In t:ItemType Complex Type] The type of DateTimeSent is xs:dateTime.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1327");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1327
            Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeCreatedSpecified,
                1327,
                @"[In t:ItemType Complex Type] The type of DateTimeCreated is xs:dateTime.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1335");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1335
            Site.CaptureRequirementIfIsTrue(
                getItems[0].HasAttachmentsSpecified,
                1335,
                @"[In t:ItemType Complex Type] The type of HasAttachments is xs:boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1340");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1340
            Site.CaptureRequirementIfIsTrue(
                getItems[0].LastModifiedTimeSpecified,
                1340,
                @"[In t:ItemType Complex Type] The type of LastModifiedTime is xs:dateTime.");

            // Verify the ReminderMinutesBeforeStartType schema.
            this.VerifyReminderMinutesBeforeStartType(getItems[0].ReminderMinutesBeforeStart, items[0].ReminderMinutesBeforeStart);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R105");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R105
            // The LastModifiedTimeSpecified is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                getItems[0].LastModifiedTimeSpecified,
                105,
                @"[In t:ItemType Complex Type] [The element ""LastModifiedTime""] Specifies an instance of the DateTime structure that represents the date and time when an item was last modified.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R78");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R78
            // The DateTimeReceivedSpecified is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeReceivedSpecified,
                78,
                @"[In t:ItemType Complex Type] [The element ""DateTimeReceived""] Specifies the date and time that an item was received in a mailbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R79");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R79
            // The SizeSpecified is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                getItems[0].SizeSpecified,
                79,
                @"[In t:ItemType Complex Type] [The element ""Size""] Specifies an integer value that represents the size of an item, in bytes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R89");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R89
            // The DateTimeSentSpecified is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeSentSpecified,
                89,
                @"[In t:ItemType Complex Type] [The element ""DateTimeSent""] Specifies the date and time when an item in a mailbox was sent.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R90");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R90
            // The DateTimeCreatedSpecified is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                getItems[0].DateTimeCreatedSpecified,
                90,
                @"[In t:ItemType Complex Type] [The element ""DateTimeCreated""] Specifies the date and time when an item in a mailbox was created.");

            if (Common.IsRequirementEnabled(4003, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R4003");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R4003
                // The IsTruncated is set and the item is created successfully, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    4003,
                    @"[In Appendix C: Product Behavior] Implementation does use the attribute ""IsTruncated"" with type ""xs:boolean ([XMLSCHEMA2])"" which specifies whether the body is truncated. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1353, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1353");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1353
                // The RetentionDate is set and the item is created successfully, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1353,
                    @"[In Appendix C: Product Behavior] Implementation does support element ""RetentionDate"" with type ""xs:dateTime"" which specifies the retention date for an item. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the ImportanceChoicesType enumeration for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC10_VerifyItemWithImportanceChoicesTypeEnums()
        {
            // Define the count of enumerations
            int enumCount = 3;
            ImportanceChoicesType[] importanceChoicesTypes = new ImportanceChoicesType[enumCount];

            importanceChoicesTypes[0] = ImportanceChoicesType.High;
            importanceChoicesTypes[1] = ImportanceChoicesType.Low;
            importanceChoicesTypes[2] = ImportanceChoicesType.Normal;

            // Define an item array to store the items got from GetItem operation response.
            // Each item should contain a ImportanceChoicesType value as its element's value
            ItemType[] items = new ItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                ImportanceChoicesType importanceChoicesType = importanceChoicesTypes[i];

                #region Step 1: Create the item.
                ItemType[] createdItems = new ItemType[] { new ItemType() };
                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);
                createdItems[0].Importance = importanceChoicesType;
                createdItems[0].ImportanceSpecified = true;

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Check the operation response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

                ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                // One created item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     createdItemIds.GetLength(0),
                     "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     createdItemIds.GetLength(0));
                #endregion

                #region Step 2: Get the item.
                GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // One item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     getItems.GetLength(0),
                     "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     getItems.GetLength(0));

                Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

                items[i] = getItems[0];

                #endregion
            }

            #region Capture Code

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R198, Expected result:{0}, Actual result:{1}", importanceChoicesTypes[0], items[0].Importance);

            // If the Importance element is present,
            // and the Importance element of items[0] equal to the High, which is the same value of importanceChoicesTypes[0] in request,
            // then this requirement can be verified.
            bool isVerifyR198 = items[0].ImportanceSpecified && items[0].Importance == ImportanceChoicesType.High;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR198,
                198,
                @"[In t:ImportanceChoicesType Simple Type] [The value ""High""] Specifies high importance.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R199, Expected result:{0}, Actual result:{1}", importanceChoicesTypes[1], items[1].Importance);

            // If the Importance element is present,
            // and the Importance element of items[1] equal to the Low, which is the same value of importanceChoicesTypes[1] in request,
            // then this requirement can be verified.
            bool isVerifyR199 = items[1].ImportanceSpecified && items[1].Importance == ImportanceChoicesType.Low;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR199,
                199,
                @"[In t:ImportanceChoicesType Simple Type] [The value ""Low""] Specifies low importance.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R200, Expected result:{0}, Actual result:{1}", importanceChoicesTypes[2], items[2].Importance);

            // If the Importance element is present,
            // and the Importance element of items[2] equal to the Normal, which is the same value of importanceChoicesTypes[2] in request,
            // then this requirement can be verified.
            bool isVerifyR200 = items[2].ImportanceSpecified && items[2].Importance == ImportanceChoicesType.Normal;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR200,
                200,
                @"[In t:ImportanceChoicesType Simple Type] [The value ""Normal""] Specifies normal importance.");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the SensitivityChoicesType enumeration for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC11_VerifyItemWithSensitivityChoicesTypeEnums()
        {
            // Define the count of enumerations
            int enumCount = 4;
            SensitivityChoicesType[] sensitivityChoicesTypes = new SensitivityChoicesType[enumCount];

            sensitivityChoicesTypes[0] = SensitivityChoicesType.Confidential;
            sensitivityChoicesTypes[1] = SensitivityChoicesType.Normal;
            sensitivityChoicesTypes[2] = SensitivityChoicesType.Personal;
            sensitivityChoicesTypes[3] = SensitivityChoicesType.Private;

            // Define an item array to store the items got from GetItem operation response.
            // Each item should contain a SensitivityChoicesType value as its element's value
            ItemType[] items = new ItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                SensitivityChoicesType sensitivityChoicesType = sensitivityChoicesTypes[i];

                #region Step 1: Create the item.
                ItemType[] createdItems = new ItemType[] { new ItemType() };
                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);
                createdItems[0].Sensitivity = sensitivityChoicesType;
                createdItems[0].SensitivitySpecified = true;

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Check the operation response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

                ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                // One created item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     createdItemIds.GetLength(0),
                     "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     createdItemIds.GetLength(0));
                #endregion

                #region Step 2: Get the item.
                // Call the GetItem operation.
                GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // One item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     getItems.GetLength(0),
                     "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     getItems.GetLength(0));

                Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

                items[i] = getItems[0];

                #endregion
            }

            #region Capture Code

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1650, Expected result:{0}, Actual result:{1}", sensitivityChoicesTypes[0], items[0].Sensitivity);

            // If the Sensitivity element is present,
            // and the Sensitivity element of items[0] equal to the Confidential, which is the same value of sensitivityChoicesTypes[0] in request,
            // then this requirement can be verified.
            bool isVerifyR1650 = items[0].SensitivitySpecified && items[0].Sensitivity == SensitivityChoicesType.Confidential;

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1650
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1650,
                1650,
                @"[In t:ItemType Complex Type] The value ""Confidential"" of ""Sensitivity"" specifies the item as confidential.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1651, Expected result:{0}, Actual result:{1}", sensitivityChoicesTypes[1], items[1].Sensitivity);

            // If the Sensitivity element is present,
            // and the Sensitivity element of items[1] equal to the Normal, which is the same value of sensitivityChoicesTypes[1] in request,
            // then this requirement can be verified.
            bool isVerifyR1651 = items[1].SensitivitySpecified && items[1].Sensitivity == SensitivityChoicesType.Normal;

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1651
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1651,
                1651,
                @"[In t:ItemType Complex Type] The value ""Normal"" of ""Sensitivity"" specifies the item as normal.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1652, Expected result:{0}, Actual result:{1}", sensitivityChoicesTypes[2], items[2].Sensitivity);

            // If the Sensitivity element is present,
            // and the Sensitivity element of items[2] equal to the Personal, which is the same value of sensitivityChoicesTypes[2] in request,
            // then this requirement can be verified.
            bool isVerifyR1652 = items[2].SensitivitySpecified && items[2].Sensitivity == SensitivityChoicesType.Personal;

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1652
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1652,
                1652,
                @"[In t:ItemType Complex Type] The value ""Personal"" of ""Sensitivity"" specifies the item as personal.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1653, Expected result:{0}, Actual result:{1}", sensitivityChoicesTypes[3], items[3].Sensitivity);

            // If the Sensitivity element is present,
            // and the Sensitivity element of items[3] equal to the Private, which is the same value of sensitivityChoicesTypes[3] in request,
            // then this requirement can be verified.
            bool isVerifyR1653 = items[3].SensitivitySpecified && items[3].Sensitivity == SensitivityChoicesType.Private;

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1653
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1653,
                1653,
                @"[In t:ItemType Complex Type] The value ""Private"" of ""Sensitivity"" specifies the item as private.");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the BodyTypeType enumeration for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC12_VerifyItemWithBodyEnums()
        {
            // Define the count of enumerations
            int enumCount = 2;
            BodyTypeType[] bodyTypeTypes = new BodyTypeType[enumCount];

            bodyTypeTypes[0] = BodyTypeType.HTML;
            bodyTypeTypes[1] = BodyTypeType.Text;

            // Define an item array to store the items got from GetItem operation response.
            // Each item should contain a BodyTypeType value as its element's value
            ItemType[] items = new ItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                BodyTypeType bodyTypeType = bodyTypeTypes[i];

                #region Step 1: Create the item.
                ItemType[] createdItems = new ItemType[] { new ItemType() };
                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);
                createdItems[0].Body = new BodyType();
                createdItems[0].Body.Value = TestSuiteHelper.BodyForBaseItem;
                createdItems[0].Body.BodyType1 = bodyTypeType;

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Check the operation response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

                ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                // One created item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     createdItemIds.GetLength(0),
                     "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     createdItemIds.GetLength(0));
                #endregion

                #region Step 2: Get the item.
                // Call the GetItem operation.
                GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // One item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     getItems.GetLength(0),
                     "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     getItems.GetLength(0));

                items[i] = getItems[0];

                ItemInfoResponseMessageType getItemResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

                Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1097");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1097
                this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                    bodyTypeType,
                    getItemResponseMessage.Items.Items[0].Body.BodyType1,
                    "MS-OXWSCDATA",
                    1097,
                    @"[In t:BodyType Complex Type] The name ""BodyType"" with type ""t:BodyTypeType"" Specifies the body content and format of an item.");
                #endregion
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1679");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1679
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.HTML,
                items[0].Body.BodyType1,
                1679,
                @"[In t:ItemType Complex Type] The value  ""HTML"" of ""Body"" specifies the item body as HTML content.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1680");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1680
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.Text,
                items[1].Body.BodyType1,
                1680,
                @"[In t:ItemType Complex Type] The value ""Text"" of ""Body"" specifies the item body as text content.");
        }

        /// <summary>
        /// This test case is intended to validate the DateTimePrecision enumeration for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC13_VerifyItemWithDateTimePrecisionEnums()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1256, this.Site), "Exchange 2007 does not support the DateTimePrecision element.");

            // Clear the soap headers.
            this.ClearSoapHeaders();

            string[] dateTimePrecisionTypes = new string[] { "Seconds", "Milliseconds" };

            #region Step 1: Create the item.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            // Define an item array to store the items got from GetItem operation response.
            // Each item should contain a DisposalType value as its element's value
            ItemType[] items = new ItemType[dateTimePrecisionTypes.Length];
            XmlElement[] lastRawResponses = new XmlElement[dateTimePrecisionTypes.Length];
            for (int i = 0; i < dateTimePrecisionTypes.Length; i++)
            {
                DateTimePrecisionType dateTimePrecision = new DateTimePrecisionType();
                dateTimePrecision.Text = new string[1];
                dateTimePrecision.Text[0] = dateTimePrecisionTypes[i];

                #region Step 2: Configure the DateTimePrecision.
                Dictionary<string, object> soapHeaders = new Dictionary<string, object>();
                soapHeaders.Add("DateTimePrecision", dateTimePrecision);
                this.COREAdapter.ConfigureSOAPHeader(soapHeaders);
                #endregion

                #region Step 3: Get the item.
                // Call the GetItem operation.
                GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

                // Get the LastRawResponseXml.
                lastRawResponses[i] = (XmlElement)this.COREAdapter.LastRawResponseXml;

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // One item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     getItems.GetLength(0),
                     "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     getItems.GetLength(0));

                Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

                items[i] = getItems[0];
                #endregion
            }

            if (Common.IsRequirementEnabled(132502, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R132502");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R132502
                // The schema is validated and the DateTimePrecision element is set in request, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    132502,
                    @"[In Appendix C: Product Behavior] Implementation does include the DateTimePrecision element which specifies precision of DateTime values that are returned in responses. This element is optional. <xs:element name=""DateTimePrecision"" type=""t:DateTimePrecisionType"" /> (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1256");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1256
                // The DateTimePrecision SOAP header is set in request and the schema is validated, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1256,
                    @"[In Appendix C: Product Behavior] Implementation does introduce the DateTimePrecision SOAP header. (<79> Section 3.1.4.4.1.1:  The DateTimePrecision SOAP header was introduced in Exchange 2010 SP2 (Exchange 2013 Preview and above follow this behavior).)");
            }

            if (Common.IsRequirementEnabled(1816, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1816");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1816
                // The DateTimePrecision SOAP header is set in request and the schema is validated, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1816,
                    @"[In Appendix C: Product Behavior] Implementation does introduce the DateTimePrecisionType simple type. (<8> Section 2.2.3.3: The DateTimePrecisionType simple type was introduced in Exchange 2010 SP2 (Exchange 2013 and above follow this behavior).)");

                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1670");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1670
                this.Site.CaptureRequirementIfIsTrue(
                    this.IsExpectedDateTimePrecision(lastRawResponses[0], dateTimePrecisionTypes[0]),
                    1670,
                    @"[In tns:GetItemSoapIn Message] The value ""Seconds"" of ""DateTimePrecision"" which specifies that the precision for date/time return values is seconds.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1671");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1671
                this.Site.CaptureRequirementIfIsTrue(
                    this.IsExpectedDateTimePrecision(lastRawResponses[1], dateTimePrecisionTypes[1]),
                    1671,
                    @"[In tns:GetItemSoapIn Message] The value ""Milliseconds"" of ""DateTimePrecision"" which specifies that the precision for date/time return values is milliseconds.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1162");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1162
                // After R1670 or R1671 was verified successfully, R1162 can be verified.
                this.Site.CaptureRequirement(
                    1162,
                    @"[In tns:GetItemSoapIn Message] [The part ""DateTimePrecision""]  Specifies a SOAP header that identifies the resolution of date/time values in responses from the server, either in seconds or in milliseconds.");

                // Clear the soap header.
                this.ClearSoapHeaders();
            }
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC14_GetItemWithItemResponseShapeType()
        {
            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC15_GetItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC16_GetItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC17_GetItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC18_GetItemWithDefaultShapeNamesTypeEnum()
        {
            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC19_GetItemWithBodyTypeResponseTypeEnum()
        {
            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate the DisposalType enumeration for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC20_VerifyItemWithDisposalTypeEnum()
        {
            DisposalType[] disposalTypes = new DisposalType[] { DisposalType.HardDelete, DisposalType.SoftDelete, DisposalType.MoveToDeletedItems };

            // Define an array to store the find item results from FindItem operation response.
            bool[] findInDeleteditems = new bool[disposalTypes.Length];
            for (int i = 0; i < disposalTypes.Length; i++)
            {
                DisposalType disposalType = disposalTypes[i];

                #region Step 1: Create the item.
                ItemType[] createdItems = new ItemType[] { new ItemType() };

                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Check the operation response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

                ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                // One created item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     createdItemIds.GetLength(0),
                     "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     createdItemIds.GetLength(0));
                #endregion

                #region Step 2: Delete the item with different DisposalType.
                DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation(disposalType);

                // Check the operation response.
                Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

                this.ExistItemIds.Remove(createdItemIds[0]);

                #endregion

                #region Step 3: Fail to get the deleted item.
                // Call the GetItem operation.
                GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // No item should be returned
                Site.Assert.AreEqual<int>(
                     0,
                     getItems.GetLength(0),
                     "No item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     0,
                     getItems.GetLength(0));

                #endregion

                #region Step 4: Find the deleted item in deleteditems folder.
                if (disposalType == DisposalType.MoveToDeletedItems)
                {
                    // Find the deleted item in deleteditems folder.
                    ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.deleteditems, createdItems[0].Subject, "User1");
                    findInDeleteditems[i] = findItemIds != null;

                    if (findInDeleteditems[i])
                    {
                        this.CallMoveItemOperation(DistinguishedFolderIdNameType.drafts, findItemIds);
                    }
                }
                else
                {
                    // Items should not be found in the deleteditems folder if the disposal type is HardDelete or SoftDelete.
                    ItemType[] findItems = this.FindItemWithRestriction(DistinguishedFolderIdNameType.deleteditems, createdItems[0].Subject);
                    findInDeleteditems[i] = findItems != null;
                }
                #endregion
            }

            #region Capture Code

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1667");

            // The item[0] is deleted using HardDelete value, and it cannot be gotten in step 3 after DeleteItem operation, also it cannot be found in deleteditems folder, this represent the item is permanently removed.
            Site.CaptureRequirementIfIsFalse(
                findInDeleteditems[0],
                1667,
                @"[In m:DeleteItemType Complex Type] The value ""HardDelete"" of  ""DeleteType"" which specifies that an item or folder is permanently removed from the store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1668");

            // The item[2] is deleted using MoveToDeletedItems value, and it cannot be got in step 3 after DeleteItem operation, also it can be found in deleteditems folder, this represent the item is moved to the Deleted Items folder.
            Site.CaptureRequirementIfIsTrue(
                findInDeleteditems[2],
                1668,
                @"[In m:DeleteItemType Complex Type] The value ""MoveToDeletedItems"" of ""DeleteType"" which specifies that an item or folder is moved to the Deleted Items folder.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate NormalizedBody element of ItemType with the successful response returned by GetItem operations for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC21_GetItemWithNormalizedBodyEnums()
        {
            bool isNormalizedBodySupported = Common.IsRequirementEnabled(1349, this.Site) && Common.IsRequirementEnabled(1683, this.Site);
            Site.Assume.IsTrue(isNormalizedBodySupported, "Exchange 2007 and Exchange 2010 do not support the NormalizedBody element.");

            // Define the count of enumerations.
            int enumCount = 2;
            BodyTypeResponseType[] bodyTypeResponseTypes = new BodyTypeResponseType[enumCount];

            bodyTypeResponseTypes[0] = BodyTypeResponseType.HTML;
            bodyTypeResponseTypes[1] = BodyTypeResponseType.Text;

            // Define an item array to store the items got from GetItem operation response.
            // Each item should contain a BodyTypeResponseType value as its element's value.
            ItemType[] items = new ItemType[enumCount];
            for (int i = 0; i < enumCount; i++)
            {
                BodyTypeResponseType bodyTypeType = bodyTypeResponseTypes[i];

                #region Step 1: Create the item.
                ItemType[] createdItems = new ItemType[] { new ItemType() };
                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Check the operation response.
                Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

                ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

                // One created item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     createdItemIds.GetLength(0),
                     "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     createdItemIds.GetLength(0));
                #endregion

                #region Step 2: Get the item.
                GetItemType getItem = new GetItemType();
                GetItemResponseType getItemResponse = new GetItemResponseType();

                // The Item properties returned
                getItem.ItemShape = new ItemResponseShapeType();
                getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
                getItem.ItemShape.BodyType = bodyTypeType;
                getItem.ItemShape.BodyTypeSpecified = true;

                // The item to get
                getItem.ItemIds = createdItemIds;

                // Set additional properties.
                getItem.ItemShape.AdditionalProperties = new PathToUnindexedFieldType[]
                {
                    new PathToUnindexedFieldType()
                    { 
                        FieldURI = UnindexedFieldURIType.itemNormalizedBody
                    }
                };

                getItemResponse = this.COREAdapter.GetItem(getItem);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

                ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

                // One item should be returned.
                Site.Assert.AreEqual<int>(
                     1,
                     getItems.GetLength(0),
                     "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                     1,
                     getItems.GetLength(0));

                Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

                // Assert the NormalizedBody elements is not null.
                Site.Assert.IsNotNull(getItems[0].NormalizedBody, "The NormalizedBody element of the item should not be null, actual: {0}.", getItems[0].NormalizedBody);

                items[i] = getItems[0];
                #endregion
            }

            #region Step 3: Verify the NormalizedBody element.

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1349");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1349
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.HTML,
                items[0].NormalizedBody.BodyType1,
                1349,
                @"[In Appendix C: Product Behavior] Implementation does support value ""HTML"" of ""NormalizedBody"" which specifies the item body as HTML content. (Exchange 2013 and above follow this behavior.)");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1683");
        
            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1683
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.Text,
                items[1].NormalizedBody.BodyType1,
                1683,
                @"[In Appendix C: Product Behavior] Implementation does support value ""Text"" of ""NormalizedBody"" which specifies the item body as text content. (Exchange 2013 and above follow this behavior.)");
            #endregion
        }
        #endregion
    }
}