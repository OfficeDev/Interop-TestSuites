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

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2169");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2169
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Success,
                createItemResponse.ResponseMessages.Items[0].ResponseClass,
                2169,
                @"[In tns:CreateItemSoapOut Message] If the request is successful, the CreateItem operation returns a CreateItemResponse element with the ResponseClass attribute of the CreateItemResponseMessage element set to ""Success"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2170");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2170
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.NoError,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2170,
                @"[In tns:CreateItemSoapOut Message] The ResponseCode element of the CreateItemResponseMessage element is set to ""NoError"".");

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

            if (Common.IsRequirementEnabled(102000000, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R102000000");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R102000000
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    createItemResponseMessage.Items.Items[0],
                    typeof(MessageType),
                    "MS-OXWSCDATA",
                    102000000,
                    @"[In Appendix C: Product Behavior] Implementation does return the items of type t:ItemType as a t:MessageType type. (Exchange 2013 and above follow this behavior.)");
            }            
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

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2008");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2008
                // The FlagStatus element is set to Complete, and StartDate and DueDate elements are not set,
                // the item is created and gotten successfully, so this requirement can be captured directly.
                this.Site.CaptureRequirement(
                    2008,
                    @"[In t:FlagType Complex Type] if the FlagStatus element is set to Complete, the StartDate and DueDate elements MUST not be set in the request;");
            }

            if (Common.IsRequirementEnabled(2281, this.Site))
            {
                this.Site.Assert.IsTrue(getItemResponseMessage.Items.Items[0].HasAttachmentsSpecified, "The HasAttachments element should be present.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1621");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1621
                this.Site.CaptureRequirementIfIsFalse(
                    getItemResponseMessage.Items.Items[0].HasAttachments,
                    1621,
                    @"[In t:ItemType Complex Type] otherwise [HasAttachments is] false, indicates [an item does not have at least one attachment].");
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
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""DeleteItemResponseMessage"" is ""m:ResponseMessageType""(section 2.2.4.67) type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1038");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1038
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                deleteItemResponse.ResponseMessages.Items[0],
                "MS-OXWSCDATA",
                1038,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""DeleteItemResponseMessage"" with type ""m:ResponseMessageType(section 2.2.4.67)"" specifies the response message for the DeleteItem operation ([MS-OXWSCORE] section 3.1.4.3).");

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

            ItemType[] getItems = Common.GetItemsFromInfoResponse<ItemType>(getItemResponse);

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


            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2503
            // The schema is validated and the item is not null, so this requirement can be captured.
            if (Common.IsRequirementEnabled(2503, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2503");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2503
                // The schema is validated and the WebClientEditFormQueryStrings is not null, so this requirement can be captured.
                Site.CaptureRequirementIfIsNotNull(
                    getItems[0].WebClientEditFormQueryString,
                    2503,
                    @"[In Appendix C: Product Behavior] Implementation does support the WebClientEditFormQueryString  element which specifies a query string that identifies a edit form accessible by using a Web browser. (<57> Section 2.2.4.24:  Exchange 2010 and above follow this behavior.)");
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
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R158");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R158
            this.Site.CaptureRequirementIfIsTrue(
                updateItemResponseMessage.Items.Items[0].ItemId.Id == createdItemIds[0].Id
                && updateItemResponseMessage.Items.Items[0].ItemId.ChangeKey != createdItemIds[0].ChangeKey,
                158,
                @"[In t:ItemIdType Complex Type] [The attribute ""ChangeKey""] Specifies a change key.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R58");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R58
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                updateItemResponseMessage,
                58,
                @"[In m:UpdateItemResponseMessageType Complex Type] The UpdateItemResponseMessageType complex type extends the ItemInfoResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.41).");

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

            // Verify the TextBody.
            this.VerifyTextBody(getItems[0].TextBody);

            // Verify the InstanceKey.
            this.VerifyInstanceKey(getItems[0].InstanceKey);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2026");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2026
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].DateTimeReceived,
                2026,
                @"[In t:ItemType Complex Type] This element [DateTimeReceived] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2028");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2028
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].Size,
                2028,
                @"[In t:ItemType Complex Type] This element [Size] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2032");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2032
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].DateTimeSent,
                2032,
                @"[In t:ItemType Complex Type] This element [DateTimeSent] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2034");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2034
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].DateTimeReceived,
                2034,
                @"[In t:ItemType Complex Type] This element [DateTimeCreated] can be returned by the server.");
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2036");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2036
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].DisplayTo,
                2036,
                @"[In t:ItemType Complex Type] This element [DisplayTo] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2038");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2038
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].EffectiveRights,
                2038,
                @"[In t:ItemType Complex Type] This element [EffectiveRights] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2040");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2040
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].LastModifiedName,
                2040,
                @"[In t:ItemType Complex Type] This element [LastModifiedName] can be returned by the server.");

            if (Common.IsRequirementEnabled(1348, this.Site))
            {

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2047");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2047
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].InstanceKey,
                2047,
                @"[In t:ItemType Complex Type] This element [InstanceKey] can be returned by the server.");
            }

            if (Common.IsRequirementEnabled(1354, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2051");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2051
                Site.CaptureRequirementIfIsNotNull(
                    getItems[0].Preview,
                    2051,
                    @"[In t:ItemType Complex Type] This element [Preview] can be returned by the server.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2055");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2055
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].GroupingAction,
                2055,
                @"[In t:ItemType Complex Type] This element [GroupingAction] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2059");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2059
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].BlockStatus,
                2059,
                @"[In t:ItemType Complex Type] This element [BlockStatus] can be returned by the server.");

            if (Common.IsRequirementEnabled(1731, this.Site))
            {
            // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2061");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2061
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].TextBody,
                2061,
                @"[In t:ItemType Complex Type] This element [TextBody] can be returned by the server.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2063");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2063
            Site.CaptureRequirementIfIsNotNull(
                getItems[0].IconIndex,
                2063,
                @"[In t:ItemType Complex Type] This element [IconIndex] can be returned by the server.");

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
                @"[In t:ItemType Complex Type] otherwise [IsSubmitted is] false, indicates [an item has not been submitted to the folder].");

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

            if (Common.IsRequirementEnabled(2281, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1314");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1314
                // The Attachments is set and the item is created successfully, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1314,
                    @"[In t:ItemType Complex Type] The type of Attachments is t:NonEmptyArrayOfAttachmentsType ([MS-OXWSCDATA] section 2.2.4.47).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2281");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2281
                // The Attachments is set and the item is created successfully, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    2281,
                    @"[In Appendix C: Product Behavior] Implementation does use the Attachments element which specifies an array of items or files that are attached to an item. (Exchange 2010 SP2 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1620");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1620
                this.Site.CaptureRequirementIfIsTrue(
                    getItems[0].HasAttachmentsSpecified && getItems[0].HasAttachments,
                    1620,
                    @"[In t:ItemType Complex Type] [HasAttachments is] True, indicates an item has at least one attachment.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1229");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1229
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1229,
                    @"[In t:NonEmptyArrayOfAttachmentsType Complex Type] The type [NonEmptyArrayOfAttachmentsType] is defined as follow:
 <xs:complexType name=""NonEmptyArrayOfAttachmentsType"">
  <xs:choice
    minOccurs=""1""
    maxOccurs=""unbounded""
  >
    <xs:element name=""ItemAttachment""
      type=""t:ItemAttachmentType""
     />
    <xs:element name=""FileAttachment""
      type=""t:FileAttachmentType""
     />
  </xs:choice>
</xs:complexType>");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1633");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1633
                // The item was created with an item attachment.
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1633,
                    @"[In t:NonEmptyArrayOfAttachmentsType Complex Type] The element ""ItemAttachment"" is ""t:ItemAttachmentType"" type.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1634");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1634
                // The item was created with an file attachment.
                this.Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    1634,
                    @"[In t:NonEmptyArrayOfAttachmentsType Complex Type] The element ""FileAttachment"" is ""t:FileAttachmentType"" type.");
            }

            if (Common.IsRequirementEnabled(2285, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2285");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2285
                this.Site.CaptureRequirementIfIsTrue(
                    getItems[0].IsAssociatedSpecified,
                    2285,
                    @"[In Appendix C: Product Behavior] Implementation does support the IsAssociated element which specifies a value that indicates whether the item is associated with a folder. (Exchange 2010 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1619");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1619
                this.Site.CaptureRequirementIfIsFalse(
                    getItems[0].IsAssociated,
                    1619,
                    @"[In t:ItemType Complex Type] otherwise [IsAssociated is] false, indicates [the item is associated with a folder].");
            }
            
            if (Common.IsRequirementEnabled(2338, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2338");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2338
                this.Site.CaptureRequirementIfIsNotNull(
                    getItems[0].WebClientReadFormQueryString,
                    2338,
                    @"[In Appendix C: Product Behavior] Implementation does support the WebClientReadFormQueryString element. (Exchange Server 2010 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1342");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1342
                // The WebClientReadFormQueryString is returned from server and pass the schema validation, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1342,
                    @"[In t:ItemType Complex Type] The type of WebClientReadFormQueryString is xs:string.");
            }

            if (Common.IsRequirementEnabled(2283, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2283");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2283
                this.Site.CaptureRequirementIfIsTrue(
                    getItems[0].ReminderNextTimeSpecified,
                    2283,
                    @"[In Appendix C: Product Behavior] Implementation does support the ReminderNextTime element which specifies the date and time for the next reminder. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1727");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1727
                // The ReminderNextTime is returned from server and pass the schema validation, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1727,
                    @"[In t:ItemType Complex Type] The type of ReminderNextTime is xs:dateTime.");
            }

            if (Common.IsRequirementEnabled(2288, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2288");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2288
                this.Site.CaptureRequirementIfIsNotNull(
                    getItems[0].ConversationId,
                    2288,
                    @"[In Appendix C: Product Behavior] Implementation does support the element ""ConversationId"" which specifies the ID of the conversation that an item is part of.. (Exchange 2010 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1344");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1344
                // The ConversationId is returned from server and pass the schema validation, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1344,
                    @"[In t:ItemType Complex Type] The type of ConversationId is t:ItemIdType.");
            }

            if (Common.IsRequirementEnabled(2923, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2923");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2923               
                this.Site.CaptureRequirementIfIsNotNull(
                    getItems[0].SortKey,                   
                    2923,
                    @"[In Appendix C: Product Behavior] Implementation does support the SortKey element which specifies a sort key. (<81> Section 2.2.4.24:  Exchange 2016 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(2932, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2932");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2932               
                this.Site.CaptureRequirementIfIsInstanceOfType(getItems[0].MentionedMe,
                    typeof(bool),                    
                    2932,
                    @"[In Appendix C: Product Behavior] Implementation does support element name MentionedMe with type xs: boolean which Specifies whether the mention applies to the mailbox owner. (<84> Section 2.2.4.24:  Exchange 2016 and above follow this behavior.)");
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
                    @"[In Appendix C: Product Behavior] Implementation does introduce the DateTimePrecision SOAP header. (<109> Section 3.1.4.4.1.1:  The DateTimePrecision SOAP header was introduced in Microsoft Exchange Server 2010 Service Pack 2 (SP2). (Exchange 2013 Preview and above follow this behavior).)");
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
                    @"[In Appendix C: Product Behavior] Implementation does introduce the DateTimePrecisionType simple type. (<79> Section 2.2.5: The DateTimePrecisionType simple type was introduced in Exchange 2010 SP2 (Exchange 2013 and above follow this behavior).)");

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

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1668");

                    // The item[2] is deleted using SoftDelete value, and it cannot be got in step 3 after DeleteItem operation, also it cannot be found in deleteditems folder, this represent the item is moved to the Deleted Items folder.
                    Site.CaptureRequirementIfIsTrue(
                        findInDeleteditems[2],
                        1668,
                        @"[In m:DeleteItemType Complex Type] The value ""MoveToDeletedItems"" of ""DeleteType"" which specifies that an item or folder is moved to the Deleted Items folder.");
                }
                else if (disposalType == DisposalType.SoftDelete)
                {
                    if (Common.IsRequirementEnabled(4000, this.Site))
                    {
                        // Find the deleted item in deleteditems folder.
                        ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.recoverableitemsdeletions, createdItems[0].Subject, "User1");
                        findInDeleteditems[i] = findItemIds != null;

                        if (findInDeleteditems[i])
                        {
                            this.CallMoveItemOperation(DistinguishedFolderIdNameType.recoverableitemsdeletions, findItemIds);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1669");

                            // The item[1] is deleted using SoftDelete value, and it cannot be got in step 3 after DeleteItem operation, also it cannot be found in deleteditems folder, this represent the item is moved to the Recoverable Items folder.
                            Site.CaptureRequirementIfIsTrue(
                                findInDeleteditems[1],
                                1669,
                                @"[In m:DeleteItemType Complex Type] The value ""SoftDelete"" of ""DeleteType"" which specifies that an item or folder is moved to the dumpster if the dumpster is enabled.");
                        }
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

        /// <summary>
        /// This test case is intended to validate element ReturnNewItemIds is ignored by server.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC22_ReturnNewItemIdsIsIgnored()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1230, this.Site), "Exchange 2007 does not support the ReturnNewItemIds element.");

            #region Step 1: Create the item.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIdsForReturnNewItemIdsValueTrue = this.CreateItemWithMinimumElements(item);
            ItemIdType[] createdItemIdsForReturnNewItemIdsValueFalse = this.CreateItemWithMinimumElements(item);
            ItemIdType[] createdItemIdsForNoReturnNewItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Copy the item with ReturnNewItemIds element setting to true.
            CopyItemType copyItemRequest = new CopyItemType();
            CopyItemResponseType copyItemResponse = new CopyItemResponseType();
            copyItemRequest.ItemIds = createdItemIdsForReturnNewItemIdsValueTrue;
            DistinguishedFolderIdType distinguishedFolderIdForCopyItem = new DistinguishedFolderIdType();
            distinguishedFolderIdForCopyItem.Id = DistinguishedFolderIdNameType.drafts;
            copyItemRequest.ToFolderId = new TargetFolderIdType();
            copyItemRequest.ToFolderId.Item = distinguishedFolderIdForCopyItem;
            copyItemRequest.ReturnNewItemIds = true;
            copyItemRequest.ReturnNewItemIdsSpecified = true;
            copyItemResponse = this.COREAdapter.CopyItem(copyItemRequest);
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);
            ItemIdType[] copiedItemIdsWithReturnNewItemIdsValueTrue = Common.GetItemIdsFromInfoResponse(copyItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 copiedItemIdsWithReturnNewItemIdsValueTrue.GetLength(0),
                 "One copied item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIdsWithReturnNewItemIdsValueTrue.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            bool isReturnNewIdsForReturnNewItemIdsTrue = !copiedItemIdsWithReturnNewItemIdsValueTrue[0].ChangeKey.Equals(createdItemIdsForReturnNewItemIdsValueTrue[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !copiedItemIdsWithReturnNewItemIdsValueTrue[0].Id.Equals(createdItemIdsForReturnNewItemIdsValueTrue[0].Id, StringComparison.InvariantCultureIgnoreCase);
            #endregion

            #region Step 3: Copy the item with ReturnNewItemIds element setting to false.
            copyItemRequest.ItemIds = createdItemIdsForReturnNewItemIdsValueFalse;
            copyItemRequest.ReturnNewItemIds = false;
            copyItemRequest.ReturnNewItemIdsSpecified = true;
            copyItemResponse = this.COREAdapter.CopyItem(copyItemRequest);
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);
            ItemIdType[] copiedItemIdsWithReturnNewItemIdsValueFalse = Common.GetItemIdsFromInfoResponse(copyItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 copiedItemIdsWithReturnNewItemIdsValueFalse.GetLength(0),
                 "One copied item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIdsWithReturnNewItemIdsValueFalse.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            bool isReturnNewIdsForReturnNewItemIdsFalse = !copiedItemIdsWithReturnNewItemIdsValueFalse[0].ChangeKey.Equals(createdItemIdsForReturnNewItemIdsValueFalse[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !copiedItemIdsWithReturnNewItemIdsValueFalse[0].Id.Equals(createdItemIdsForReturnNewItemIdsValueFalse[0].Id, StringComparison.InvariantCultureIgnoreCase);
            #endregion

            #region Step 4: Copy the item without ReturnNewItemIds element.
            copyItemRequest.ItemIds = createdItemIdsForNoReturnNewItemIds;
            copyItemRequest.ReturnNewItemIdsSpecified = false;
            copyItemResponse = this.COREAdapter.CopyItem(copyItemRequest);
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);
            ItemIdType[] copiedItemIdsWithoutReturnNewItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 copiedItemIdsWithoutReturnNewItemIds.GetLength(0),
                 "One copied item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIdsWithoutReturnNewItemIds.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            bool isReturnNewIdsForNoReturnNewItemIds = !copiedItemIdsWithoutReturnNewItemIds[0].ChangeKey.Equals(createdItemIdsForNoReturnNewItemIds[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !copiedItemIdsWithoutReturnNewItemIds[0].Id.Equals(createdItemIdsForNoReturnNewItemIds[0].Id, StringComparison.InvariantCultureIgnoreCase);

            Site.Assert.IsTrue(
                isReturnNewIdsForReturnNewItemIdsTrue == isReturnNewIdsForReturnNewItemIdsFalse
                && isReturnNewIdsForReturnNewItemIdsFalse == isReturnNewIdsForNoReturnNewItemIds,
                "New item id should be always returned or not for CopyItem regardless of wheter including ReturnNewItemIds element and the value for it.");
            #endregion

            #region Step 5: Move the item with ReturnNewItemIds element setting to true.
            createdItemIdsForReturnNewItemIdsValueTrue = this.CreateItemWithMinimumElements(item);
            createdItemIdsForReturnNewItemIdsValueFalse = this.CreateItemWithMinimumElements(item);
            createdItemIdsForNoReturnNewItemIds = this.CreateItemWithMinimumElements(item);

            MoveItemType moveItemRequest = new MoveItemType();
            moveItemRequest.ItemIds = createdItemIdsForReturnNewItemIdsValueTrue;
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = DistinguishedFolderIdNameType.inbox;
            moveItemRequest.ToFolderId = new TargetFolderIdType();
            moveItemRequest.ToFolderId.Item = distinguishedFolderId;
            moveItemRequest.ReturnNewItemIds = true;
            moveItemRequest.ReturnNewItemIdsSpecified = true;
            MoveItemResponseType moveItemResponse = this.COREAdapter.MoveItem(moveItemRequest);
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            ItemIdType[] movedItemIdsForReturnNewItemIdsValueTrue = Common.GetItemIdsFromInfoResponse(moveItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 movedItemIdsForReturnNewItemIdsValueTrue.GetLength(0),
                 "One moved item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIdsForReturnNewItemIdsValueTrue.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");
            this.ExistItemIds.Remove(createdItemIdsForReturnNewItemIdsValueTrue[0]);

            isReturnNewIdsForReturnNewItemIdsTrue = !movedItemIdsForReturnNewItemIdsValueTrue[0].ChangeKey.Equals(createdItemIdsForReturnNewItemIdsValueTrue[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !movedItemIdsForReturnNewItemIdsValueTrue[0].Id.Equals(createdItemIdsForReturnNewItemIdsValueTrue[0].Id, StringComparison.InvariantCultureIgnoreCase);
            #endregion

            #region Step 6: Move the item with ReturnNewItemIds element setting to false.
            moveItemRequest.ItemIds = createdItemIdsForReturnNewItemIdsValueFalse;
            moveItemRequest.ReturnNewItemIds = false;
            moveItemRequest.ReturnNewItemIdsSpecified = true;
            moveItemResponse = this.COREAdapter.MoveItem(moveItemRequest);
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            ItemIdType[] movedItemIdsForReturnNewItemIdsValueFalse = Common.GetItemIdsFromInfoResponse(moveItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 movedItemIdsForReturnNewItemIdsValueFalse.GetLength(0),
                 "One moved item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIdsForReturnNewItemIdsValueFalse.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");
            this.ExistItemIds.Remove(createdItemIdsForReturnNewItemIdsValueFalse[0]);

            isReturnNewIdsForReturnNewItemIdsFalse = !movedItemIdsForReturnNewItemIdsValueFalse[0].ChangeKey.Equals(createdItemIdsForReturnNewItemIdsValueFalse[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !movedItemIdsForReturnNewItemIdsValueFalse[0].Id.Equals(createdItemIdsForReturnNewItemIdsValueFalse[0].Id, StringComparison.InvariantCultureIgnoreCase);
            #endregion

            #region Step 7: Move the item without ReturnNewItemIds element.
            moveItemRequest.ItemIds = createdItemIdsForNoReturnNewItemIds;
            moveItemRequest.ReturnNewItemIdsSpecified = false;
            moveItemResponse = this.COREAdapter.MoveItem(moveItemRequest);
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            ItemIdType[] movedItemIdsForNoReturnNewItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);
            Site.Assert.AreEqual<int>(
                 1,
                 movedItemIdsForNoReturnNewItemIds.GetLength(0),
                 "One moved item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIdsForReturnNewItemIdsValueFalse.GetLength(0));
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");
            this.ExistItemIds.Remove(createdItemIdsForNoReturnNewItemIds[0]);

            isReturnNewIdsForNoReturnNewItemIds = !movedItemIdsForNoReturnNewItemIds[0].ChangeKey.Equals(createdItemIdsForNoReturnNewItemIds[0].ChangeKey, StringComparison.InvariantCultureIgnoreCase)
                || !movedItemIdsForNoReturnNewItemIds[0].Id.Equals(createdItemIdsForNoReturnNewItemIds[0].Id, StringComparison.InvariantCultureIgnoreCase);

            Site.Assert.IsTrue(
                isReturnNewIdsForReturnNewItemIdsTrue == isReturnNewIdsForReturnNewItemIdsFalse
                && isReturnNewIdsForReturnNewItemIdsFalse == isReturnNewIdsForNoReturnNewItemIds,
                "New item id should be always returned or not for MoveItem regardless of wheter including ReturnNewItemIds element and the value for it.");

            if (Common.IsRequirementEnabled(1230, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1230");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1230
                // The ReturnNewItemId was set in the MoveItem and CopyItem request and the operations was executed successfully as above, so this requirement can be captured.
                this.Site.CaptureRequirement(
                    1230,
                    @"[In Appendix C: Product Behavior] Implementation does introduce the ReturnNewItemIds element. (<45> Section 2.2.4.16: The ReturnNewItemIds element was introduced in Microsoft Exchange Server 2010 Service Pack 1 (SP1) (Exchange 2010 SP1 and above follow this behavior).)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R47");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R47
                // ReturnNewItemIds is ignored by server, this requirement can be covered.
                this.Site.CaptureRequirement(
                    47,
                    @"[In m:BaseMoveCopyItemType Complex Type] [The element ""ReturnNewItemIds""] Specifies a Boolean return value that indicates whether the ItemId element is to be returned for new items.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate ErrorInvalidPropertySet is returned if WebClientReadFormQueryString is specified in request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC23_WebClientReadFormQueryStringIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2338, this.Site), "Exchange 2007 does not support the WebClientReadFormQueryString element.");

            #region Step 1: Create the item with setting WebClientReadFormQueryString.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].WebClientReadFormQueryString = Common.GenerateResourceName(this.Site, "WebClientReadFormQueryString");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2043");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2043
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2043,
                @"[In t:ItemType Complex Type] This element [WebClientReadFormQueryString] is read-only and may be returned by the server but if specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion

            #region Step 2: Update created item with setting WebClientReadFormQueryString.

            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemWebClientReadFormQueryString
            };
            setItem.Item1 = new ItemType()
            {
                WebClientReadFormQueryString = Common.GenerateResourceName(this.Site, "WebClientReadFormQueryString")
            };
            itemChanges[0].Updates[0] = setItem;

            // Call UpdateItem to update the body of the created item, by using ItemId in CreateItem response.
            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            ItemType[] updateItems = Common.GetItemsFromInfoResponse<ItemType>(updateItemResponse);
                       
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2043");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2043
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2043,
                @"[In t:ItemType Complex Type] This element [WebClientReadFormQueryString] is read-only and may be returned by the server but if specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate IsAssociated in ItemType is set to true if the item is associated with folder.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC24_CreateItemAssociatedWithFolder()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2285, this.Site), "Exchange 2007 does not support the IsAssociated element.");

            #region Create an user configuration object.
            // User configuration objects are items that are associated with folders in a mailbox.
            string userConfiguratioName = Common.GenerateResourceName(this.Site, "UserConfigurationSampleName").Replace("_", string.Empty);
            bool isSuccess = this.USRCFGSUTControlAdapter.CreateUserConfiguration(
                Common.GetConfigurationPropertyValue("User1Name", this.Site),
                Common.GetConfigurationPropertyValue("User1Password", this.Site),
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                userConfiguratioName);
            Site.Assert.IsTrue(isSuccess, "The user configuration object should be created successfully.");
            #endregion

            #region Find the created user configuration object
            FindItemType findRequest = new FindItemType();
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            findRequest.ParentFolderIds = new BaseFolderIdType[1]
            {
                new DistinguishedFolderIdType() { Id = DistinguishedFolderIdNameType.inbox }
            };
            PathToUnindexedFieldType itemSubject = new PathToUnindexedFieldType();
            itemSubject.FieldURI = UnindexedFieldURIType.itemItemClass;
            ContainsExpressionType expressionType = new ContainsExpressionType();
            expressionType.Item = itemSubject;
            expressionType.ContainmentMode = ContainmentModeType.Substring;
            expressionType.ContainmentModeSpecified = true;
            expressionType.ContainmentComparison = ContainmentComparisonType.IgnoreCaseAndNonSpacingCharacters;
            expressionType.ContainmentComparisonSpecified = true;
            expressionType.Constant = new ConstantValueType();
            expressionType.Constant.Value = "IPM.Configuration";

            RestrictionType restriction = new RestrictionType();
            restriction.Item = expressionType;
            findRequest.Restriction = restriction;
            findRequest.Traversal = ItemQueryTraversalType.Associated;

            FindItemResponseType findResponse = this.SRCHAdapter.FindItem(findRequest);
            ItemType[] foundItems = (((FindItemResponseMessageType)findResponse.ResponseMessages.Items[0]).RootFolder.Item as ArrayOfRealItemsType).Items;
            ItemType item = null;
            foreach (ItemType foundItem in foundItems)
            {
                if (foundItem.ItemClass.Contains(userConfiguratioName))
                {
                    item = foundItem;
                    break;
                }
            }

            Site.Assert.IsNotNull(item, "The created user configuration object should be found!");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1618");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1618
            this.Site.CaptureRequirementIfIsTrue(
                item.IsAssociatedSpecified && item.IsAssociated,
                1618,
                @"[In t:ItemType Complex Type] [IsAssociated is] True, indicates the item is associated with a folder.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate UniqueBody element of ItemType with the successful response returned by GetItem operations for base item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC25_GetItemWithUniqueBodyEnums()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2290, this.Site), "Exchange 2007 does not support the UniqueBody element.");

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
                        FieldURI = UnindexedFieldURIType.itemUniqueBody
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
                Site.Assert.IsNotNull(getItems[0].UniqueBody, "The UniqueBody element of the item should not be null, actual: {0}.", getItems[0].UniqueBody);

                items[i] = getItems[0];
                #endregion
            }

            #region Step 3: Verify the UniqueBody element.

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1681");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1681
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.HTML,
                items[0].UniqueBody.BodyType1,
                1681,
                @"[In t:ItemType Complex Type] The value  ""HTML"" of ""UniqueBody"" specifies the item body as HTML content.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1682");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1682
            this.Site.CaptureRequirementIfAreEqual<BodyTypeType>(
                BodyTypeType.Text,
                items[1].UniqueBody.BodyType1,
                1682,
                @"[In t:ItemType Complex Type] The value ""Text"" of ""UniqueBody"" specifies the item body as text content.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2290");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2290
            // The element UniqueBody is returned, so this requirement can be captured.
            this.Site.CaptureRequirement(
                2290,
                @"[In Appendix C: Product Behavior] Implementation does support the element ""UniqueBody"" which specifies the body part that is unique to the conversation that an item is part of. (Exchange 2010 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate if the FlagStatus element is set to Flagged, the CompleteDate element MUST not be set in the request, and the StartDate and DueDate elements MUST be set or unset in pair.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC26_CreateItemWithFlagStatusFlagged()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1271, this.Site), "Exchange 2007 and Exchange 2010 do not support the FlagType complex type.");

            #region Step 1: Create the item, set FlagStatus to Flagged, set element StartDate and DueDate and do not set CompleteDate element.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                1);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Flagged;
            createdItems[0].Flag.StartDateSpecified = true;
            createdItems[0].Flag.StartDate = DateTime.Now;
            createdItems[0].Flag.DueDateSpecified = true;
            createdItems[0].Flag.DueDate = DateTime.Now.AddDays(1);

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);
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
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2002");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2002
            // FlagStatus element is set to Flagged, CompleteDate element is not set, and StartDate and DueDate elements are set,
            // the item is created and gotten successfully, so this requirement can be captured directly.
            this.Site.CaptureRequirement(
                2002,
                @"[In t:FlagType Complex Type] If the FlagStatus element is set to Flagged, the CompleteDate element MUST not be set in the request, and the StartDate and DueDate elements MUST be set or unset in pair;");

            #endregion

            #region Step 3: Create the item, set FlagStatus to Flagged, and do not set StartDate/DueDate/CompleteDate element.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                2);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Flagged;

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);
            #endregion

            #region Step 2: Get the item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

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

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2002");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2002
            // FlagStatus element is set to Flagged, and do not set StartDate/DueDate/CompleteDate element,
            // the item is created and gotten successfully, so this requirement can be captured directly.
            this.Site.CaptureRequirement(
                2002,
                @"[In t:FlagType Complex Type] If the FlagStatus element is set to Flagged, the CompleteDate element MUST not be set in the request, and the StartDate and DueDate elements MUST be set or unset in pair;");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate if the FlagStatus element is set to NotFlagged, the CompleteDate, StartDate, and DueDate elements MUST not be set in the request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC27_CreateItemWithFlagStatusNotFlagged()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1271, this.Site), "Exchange 2007 and Exchange 2010 do not support the FlagType complex type.");

            #region Step 1: Create the item.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.NotFlagged;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);
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
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2005");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2005
            // FlagStatus element is set to NotFlagged, CompleteDate/StartDate/DueDate element is not set,
            // the item is created and gotten successfully, so this requirement can be captured directly.
            this.Site.CaptureRequirement(
                2005,
                @"[In t:FlagType Complex Type] if the FlagStatus element is set to NotFlagged, the CompleteDate, StartDate, and DueDate elements MUST not be set in the request.");

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate ErrorInvalidArgument will be returned if elements in FlagType are not set correctly.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC28_CreateItemWithFlagStatusFailed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1271, this.Site), "Exchange 2007 and Exchange 2010 do not support the FlagType complex type.");

            #region Step 1: Create the item, set FlagStatus to Flagged, and set CompleteDate element.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                1);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Flagged;
            createdItems[0].Flag.CompleteDateSpecified = true;
            createdItems[0].Flag.CompleteDate = DateTime.UtcNow;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set the CompleteDate element when the FlagStatus element is set to Flagged.");
            #endregion

            #region Step 2: Create the item, set FlagStatus to Flagged, and set StartDate but do not set DueDate.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                2);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Flagged;
            createdItems[0].Flag.StartDateSpecified = true;
            createdItems[0].Flag.StartDate = DateTime.Now;

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set the StartDate element but not set DueDate element when the FlagStatus element is set to Flagged.");
            #endregion

            #region Step 3: Create the item, set FlagStatus to Flagged, and set DueDate but do not set StartDate.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                3);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Flagged;
            createdItems[0].Flag.DueDateSpecified = true;
            createdItems[0].Flag.DueDate = DateTime.Now.AddDays(1);

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set the DueDate element but not set StartDate element when the FlagStatus element is set to Flagged.");
            #endregion

            #region Step 4: Create the item, set the FlagStatus to Complete, and set the StartDate and DueDate element.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                4);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.Complete;
            createdItems[0].Flag.CompleteDateSpecified = true;
            createdItems[0].Flag.CompleteDate = DateTime.UtcNow;
            createdItems[0].Flag.StartDateSpecified = true;
            createdItems[0].Flag.StartDate = DateTime.Now;
            createdItems[0].Flag.DueDateSpecified = true;
            createdItems[0].Flag.DueDate = DateTime.Now.AddDays(1);

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set StartDate and DueDate element when the FlagStatus element is set to Complete.");
            #endregion

            #region Step 5: Create the item, set the FlagStatus to NotFlagged, and set the CompleteDate element.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                5);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.NotFlagged;
            createdItems[0].Flag.CompleteDateSpecified = true;
            createdItems[0].Flag.CompleteDate = DateTime.UtcNow;

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set CompleteDate element when the FlagStatus element is set to NotFlagged.");
            #endregion

            #region Step 6: Create the item, set the FlagStatus to NotFlagged, and set the StartDate and DueDate element.
            createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                6);

            createdItems[0].Flag = new FlagType();
            createdItems[0].Flag.FlagStatus = FlagStatusType.NotFlagged;
            createdItems[0].Flag.StartDateSpecified = true;
            createdItems[0].Flag.StartDate = DateTime.Now;
            createdItems[0].Flag.DueDateSpecified = true;
            createdItems[0].Flag.DueDate = DateTime.Now.AddDays(1);

            createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidArgument,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorInvalidArgument should be returned if set StartDate and DueDate element when the FlagStatus element is set to NotFlagged.");
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2010");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2010
            // The requirement has been verified by above steps.
            this.Site.CaptureRequirement(
                2010,
                @"[In t:FlagType Complex Type] Otherwise [If the FlagStatus element is set to Flagged, the CompleteDate element MUST not be set in the request, and the StartDate and DueDate elements MUST be set or unset in pair; if the FlagStatus element is set to Complete, the StartDate and DueDate elements MUST not be set in the request; if the FlagStatus element is set to NotFlagged, the CompleteDate, StartDate, and DueDate elements MUST not be set in the request.], ErrorInvalidArgument ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
        }

        /// <summary>
        /// This test case is intended to validate the element ItemId is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC29_ItemIdIsReadOnly()
        {
            #region Create an item with setting ItemId
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].ItemId = new ItemIdType();
            createdItems[0].ItemId.Id = Common.GenerateResourceName(this.Site, "Id");
            createdItems[0].ItemId.ChangeKey = Common.GenerateResourceName(this.Site, "ChangeKey");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2014");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2014
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2014,
                @"[In t:ItemType Complex Type] This element [ItemId] is read-only.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2171");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2171
            this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                ResponseClassType.Error,
                createItemResponse.ResponseMessages.Items[0].ResponseClass,
                2171,
                @"[In tns:CreateItemSoapOut Message] If the request is unsuccessful, the CreateItem operation returns a CreateItemResponse element with the ResponseClass attribute of the CreateItemResponseMessage element set to ""Error"".");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2172");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2172
            // MS-OXWSCORE_R2014 is captured, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                2172,
                @"[In tns:CreateItemSoapOut Message] The ResponseCode element of the CreateItemResponseMessage element is set to a value of the ResponseCodeType simple type, as specified in [MS-OXWSCDATA] section 2.2.5.24.");
            #endregion

            #region Update an item with setting ItemId
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemItemId
            };
            setItem.Item1 = new ItemType()
            {
                ItemId = new ItemIdType()
                {
                    Id = Common.GenerateResourceName(this.Site, "Id"),
                    ChangeKey = Common.GenerateResourceName(this.Site, "ChangeKey")
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2347");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2347
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2347,
                @"[In t:ItemType Complex Type] but if [ItemId] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element ParentFolderId is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC30_ParentFolderIdIsReadOnly()
        {
            #region Create an item with setting ParentFolderId
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].ParentFolderId = new FolderIdType();
            createdItems[0].ParentFolderId.Id = Common.GenerateResourceName(this.Site, "Id");
            createdItems[0].ParentFolderId.ChangeKey = Common.GenerateResourceName(this.Site, "ChangeKey");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2016");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2016
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2016,
                @"[In t:ItemType Complex Type] This element [ParentFolderId] is read-only.");
            #endregion

            #region Update an item with setting ParentFolderId
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemParentFolderId
            };
            setItem.Item1 = new ItemType()
            {
                ParentFolderId = new FolderIdType
                {
                    Id = Common.GenerateResourceName(this.Site, "Id"),
                    ChangeKey = Common.GenerateResourceName(this.Site, "ChangeKey")
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2348");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2348
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2348,
                @"[In t:ItemType Complex Type] but if [ParentFolderId] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element DateTimeReceived is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC31_DateTimeReceivedIsReadOnly()
        {
            #region Create an item with setting DateTimeReceived
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].DateTimeReceivedSpecified = true;
            createdItems[0].DateTimeReceived = DateTime.UtcNow;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2025");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2025
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2025,
                @"[In t:ItemType Complex Type] This element [DateTimeReceived] is read-only.");
            #endregion

            #region Update an item with setting DateTimeReceived
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemDateTimeReceived
            };
            setItem.Item1 = new ItemType()
            {
                DateTimeReceivedSpecified = true,
                DateTimeReceived = DateTime.UtcNow
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2349");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2349
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2349,
                @"[In t:ItemType Complex Type] but if [DateTimeReceived] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element Size is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC32_SizeIsReadOnly()
        {
            #region Create an item with setting Size
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].SizeSpecified = true;
            createdItems[0].Size = 10;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2027");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2027
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2027,
                @"[In t:ItemType Complex Type] This element [Size] is read-only.");
            #endregion

            #region Update an item with setting Size
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemSize
            };
            setItem.Item1 = new ItemType()
            {
                Size = 10,
                SizeSpecified = true
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2272");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2272
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2272,
                @"[In t:ItemType Complex Type] but if [Size] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element DateTimeSent is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC33_DateTimeSentIsReadOnly()
        {
            #region Create an item with setting DateTimeSent
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].DateTimeSentSpecified = true;
            createdItems[0].DateTimeSent = DateTime.UtcNow;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2031");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2031
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2031,
                @"[In t:ItemType Complex Type] This element [DateTimeSent] is read-only.");
            #endregion

            #region Update an item with setting DateTimeSent
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemDateTimeSent
            };
            setItem.Item1 = new ItemType()
            {
                DateTimeSent = DateTime.UtcNow,
                DateTimeSentSpecified = true
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2273");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2273
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2273,
                @"[In t:ItemType Complex Type] but if [DateTimeSent] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element DateTimeCreated is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC34_DateTimeCreatedIsReadOnly()
        {
            #region Create an item with setting DateTimeCreated
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].DateTimeCreatedSpecified = true;
            createdItems[0].DateTimeCreated = DateTime.UtcNow;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2033");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2033
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2033,
                @"[In t:ItemType Complex Type] This element [DateTimeCreated] is read-only.");
            #endregion

            #region Update an item with setting DateTimeCreated
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemDateTimeCreated
            };
            setItem.Item1 = new ItemType()
            {
                DateTimeCreatedSpecified = true,
                DateTimeCreated = DateTime.UtcNow
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2274");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2274
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2274,
                @"[In t:ItemType Complex Type] but if [DateTimeCreated] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element DisplayTo is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC35_DisplayToIsReadOnly()
        {
            #region Create an item with setting DisplayTo
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].DisplayTo = Common.GetConfigurationPropertyValue("User1Name", this.Site);

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2035");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2035
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2035,
                @"[In t:ItemType Complex Type] This element [DisplayTo] is read-only.");
            #endregion

            #region Update an item with setting DisplayTo
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemDisplayTo
            };
            setItem.Item1 = new ItemType()
            {
                DisplayTo = Common.GetConfigurationPropertyValue("User1Name", this.Site)
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2275");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2275
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2275,
                @"[In t:ItemType Complex Type] but if [DisplayTo] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element EffectiveRights is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC36_EffectiveRightsIsReadOnly()
        {
            #region Create an item with setting EffectiveRights
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].EffectiveRights = new EffectiveRightsType();
            createdItems[0].EffectiveRights.CreateAssociated = true;
            createdItems[0].EffectiveRights.CreateContents = true;
            createdItems[0].EffectiveRights.CreateHierarchy = true;
            createdItems[0].EffectiveRights.Delete = true;
            createdItems[0].EffectiveRights.Modify = true;
            createdItems[0].EffectiveRights.Read = true;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2037");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2037
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2037,
                @"[In t:ItemType Complex Type] This element [EffectiveRights] is read-only.");
            #endregion

            #region Update an item with setting EffectiveRights
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemEffectiveRights
            };
            setItem.Item1 = new ItemType()
            {
                EffectiveRights = new EffectiveRightsType
                {
                    CreateAssociated = true,
                    CreateContents = true,
                    CreateHierarchy = true,
                    Delete = true,
                    Modify = true,
                    Read = true
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2276");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2276
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2276,
                @"[In t:ItemType Complex Type] but if [EffectiveRights] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element LastModifiedName is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC37_LastModifiedNameIsReadOnly()
        {
            #region Create an item with setting LastModifiedName
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].LastModifiedName = Common.GenerateResourceName(this.Site, "LastModifiedName");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2039");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2039
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2039,
                @"[In t:ItemType Complex Type] This element [LastModifiedName] is read-only.");
            #endregion

            #region Update an item with setting LastModifiedName
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemLastModifiedName
            };
            setItem.Item1 = new ItemType()
            {
                LastModifiedName = Common.GenerateResourceName(this.Site, "LastModifiedName")
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2277");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2277
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2277,
                @"[In t:ItemType Complex Type] but if [LastModifiedName] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element InstanceKey is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC38_InstanceKeyIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1348, this.Site), "Exchange 2007 and Exchange 2010 do not support the InstanceKey element.");

            #region Create an item with setting InstanceKey
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].InstanceKey = Convert.FromBase64String("AQAAAAAAAQ0BAAAAAAAFKAAAAAA=");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2046");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2046
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2046,
                @"[In t:ItemType Complex Type] This element [InstanceKey] is read-only.");
            #endregion

            #region Update an item with setting InstanceKey
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemInstanceKey
            };
            setItem.Item1 = new ItemType()
            {
                InstanceKey = Convert.FromBase64String("AQAAAAAAAQ0BAAAAAAAFKAAAAAA=")
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2350");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2350
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2350,
                @"[In t:ItemType Complex Type]but if [InstanceKey] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element NormalizedBody is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC39_NormalizedBodyIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1349, this.Site), "Exchange 2007 and Exchange 2010 do not support the NormalizedBody element.");

            #region Create an item with setting NormalizedBody
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].NormalizedBody = new BodyType();
            createdItems[0].NormalizedBody.BodyType1 = BodyTypeType.HTML;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2048");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2048
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2048,
                @"[In t:ItemType Complex Type] This element [NormalizedBody] is read-only.");
            #endregion

            #region Update an item with setting NormalizedBody
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemNormalizedBody
            };
            setItem.Item1 = new ItemType()
            {
                NormalizedBody = new BodyType
                {
                    BodyType1 = BodyTypeType.HTML
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2351");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2351
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2351,
                @"[In t:ItemType Complex Type] but if [NormalizedBody] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element Preview is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC40_PreviewIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1354, this.Site), "Exchange 2007 and Exchange 2010 do not support the Preview element.");

            #region Create an item with setting Preview
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].Preview = Common.GenerateResourceName(this.Site, "Preview");

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2050");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2050
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2050,
                @"[In t:ItemType Complex Type] This element [Preview] is read-only.");
            #endregion

            #region Update an item with setting Preview
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemPreview
            };
            setItem.Item1 = new ItemType()
            {
                Preview = Common.GenerateResourceName(this.Site, "Preview")
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2352");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2352
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2352,
                @"[In t:ItemType Complex Type] but if [Preview] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element TextBody is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC41_TextBodyIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1731, this.Site), "Exchange 2007 and Exchange 2010 do not support the TextBody element.");

            #region Create an item with setting TextBody
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].TextBody = new BodyType();
            createdItems[0].TextBody.BodyType1 = BodyTypeType.HTML;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2060");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2060
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2060,
                @"[In t:ItemType Complex Type] This element [TextBody] is read-only.");
            #endregion

            #region Update an item with setting TextBody
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemTextBody
            };
            setItem.Item1 = new ItemType()
            {
                TextBody = new BodyType
                {
                    BodyType1 = BodyTypeType.HTML
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2356");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2356
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2356,
                @"[In t:ItemType Complex Type] but if [TextBody] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element IconIndex is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC42_IconIndexIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1917, this.Site), "Exchange 2007 and Exchange 2010 do not support the IconIndex element.");

            #region Create an item with setting IconIndex
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].IconIndexSpecified = true;
            createdItems[0].IconIndex = IconIndexType.TaskRecur;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2062");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2062
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2062,
                @"[In t:ItemType Complex Type] This element [IconIndex] is read-only.");
            #endregion

            #region Update an item with setting IconIndex
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemIconIndex
            };
            setItem.Item1 = new ItemType()
            {
                IconIndex = IconIndexType.TaskRecur,
                IconIndexSpecified = true
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2358");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2358
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2358,
                @"[In t:ItemType Complex Type] but if [IconIndex] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the element RightsManagementLicenseData is read-only.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC43_RightsManagementLicenseDataIsReadOnly()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1355, this.Site), "Exchange 2007 and Exchange 2010 do not support the RightsManagementLicenseData element.");

            #region Create an item with setting RightsManagementLicenseData
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].RightsManagementLicenseData = new RightsManagementLicenseDataType();
            createdItems[0].RightsManagementLicenseData.EditAllowedSpecified = true;
            createdItems[0].RightsManagementLicenseData.EditAllowed = true;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2052");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2052
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                2052,
                @"[In t:ItemType Complex Type] This element [RightsManagementLicenseData] is read-only.");
            #endregion

            #region Update an item with setting RightsManagementLicenseData
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemRightsManagementLicenseData
            };
            setItem.Item1 = new ItemType()
            {
                RightsManagementLicenseData = new RightsManagementLicenseDataType
                {
                    EditAllowed = true,
                    EditAllowedSpecified = true
                }
            };
            itemChanges[0].Updates[0] = setItem;

            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2353");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2353
            this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorInvalidPropertySet,
                updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                2353,
                @"[In t:ItemType Complex Type] but if [RightsManagementLicenseData] specified in a CreateItem or UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1355");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1355
            // Server handles the element RightsManagementLicenseData and returns ErrorInvalidPropertySet, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                1355,
                @"[In Appendix C: Product Behavior] Implementation does support element ""RightsManagementLicenseData"" with type ""t:RightsManagementLicenseDataType (section 2.2.4.37)"" which specifies rights management license data. (Exchange 2013 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which FilterHtmlContent element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC44_GetItemWithFilterHtmlContent()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2119413, this.Site), "Exchange 2007 do not support the FilterHtmlContent element.");

            ItemType item = new ItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_FilterHtmlContentBoolean(item);
        }

        /// <summary>
        /// This test case is intended to validate ErrorInvalidPropertySet is returned if WebClientEditFormQueryString is specified in request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC45_VerifyStoreEntryIdIsReadOnly()
        {
            #region Step 1: Create the item with setting WebClientReadFormQueryString.
            ItemType[] createdItems = new ItemType[] { new ItemType() };
            createdItems[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            createdItems[0].StoreEntryId = new byte[5] { 1, 2, 3, 4, 5 };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);
            if (Common.IsRequirementEnabled(204511, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R204511");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R204511
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.ErrorInvalidPropertySet,
                    createItemResponse.ResponseMessages.Items[0].ResponseCode,
                    204511,
                    @"[In Appendix C: Product Behavior]It [StoreEntryId] is read-only for the client and will be ignored by the server.(<62> Section 2.2.4.24:  In Exchange 2010 SP2, if the StoreEntryId is specified in a CreateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.)");
            }

            if (Common.IsRequirementEnabled(204521, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R204521");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R204521
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.NoError,
                    createItemResponse.ResponseMessages.Items[0].ResponseCode,
                    204521,
                    @"[In Appendix C: Product Behavior] It [StoreEntryId] is read-only for the client and will be ignored by the server.(<62> Section 2.2.4.24:  In Exchange 2013 and above,  if the StoreEntryId is specified in a CreateItem request, will be ignored by server.)");
            }
            #endregion

            #region Step 2: Update created item with setting StoreEntryId.
            ItemType item = new ItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);

            UpdateItemResponseType updateItemResponse;
            ItemChangeType[] itemChanges;

            itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();

            // Update the created item.
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemStoreEntryId
            };
            setItem.Item1 = new ItemType()
            {
                StoreEntryId = new byte[5] { 1, 2, 3, 4, 5 }
            };
            itemChanges[0].Updates[0] = setItem;

            // Call UpdateItem to update the body of the created item, by using ItemId in CreateItem response.
            updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            ItemType[] updateItems = Common.GetItemsFromInfoResponse<ItemType>(updateItemResponse);
 
            if (Common.IsRequirementEnabled(204512, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R204512");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R204512
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.ErrorInvalidPropertySet,
                    updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                    204512,
                    @"[In Appendix C: Product Behavior]It [StoreEntryId] is read-only for the client and will be ignored by the server.(<62> Section 2.2.4.24:  In Exchange 2010 SP2, if the StoreEntryId is specified in a  UpdateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.)");
            }

            if (Common.IsRequirementEnabled(204522, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R204522");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R204522
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.ErrorInvalidPropertySet,
                    updateItemResponse.ResponseMessages.Items[0].ResponseCode,
                    204522,
                    @"[In Appendix C: Product Behavior] It [StoreEntryId] is read-only for the client and will be ignored by the server.(<62> Section 2.2.4.24:  In Exchange 2013 and above,  if the StoreEntryId is specified in a UpdateItem request, will be ignored by server.)");
            }
            #endregion

        }

        /// <summary>
        /// This test case is intended to validate ErrorInvalidPropertySet is returned if WebClientEditFormQueryString is specified in request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S01_TC46_VerifyBlockStatusIsSetInCreateItem()
        {
            if (Common.IsRequirementEnabled(1357, this.Site))
            {
                #region Step 1: Create the item with BlockStatus.
                ItemType[] createdItems = new ItemType[] { new ItemType() };
                createdItems[0].Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem);
                createdItems[0].BlockStatusSpecified= true;
                createdItems[0].BlockStatus = true;

                CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, createdItems);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2355");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2355
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.ErrorInvalidPropertySet,
                    createItemResponse.ResponseMessages.Items[0].ResponseCode,
                    2355,
                    @"[In t:ItemType Complex Type] but if [BlockStatus] specified in a CreateItem request, an ErrorInvalidPropertySet ([MS-OXWSCDATA] section 2.2.5.24) will be returned.");
                #endregion
            }
        }
    }
    #endregion
}