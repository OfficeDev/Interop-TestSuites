//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;
    
    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, deletion and mark of calendar items on the server.
    /// </summary>
    [TestClass]
    public class S05_ManageMeetingItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for calendar item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC01_CreateGetDeleteMeetingItemSuccessfully()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for calendar item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC02_CopyMeetingItemSuccessfully()
        {
            #region Step 1: Create the calendar item.
            CalendarItemType item = new CalendarItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            ItemIdId itemIdId = this.ITEMIDAdapter.ParseItemId(createdItemIds[0]);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R66");

            // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R66
            Site.CaptureRequirementIfAreEqual<IdProcessingInstructionType>(
                IdProcessingInstructionType.Normal,
                (IdProcessingInstructionType)itemIdId.IdProcessingInstruction,
                "MS-OXWSITEMID",
                66,
                @"[In Id Storage Type (byte)] The Id processing uses the value of the following enumeration.
                        /// <summary>
                        /// Indicates any special processing to perform on an Id when deserializing it.
                        /// </summary>
                        internal enum IdProcessingInstruction : byte
                        {
                            /// <summary>
                            /// No special processing.  The Id represents a PR_ENTRY_ID
                            /// </summary>
                            Normal = 0,

                    [        /// <summary>
                            /// The Id represents an OccurenceStoreObjectId and therefore
                           /// must be deserialized as a StoreObjectId.
                            /// </summary>
                            Recurrence = 1,

                            /// <summary>
                            /// The Id represents a series.
                            /// </summary>
                            Series = 2,]
                        }");
            #endregion

            #region Step 2: Copy the calendar item.
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion 

            #region Step 3: Get the first created calendar item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion 

            #region Step 4: Get the second copied calendar item success.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for calendar item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC03_MoveMeetingItemSuccessfully()
        {
            #region Step 1: Create the calendar item.
            CalendarItemType item = new CalendarItemType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the calendar item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion 

            #region Step 3: Get the created calendar item failed.
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
                    "Get calendar item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion 

            #region Step 4: Get the moved calendar item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for calendar item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC04_UpdateMeetingItemSuccessfully()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for calendar items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC05_MarkAllMeetingItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            CalendarItemType[] items = new CalendarItemType[] { new CalendarItemType(), new CalendarItemType() };
            this.TestSteps_VerifyMarkAllItemsAsRead(items);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for calendar item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC06_UpdateMeetingItemFailed()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for calendar item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC07_CreateMeetingItemFailed()
        {
            #region Step 1: Create the calendar item with invalid item class.
            CalendarItemType[] items = new CalendarItemType[]
            { 
                new CalendarItemType()
                { 
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem),

                    // Set an invalid ItemClass to calendar item.
                    ItemClass = TestSuiteHelper.InvalidItemClass
                }
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.calendar, items);
            #endregion 

            // Get ResponseCode from CreateItem operation response.
            ResponseCodeType responseCode = createItemResponse.ResponseMessages.Items[0].ResponseCode;

            // Verify MS-OXWSCDATA_R619.
            this.VerifyErrorObjectTypeChanged(responseCode);
        }

        /// <summary>
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for calendar item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC08_VerifyExtendPropertyType()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.calendar, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.calendar, item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get and copy the multiple calendar items with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC09_OperateMultipleMeetingItemsSuccessfully()
        {
            CalendarItemType[] items = new CalendarItemType[] { new CalendarItemType(), new CalendarItemType() };
            this.TestSteps_VerifyOperateMultipleItems(items);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC10_GetMeetingItemWithItemResponseShapeType()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which IncludeMimeContent element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC11_GetMeetingItemWithIncludeMimeContent()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC12_GetMeetingItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC13_GetMeetingItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC14_GetMeetingItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC15_GetMeetingItemWithDefaultShapeNamesTypeEnum()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC16_GetMeetingItemWithBodyTypeResponseTypeEnum()
        {
            CalendarItemType item = new CalendarItemType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by GetItem with different ItemId types in GetItem request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC17_GetMeetingItemWithFourItemIdTypesSuccessfully()
        {
            #region Step 1: Create a recurring calendar item.
            DateTime start = DateTime.Now;
            int numberOfOccurrences = 5;
            CalendarItemType item = this.CreateAndGetRecurringCalendarItem(start, numberOfOccurrences);
            #endregion

            #region Step 2: Get the first occurrence of the recurring calendar item by OccurrenceItemIdType.
            // The calendar item to get.
            OccurrenceItemIdType[] occurrenceItemId = new OccurrenceItemIdType[1];
            occurrenceItemId[0] = new OccurrenceItemIdType();
            occurrenceItemId[0].RecurringMasterId = item.ItemId.Id;
            occurrenceItemId[0].ChangeKey = item.FirstOccurrence.ItemId.ChangeKey;
            occurrenceItemId[0].InstanceIndex = 1;

            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(occurrenceItemId);
            
            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            CalendarItemType[] getCalendarOccurences = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getCalendarOccurences.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getCalendarOccurences.GetLength(0));

            ItemIdType[] itemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);
            ItemIdId itemIdId = this.ITEMIDAdapter.ParseItemId(itemIds[0]);
            #endregion

            #region Step 3: Get the recurrence master calendar item by RecurringMasterItemIdType.
            // The calendar item to get.
            RecurringMasterItemIdType[] recurringMasterItemId = new RecurringMasterItemIdType[1];
            recurringMasterItemId[0] = new RecurringMasterItemIdType();

            // Use the first occurrence item id and change key to form the recurringMasterItemId
            recurringMasterItemId[0].OccurrenceId = item.FirstOccurrence.ItemId.Id;
            recurringMasterItemId[0].ChangeKey = item.FirstOccurrence.ItemId.ChangeKey;

            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(recurringMasterItemId);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            CalendarItemType[] getCalendarRecurring = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getCalendarRecurring.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getCalendarRecurring.GetLength(0));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R67");

            // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R67
            Site.CaptureRequirementIfAreEqual<IdProcessingInstructionType>(
                IdProcessingInstructionType.Recurrence,
                (IdProcessingInstructionType)itemIdId.IdProcessingInstruction,
                "MS-OXWSITEMID",
                67,
                @"[In Id Storage Type (byte)] The Id processing uses the value of the following enumeration.
                        /// <summary>
                        /// Indicates any special processing to perform on an Id when deserializing it.
                        /// </summary>
                        internal enum IdProcessingInstruction : byte
                        {
                    [        /// <summary>
                            /// No special processing.  The Id represents a PR_ENTRY_ID
                            /// </summary>
                            Normal = 0,
                    ]
                             /// <summary>
                            /// The Id represents an OccurenceStoreObjectId and therefore
                           /// must be deserialized as a StoreObjectId.
                            /// </summary>
                            Recurrence = 1,

                    [       /// <summary>
                            /// The Id represents a series.
                            /// </summary>
                            Series = 2,]
                    }");
            #endregion
            
            // Clear ExistItemIds for DeleteItem.
            this.ExistItemIds.Clear();
            this.ExistItemIds.Add(item.ItemId);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by DeleteItem with three different ItemId types in DeleteItem request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC18_DeleteMeetingItemWithThreeItemIdTypesSuccessfully()
        {
            #region Step 1: Create and get a recurring calendar item.
            DateTime start = DateTime.Now;
            int numberOfOccurrences = 5;
            CalendarItemType calendar = this.CreateAndGetRecurringCalendarItem(start, numberOfOccurrences);

            #endregion

            #region Step 2: Delete the recurring calendar item by ItemIdType.
            DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion 

            #region Step 3: Create and get a recurring calendar item.
            calendar = this.CreateAndGetRecurringCalendarItem(start, numberOfOccurrences);

            #endregion 

            #region Step 4: Delete the recurrence master calendar item by RecurringMasterItemIdType.
            DeleteItemType deleteItemRequest = new DeleteItemType();

            // The calendar item to delete.
            RecurringMasterItemIdType[] recurringMasterItemIds = new RecurringMasterItemIdType[1];
            recurringMasterItemIds[0] = new RecurringMasterItemIdType();
            recurringMasterItemIds[0].OccurrenceId = calendar.FirstOccurrence.ItemId.Id;
            recurringMasterItemIds[0].ChangeKey = calendar.FirstOccurrence.ItemId.ChangeKey;

            deleteItemRequest.ItemIds = recurringMasterItemIds;

            // Enumeration value to describe how an item is to be deleted.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            // AffectedTaskOccurrences indicates whether a task instance or a task master is to be deleted.
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;

            // SendMeetingCancellations describes how cancellations are to be handled for deleted meetings.
            deleteItemRequest.SendMeetingCancellationsSpecified = true;
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 5: Create and get a recurring calendar item.
            calendar = this.CreateAndGetRecurringCalendarItem(start, numberOfOccurrences);

            #endregion

            #region Step 6: Delete the first occurrence of the recurring calendar item by OccurrenceItemIdType.
            deleteItemRequest = new DeleteItemType();

            // The calendar item to delete.
            OccurrenceItemIdType[] occurrenceItemIds = new OccurrenceItemIdType[1];
            occurrenceItemIds[0] = new OccurrenceItemIdType();
            occurrenceItemIds[0].RecurringMasterId = calendar.ItemId.Id;
            occurrenceItemIds[0].ChangeKey = calendar.FirstOccurrence.ItemId.ChangeKey;
            occurrenceItemIds[0].InstanceIndex = 1;

            deleteItemRequest.ItemIds = occurrenceItemIds;

            // Enumeration value to describe how an item is to be deleted.
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            // AffectedTaskOccurrences indicates whether a task instance or a task master is to be deleted.
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;

            // SendMeetingCancellations describes how cancellations are to be handled for deleted meetings.
            deleteItemRequest.SendMeetingCancellationsSpecified = true;
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the ReminderIsSet Boolean values for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC19_VerifyItemWithReminderIsSet()
        {
            #region Step 1: Create the items.
            // Define two calendar items, the first one has the reminder set and the second one hasn't the reminder set.
            CalendarItemType[] calendar = new CalendarItemType[2]
            {
                new CalendarItemType
                {
                    ReminderIsSetSpecified = true,
                    ReminderIsSet = true,
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        "ReminderIsSet")
                },
                new CalendarItemType
                {
                    ReminderIsSetSpecified = true,
                    ReminderIsSet = false,
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        "ReminderIsNotSet")
                }
            };

            // Call the CreateItem operation.
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.calendar, calendar);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // Two created items should be returned.
            Site.Assert.AreEqual<int>(
                 2,
                 createdItemIds.GetLength(0),
                 "Two created items should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 2,
                 createdItemIds.GetLength(0));

            #endregion

            #region Step 2: Get the items.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                 2,
                 getItemIds.GetLength(0),
                 "Two items should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 2,
                 getItemIds.GetLength(0));

            // The ReminderIsSet of the second item should not be set to true.
            ItemInfoResponseMessageType itemInfoResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            ArrayOfRealItemsType arrayOfRealItemsType = itemInfoResponseMessage.Items;

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1616");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1616
            // The schema is validated and the ReminderIsSet element is true, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsTrue(
                arrayOfRealItemsType.Items[0].ReminderIsSet,
                1616,
                @"[In t:ItemType Complex Type] [ReminderIsSet is] True, indicates a reminder has been set for an item.");

            // The ReminderIsSet of the second item should not be set to false.
            itemInfoResponseMessage = getItemResponse.ResponseMessages.Items[1] as ItemInfoResponseMessageType;
            arrayOfRealItemsType = itemInfoResponseMessage.Items;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1617");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1617
            // The schema is validated and the ReminderIsSet element is false, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsFalse(
                arrayOfRealItemsType.Items[0].ReminderIsSet,
                1617,
                @"[In t:ItemType Complex Type] otherwise [ReminderIsSet is] false, indicates [a reminder has not been set for an item].");
            #endregion
        }

        /// <summary>
        /// Create and get a recurring calendar item.
        /// </summary>
        /// <param name="start">The start time of a recurring calendar item</param>
        /// <param name="numberOfOccurrences">The number of occurrences of a recurring calendar item</param>
        /// <returns>The created recurring calendar item</returns>
        private CalendarItemType CreateAndGetRecurringCalendarItem(DateTime start, int numberOfOccurrences)
        {
            #region Step 1: Create a recurring calendar item.
            CalendarItemType[] items = new CalendarItemType[] { new CalendarItemType() };
            items[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            items[0].Recurrence = new RecurrenceType();

            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = 1;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = numberOfOccurrences;
            range.StartDate = start;

            items[0].Recurrence.Item = pattern;
            items[0].Recurrence.Item1 = range;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.calendar, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdCalendarItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 createdCalendarItemIds.GetLength(0),
                 "One created calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdCalendarItemIds.GetLength(0));
            #endregion 

            #region Step 2: Get the recurring calendar item by ItemIdType.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdCalendarItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            CalendarItemType[] getCalendarItems = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse);

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getCalendarItems.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getCalendarItems.GetLength(0));

            Site.Assert.IsNotNull(
                 getCalendarItems[0].FirstOccurrence,
                 "FirstOccurrence element in calendar item should not be null!");

            return getCalendarItems[0];
            #endregion 
        }
        #endregion
    }
}