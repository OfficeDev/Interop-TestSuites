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

            CalendarItemType[] getItems = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse);

            if (Common.IsRequirementEnabled(2935, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2935");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2935               
                this.Site.CaptureRequirementIfIsNull(getItems[0].Location,
                    "MS-OXWSCORE",
                    2935,
                    @"[In Appendix C: Product Behavior] Implementation will return no data for the Location element which represents the location of the meeting. (<88> Section 2.2.4.27:  Exchange 2013, Exchange 2016, and Exchange 2019 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(2936, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2936");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2936               
                this.Site.CaptureRequirementIfIsNotNull(getItems[0].Location,
                    "MS-OXWSCORE",
                    2936,
                    @"[In Appendix C: Product Behavior] Implementation will return data for the Location element which represents the location of the meeting. (<88> Section 2.2.4.27:  Exchange 2010 follow this behavior.)");
            }

            // One calendar item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One calendar item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2019");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2019
            this.Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Appointment",
                ((ItemInfoResponseMessageType)getItemResponse.ResponseMessages.Items[0]).Items.Items[0].ItemClass,
                2019,
                @"[In t:ItemType Complex Type] This value is ""IPM.Appointment"" for calendar item.");
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
        /// This test case is intended to validate if invalid ItemClass values are set for Meeting items in the CreateItem request,
        /// an ErrorObjectTypeChanged response code will be returned in the CreateItem response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC20_CreateMeetingItemWithInvalidItemClass()
        {
            #region Step 1: Create the Meeting item with ItemClass set to IPM.DistList.
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            CalendarItemType item = new CalendarItemType();
            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.Items.Items[0].ItemClass = "IPM.DistList";
            DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
            folderIdForCreateItems.Id = DistinguishedFolderIdNameType.calendar;
            createItemRequest.SavedItemFolderId = new TargetFolderIdType();
            createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
            createItemRequest.SendMeetingInvitationsSpecified = true;

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Meeting item with ItemClass IPM.DistList.");
            #endregion

            #region Step 2: Create the Meeting item with ItemClass set to IPM.Post.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Post";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Meeting item with ItemClass IPM.Post.");
            #endregion

            #region Step 3: Create the Meeting item with ItemClass set to IPM.Task.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 3);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Task";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Meeting item with ItemClass IPM.Task.");
            #endregion

            #region Step 4: Create the Meeting item with ItemClass set to IPM.Contact.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 4);
            createItemRequest.Items.Items[0].ItemClass = "IPM.Contact";
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Meeting item with ItemClass IPM.Contact.");
            #endregion

            #region Step 5: Create the Meeting item with ItemClass set to random string.
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 5);
            createItemRequest.Items.Items[0].ItemClass = Common.GenerateResourceName(this.Site, "ItemClass");
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Site.Assert.AreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorObjectTypeChanged,
                createItemResponse.ResponseMessages.Items[0].ResponseCode,
                "ErrorObjectTypeChanged should be returned if create a Meeting item with ItemClass is set to a random string.");
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2023");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2023
            this.Site.CaptureRequirement(
                2023,
                @"[In t:ItemType Complex Type] If invalid values are set for these items in the CreateItem request, an ErrorObjectTypeChanged ([MS-OXWSCDATA] section 2.2.5.24) response code will be returned in the CreateItem response.");
        }

        /// <summary>
        /// This case is intended to validate ProposeNewTime in ResponseObjects for calendar item from successful response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC21_ResponseObjectsProposeNewTime()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2302, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the ProposeNewTime element. ");

            #region Organizer sends meeting to attendee.
            CalendarItemType item = new CalendarItemType();
            item.RequiredAttendees = new AttendeeType[1];
            EmailAddressType attendeeEmail = new EmailAddressType();
            attendeeEmail.EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            AttendeeType attendee = new AttendeeType();
            attendee.Mailbox = attendeeEmail;
            item.RequiredAttendees[0] = attendee;

            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem);
            createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendOnlyToAll;
            createItemRequest.SendMeetingInvitationsSpecified = true;

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

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

            #region Attendee gets the meeting request
            ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User2");
            Site.Assert.AreEqual<int>(1, findItemIds.Length, "Attendee should receive the meeting request");

            GetItemResponseType getItemResponse = this.CallGetItemOperation(findItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            // Check whether the child elements of ResponseObjects have been returned successfully.
            ItemInfoResponseMessageType getItems = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            ResponseObjectType[] responseObjects = getItems.Items.Items[0].ResponseObjects;
            foreach (ResponseObjectType responseObject in responseObjects)
            {
                if (responseObject.GetType() == typeof(ProposeNewTimeType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2302");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2302
                    // Element ProposeNewTime is returned from server, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        2302,
                        @"[In Appendix C: Product Behavior] Implementation does support the ProposeNewTime element which specifies the response object that is used to propose a new time. (<82> Section 2.2.4.33:  This element [ProposeNewTime] was introduced in Exchange 2013 SP1.)");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2135");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2135
                    // Element ProposeNewTime is returned from server and pass schema validation, this requirement can be captured directly.
                    this.Site.CaptureRequirement(
                        2135,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] The type of ProposeNewTime is t:ProposeNewTimeType ([MS-OXWSCDATA] section 2.2.4.38).");
                    break;
                }
            }

            this.CleanItemsSentOut(new string[] { createItemRequest.Items.Items[0].Subject });
            this.ExistItemIds.Remove(getItems.Items.Items[0].ItemId);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the CompareOriginalStartTime.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S05_TC22_VerifyCompareOriginalStartTime()
        {
            #region Step 1: Create and get a recurring calendar item.
            DateTime start = DateTime.Now;
            int numberOfOccurrences = 5;
            CalendarItemType calendar = this.CreateAndGetRecurringCalendarItem(start, numberOfOccurrences);

            #endregion

            #region Step 2: Get the first occurrence of the recurring calendar item by OccurrenceItemIdType.
            // The calendar item to get.
            OccurrenceItemIdType[] occurrenceItemId = new OccurrenceItemIdType[1];
            occurrenceItemId[0] = new OccurrenceItemIdType();
            occurrenceItemId[0].RecurringMasterId = calendar.ItemId.Id;
            occurrenceItemId[0].ChangeKey = calendar.FirstOccurrence.ItemId.ChangeKey;
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

            #region Step 3: Update the start date of the calendar item.

            ItemChangeType itemChange = new ItemChangeType();
            itemChange.Item = itemIds[0];

            CalendarItemType calendarChange = new CalendarItemType();
            calendarChange.Start = calendar.Start.AddMinutes(20);
            calendarChange.StartSpecified = true;

            itemChange.Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItemField = new SetItemFieldType();
            setItemField.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.calendarStart
            };

            setItemField.Item1 = calendarChange;
            itemChange.Updates[0] = setItemField;

            UpdateItemResponseType updatedItem = this.CallUpdateItemOperation(DistinguishedFolderIdNameType.calendar, true, new ItemChangeType[] { itemChange });

            #endregion

            SutVersion currentSutVersion = (SutVersion)Enum.Parse(typeof(SutVersion), Common.GetConfigurationPropertyValue("SutVersion", this.Site));
            if (currentSutVersion.Equals(SutVersion.ExchangeServer2016))
            {
                #region Step 4: Get the recurring calendar item by RecurringMasterItemIdRangesType with set CompareOriginalStartTime to true.

                // The calendar item to get.
                RecurringMasterItemIdRangesType[] recurringMasterItemIdRanges = new RecurringMasterItemIdRangesType[1];
                recurringMasterItemIdRanges[0] = new RecurringMasterItemIdRangesType();

                // Use the first occurrence item id and change key to form the recurringMasterItemId
                recurringMasterItemIdRanges[0].Id = calendar.ItemId.Id;
                recurringMasterItemIdRanges[0].ChangeKey = calendar.ItemId.ChangeKey;
                recurringMasterItemIdRanges[0].Ranges = new OccurrencesRangeType[1];
                recurringMasterItemIdRanges[0].Ranges[0] = new OccurrencesRangeType();
                recurringMasterItemIdRanges[0].Ranges[0].Start = calendar.Start.AddMinutes(10);
                recurringMasterItemIdRanges[0].Ranges[0].StartSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].End = start.AddDays(5);
                recurringMasterItemIdRanges[0].Ranges[0].EndSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].Count = 5;
                recurringMasterItemIdRanges[0].Ranges[0].CountSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTime = true;
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTimeSpecified = true;

                // Call the GetItem operation.
                GetItemResponseType getItemResponse1 = this.CallGetItemOperation(recurringMasterItemIdRanges);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse1, 1, this.Site);

                CalendarItemType[] getCalendarRecurring = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse1);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1697");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1697
                this.Site.CaptureRequirementIfAreEqual(
                    5,
                    getCalendarRecurring.Length,
                    1697,
                    @"[In t:OccurrencesRangeType Complex Type] [CompareOriginalStartTime is] True, indicates comparing the specified ranges to an original start time.");

                #endregion

                #region Step 5: Get the recurrence master calendar item by RecurringMasterItemIdRangesType with set CompareOriginalStartTime to false.

                // The calendar item to get.
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTime = false;
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTimeSpecified = true;

                // Call the GetItem operation.
                getItemResponse1 = this.CallGetItemOperation(recurringMasterItemIdRanges);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse1, 1, this.Site);

                getCalendarRecurring = Common.GetItemsFromInfoResponse<CalendarItemType>(getItemResponse1);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1698");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1698
                this.Site.CaptureRequirementIfAreEqual(
                    6,
                    getCalendarRecurring.Length,
                    1698,
                    @"[In t:OccurrencesRangeType Complex Type] otherwise [CompareOriginalStartTime is] false, indicates comparing the specified ranges to a pair of start and end values.");
                #endregion
            }

            // Clear ExistItemIds for DeleteItem.
            this.ExistItemIds.Clear();
            this.ExistItemIds.Add(calendar.ItemId);
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