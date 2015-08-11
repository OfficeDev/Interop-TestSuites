namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, deletion and mark of task items on the server.
    /// </summary>
    [TestClass]
    public class S07_ManageTaskItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for task item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC01_CreateGetDeleteTaskItemSuccessfully()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for task item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC02_CopyTaskItemSuccessfully()
        {
            #region Step 1: Create the task item.
            TaskType item = new TaskType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Copy the task item.
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied task item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the first created task item success.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One task item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion

            #region Step 4: Get the second copied task item success.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One task item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for task item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC03_MoveTaskItemSuccessfully()
        {
            #region Step 1: Create the task item.
            TaskType item = new TaskType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the task item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved task item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the created task item failed.
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
                    "Get task item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion

            #region Step 4: Get the moved task item.
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One task item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for task item with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC04_UpdateTaskItemSuccessfully()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for task items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC05_MarkAllTaskItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            TaskType[] items = new TaskType[] { new TaskType(), new TaskType() };
            this.TestSteps_VerifyMarkAllItemsAsRead(items);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by CreateItem operation with ErrorObjectTypeChanged response code for task item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC06_CreateTaskItemFailed()
        {
            #region Step 1: Create the task item with invalid item class.
            TaskType[] items = new TaskType[]
            { 
                new TaskType() 
                { 
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem),

                    // Set an invalid ItemClass to post item.
                    ItemClass = TestSuiteHelper.InvalidItemClass
                }
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.tasks, items);

            #endregion

            // Get ResponseCode from CreateItem operation response.
            ResponseCodeType responseCode = createItemResponse.ResponseMessages.Items[0].ResponseCode;

            // Verify MS-OXWSCDATA_R619.
            this.VerifyErrorObjectTypeChanged(responseCode);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for task item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC07_UpdateTaskItemFailed()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for task item.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC08_VerifyExtendPropertyType()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.tasks, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.tasks, item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get and copy the multiple task items with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC09_OperateMultipleTaskItemsSuccessfully()
        {
            TaskType[] items = new TaskType[] { new TaskType(), new TaskType() };
            this.TestSteps_VerifyOperateMultipleItems(items);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC10_GetTaskItemWithItemResponseShapeType()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC11_GetTaskItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC12_GetTaskItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC13_GetTaskItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC14_GetTaskItemWithDefaultShapeNamesTypeEnum()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC15_GetTaskItemWithBodyTypeResponseTypeEnum()
        {
            TaskType item = new TaskType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by GetItem operation for ItemIdType and RecurringMasterItemIdRangesType item ID types in NonEmptyArrayOfBaseItemIdsType element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC16_GetTaskItemWithTwoItemIdTypesSuccessfully()
        {
            #region Step 1: Create a recurring task item.
            // Define the pattern and range of the recurring task item.
            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = 1;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            int numberOfOccurrences = 5;
            range.NumberOfOccurrences = numberOfOccurrences;
            System.DateTime start = System.DateTime.Now;
            range.StartDate = start;

            // Define the TaskType item.
            TaskType[] items = new TaskType[] { new TaskType() };
            items[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            items[0].Recurrence = new TaskRecurrenceType();
            items[0].Recurrence.Item = pattern;
            items[0].Recurrence.Item1 = range;

            // Call CreateItem operation.
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.tasks, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            BaseItemIdType[] createdTaskItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created task item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 createdTaskItemIds.GetLength(0),
                 "One created task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdTaskItemIds.GetLength(0));
            #endregion

            #region Step 2: Get the recurring task item by ItemIdType.
            // Call GetItem operation using the created task item ID.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdTaskItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
            #endregion

            // Exchange 2007 and Exchange 2010 do not support the RecurringMasterItemIdRanges element.
            SutVersion currentSutVersion = (SutVersion)Enum.Parse(typeof(SutVersion), Common.GetConfigurationPropertyValue("SutVersion", this.Site));
            if (currentSutVersion.Equals(SutVersion.ExchangeServer2013))
            {
                #region Step 3: Get the recurring task item by RecurringMasterItemIdRangesType.
                // Define the RecurringMasterItemIdRanges using the created task item ID.
                RecurringMasterItemIdRangesType[] recurringMasterItemIdRanges = new RecurringMasterItemIdRangesType[1];
                recurringMasterItemIdRanges[0] = new RecurringMasterItemIdRangesType();
                recurringMasterItemIdRanges[0].Id = (createdTaskItemIds[0] as ItemIdType).Id;
                recurringMasterItemIdRanges[0].ChangeKey = (createdTaskItemIds[0] as ItemIdType).ChangeKey;
                recurringMasterItemIdRanges[0].Ranges = new OccurrencesRangeType[1];
                recurringMasterItemIdRanges[0].Ranges[0] = new OccurrencesRangeType();
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTimeSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].CompareOriginalStartTime = true;
                recurringMasterItemIdRanges[0].Ranges[0].StartSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].Start = start;
                recurringMasterItemIdRanges[0].Ranges[0].EndSpecified = true;
                recurringMasterItemIdRanges[0].Ranges[0].End = start.AddDays(numberOfOccurrences);

                // Call GetItem operation using the recurringMasterItemIdRanges.
                getItemResponse = this.CallGetItemOperation(recurringMasterItemIdRanges);

                // Check the operation response.
                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
                #endregion
            }
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by GetItem operation with IconIndex child element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S07_TC17_GetTaskItemWithIconIndexSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1917, this.Site), "Exchange 2007 and Exchange 2010 do not support the IconIndex element.");

            #region Step 1: Create a recurring task item.
            // Define the pattern and range of the recurring task item.
            DailyRecurrencePatternType pattern = new DailyRecurrencePatternType();
            pattern.Interval = 1;

            NumberedRecurrenceRangeType range = new NumberedRecurrenceRangeType();
            range.NumberOfOccurrences = 5;
            range.StartDate = System.DateTime.Now;

            // Define the TaskType item.
            TaskType[] items = new TaskType[] { new TaskType() };
            items[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            items[0].Recurrence = new TaskRecurrenceType();
            items[0].Recurrence.Item = pattern;
            items[0].Recurrence.Item1 = range;

            // Call CreateItem operation.
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.tasks, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            BaseItemIdType[] createdTaskItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created task item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 createdTaskItemIds.GetLength(0),
                 "One created task item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdTaskItemIds.GetLength(0));
            #endregion

            #region Step 2: Get the recurring task item by ItemIdType.
            GetItemType getItem = new GetItemType();

            // Create item and return the item id.
            getItem.ItemIds = createdTaskItemIds;

            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItem.ItemShape.AdditionalProperties = new PathToUnindexedFieldType[] { new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.itemIconIndex } };

            // Call GetItem operation using the created task item ID.
            GetItemResponseType getItemResponse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
            #endregion

            #region Step 3: Assert the value of IconIndex element is "TaskRecur".
            TaskType[] taskItems = Common.GetItemsFromInfoResponse<TaskType>(getItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 taskItems.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 taskItems.GetLength(0));

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1917");
        
            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1917
            this.Site.CaptureRequirementIfAreEqual<IconIndexType>(
                IconIndexType.TaskRecur,
                taskItems[0].IconIndex,
                1917,
                @"[In Appendix C: Product Behavior] Implementation does support value ""TaskRecur"" of ""IconIndex"" simple type which specifies the recurring task icon. (Exchange 2013 and above follow this behavior.)");
            #endregion
        }
        #endregion
    }
}