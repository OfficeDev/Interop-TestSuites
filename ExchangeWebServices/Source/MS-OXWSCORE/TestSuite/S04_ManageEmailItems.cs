namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, sending, deletion and mark of email items on the server.
    /// </summary>
    [TestClass]
    public class S04_ManageEmailItems : TestSuiteBase
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
        /// This test case is intended to validate the successful responses returned by CreateItem, GetItem and DeleteItem operations for message with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC01_CreateGetDeleteEmailItemSuccessfully()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyCreateGetDeleteItem(item);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, CopyItem and GetItem operations for message with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC02_CopyEmailItemSuccessfully()
        {
            #region Step 1:Create the message type item
            MessageType item = new MessageType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2:Copy the message type item
            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 1, this.Site);

            ItemIdType[] copiedItemIds = Common.GetItemIdsFromInfoResponse(copyItemResponse);

            // One copied message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 copiedItemIds.GetLength(0),
                 "One copied message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 copiedItemIds.GetLength(0));
            #endregion

            #region Step 3:Get the first created message type item success
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);            

            // One message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion

            #region Step 4:Get the second copied message type item success
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(copiedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MoveItem and GetItem operations for message with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC03_MoveEmailItemSuccessfully()
        {
            #region Step 1: Create the message type item.
            MessageType item = new MessageType();
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Move the message type item.
            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);

            ItemIdType[] movedItemIds = Common.GetItemIdsFromInfoResponse(moveItemResponse);

            // One moved message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 movedItemIds.GetLength(0),
                 "One moved message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 movedItemIds.GetLength(0));
            #endregion

            #region Step 3: Get the created message type item failed.
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
                    "Get message type item operation should be failed with error! Actual response code: {0}",
                    getItemResponse.ResponseMessages.Items[0].ResponseCode));
            #endregion

            #region Step 4:Get the moved message type item
            // Call the GetItem operation.
            getItemResponse = this.CallGetItemOperation(movedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, GetItem and SendItem operations for message with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC04_SendEmailItemSuccessfully()
        {
            #region Step 1: Create the message type item.
            string itemSubject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            ItemIdType[] createdItemIds = this.CreateItemWithRecipient(itemSubject);
            #endregion

            #region Step 2: Get the message type item.
            // Call the GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One message type item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 getItemIds.GetLength(0),
                 "One message type item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 getItemIds.GetLength(0));

            #endregion

            #region Step 3: Send the message type item.
            // Call SendItem to send the created message item from Inbox folder to the recipient identified by the ToRecipients element, by using createdItemIds in CreateItem response.
            SendItemResponseType sendResponse = this.CallSendItemOperation(
                createdItemIds,
                DistinguishedFolderIdNameType.sentitems,
                false);

            // Check the operation response.
            Common.CheckOperationSuccess(sendResponse, 1, this.Site);

            // Clear ExistItemIds for SendItem, since the item has been sent out and no copy remains.
            this.InitializeCollection();

            #endregion

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R456");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R456
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                sendResponse,
                456,
                @"[In m:SendItemResponseType Complex Type] The SendItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1041");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1041
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                sendResponse.ResponseMessages.Items[0],
                "MS-OXWSCDATA",
                1041,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""SendItemResponseMessage"" with type ""m:ResponseMessageType(section 2.2.4.57)"" specifies the response message for the SendItem operation ([MS-OXWSCORE] section 3.1.4.8).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1588");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1588
            // The schema is validated and the response is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                sendResponse.ResponseMessages.Items[0],
                "MS-OXWSCDATA",
                1588,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""SendItemResponseMessage"" is ""m:ResponseMessageType"" (section 2.2.4.57) type.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R462");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R462
            // The item was sent successfully, so the ItemIds in SendItem operation request specifies the items to send.
            this.Site.CaptureRequirement(
                462,
                @"[In m:SendItemType Complex Type] [The element ""ItemIds""] Specifies an array of item identifiers for the items to try to send.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R49");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R49
            // The schema is validated, so this requirement can be captured.
            this.Site.CaptureRequirement(
                49,
                @"[In m:SendItemResponseType Complex Type] The SendItemResponseType complex type extends the BaseResponseMessageType complex type ([MS-OXWSCDATA] section 2.2.4.16).");

            #region Step 4: Clean all items sent out.
            this.CleanItemsSentOut(new string[] { itemSubject });
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, MarkAllItemsAsRead and GetItem operations for messages with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC05_MarkAllEmailItemsAsReadSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1290, this.Site), "Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.");

            MessageType[] items = new MessageType[] { new MessageType(), new MessageType() };
            this.TestSteps_VerifyMarkAllItemsAsRead(items);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem, UpdateItem and GetItem operations for message with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC06_UpdateEmailItemSuccessfully()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyUpdateItemSuccessfulResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for message.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC07_UpdateEmailItemFailed()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyUpdateItemFailedResponse(item);
        }

        /// <summary>
        /// This test case is intended to validate the PathToExtendedFieldType complex type returned by CreateItem operation for message.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC08_VerifyExtendPropertyType()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertySetIdConflictsWithPropertyTag(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyTagRepresentation(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyName(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyTagConflictsWithPropertyId(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.inbox, item);

            this.TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(DistinguishedFolderIdNameType.inbox, item);
        }

        /// <summary>
        /// This test case is intended to create, update, move, get, copy and send the multiple messages with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC09_OperateMultipleEmailItemsSuccessfully()
        {
            #region Step 1: Create multiple items.
            EmailAddressType address = new EmailAddressType()
            {
                EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site),
            };

            MessageType[] items = new MessageType[]
            {
                new MessageType
                {
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem,
                        1),

                    ToRecipients = new EmailAddressType[]
                    {
                        address
                    }
                },
                new MessageType
                {
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForCreateItem,
                        2),

                    ToRecipients = new EmailAddressType[]
                    {
                        address
                    }
                }
            };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 2, this.Site);

            ItemType[] createdItems = Common.GetItemsFromInfoResponse<ItemType>(createItemResponse);

            // Two created items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    createdItems.GetLength(0),
                    "Two created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    createdItems.GetLength(0));
            #endregion

            #region Step 2 - 5: Update, move, get and copy the items.
            this.OperateMultipleItems(createdItems);
            #endregion

            #region Step 6: Send the items and clean the items sent out.
            // Call SendItem to send the created message items from Inbox folder to the recipient identified by the ToRecipients element, by using createdItemIds in CreateItem response.
            ItemIdType[] itemArray = new ItemIdType[this.CopiedItemIds.Count];
            this.CopiedItemIds.CopyTo(itemArray, 0);
            SendItemResponseType sendResponse = this.CallSendItemOperation(
                itemArray,
                DistinguishedFolderIdNameType.sentitems,
                false);

            // Check the operation response.
            Common.CheckOperationSuccess(sendResponse, 2, this.Site);

            // Remove the sent items from ExistItemIds collection, since the items have been sent out and no copy remains.
            foreach (ItemIdType itemId in itemArray)
            {
                this.ExistItemIds.Remove(itemId);
            }

            // Clear the sent items.
            string[] itemSubjects = new string[2];
            itemSubjects[0] = items[0].Subject;
            itemSubjects[1] = items[1].Subject;
            this.CleanItemsSentOut(itemSubjects);
            #endregion
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC10_GetEmailItemWithItemResponseShapeType()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which IncludeMimeContent element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC11_GetEmailItemWithIncludeMimeContent()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC12_GetEmailItemWithConvertHtmlCodePageToUTF8()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(21498, this.Site), "Exchange 2007 and Exchange 2010 do not include the ConvertHtmlCodePageToUTF8 element.");

            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC13_GetEmailItemWithAddBlankTargetToLinks()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149908, this.Site), "Exchange 2007 and Exchange 2010 do not use the AddBlankTargetToLinks element.");

            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the response returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC14_GetEmailItemWithBlockExternalImages()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2149905, this.Site), "Exchange 2007 and Exchange 2010 do not use the BlockExternalImages element.");

            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC15_GetEmailItemWithDefaultShapeNamesTypeEnum()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC16_GetEmailItemWithBodyTypeResponseTypeEnum()
        {
            MessageType item = new MessageType();
            this.TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum(item);
        }

        /// <summary>
        /// This case is intended to validate ReplyToItemType, ForwardItemType, ReplyAllToItemType and SuppressReadReceiptType in ResponseObjects for email item from successful response.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC17_VerifyEmailWithResponseObjects()
        {
            #region Step 1: Create a message with IsReadReceiptRequested.
            // Define the MessageType item which will be created.
            MessageType[] messages = new MessageType[] { new MessageType() };
            messages[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            EmailAddressType email = new EmailAddressType();
            email.EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            messages[0].ToRecipients = new EmailAddressType[1];
            messages[0].ToRecipients[0] = email;
            messages[0].IsReadReceiptRequestedSpecified = true;
            messages[0].IsReadReceiptRequested = true;
            
            // Define the request of CreateItem operation.
            CreateItemType requestItem = new CreateItemType();
            requestItem.MessageDispositionSpecified = true;
            requestItem.MessageDisposition = MessageDispositionType.SendOnly;
            requestItem.Items = new NonEmptyArrayOfAllItemsType();
            requestItem.Items.Items = messages;

            // Call the CreateItem operation.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(requestItem);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2: Get the message with response objects in inbox.
            // Find the received item in the Inbox folder.
            ItemIdType[] foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, messages[0].Subject, "User1");

            // The result of FindItemsInFolder should not be null.
            Site.Assert.IsNotNull(
                foundItems,
                "The result of FindItemsInFolder should not be null.");

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 foundItems.GetLength(0),
                 "One item should be returned! Expected count: {0}, actual count: {1}",
                 1,
                 foundItems.GetLength(0));

            // Get information from the found item.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(foundItems);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            // Check whether the child elements of ResponseObjects have been returned successfully.
            ItemInfoResponseMessageType getItems = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            ResponseObjectType[] responseObjects = getItems.Items.Items[0].ResponseObjects;
 
            Site.Assert.IsNotNull(responseObjects, "The ResponseObjects should not be null.");

            // Receivers could reply, reply all, forward or send read receipt for the received item, so there should be four child elements in ResponseObjects.
            Site.Assert.AreEqual<int>(
                4,
                responseObjects.Length,
                "Four child elements in ResponseObjects should be returned! Expected count: {0}, actual count: {1}",
                4,
                responseObjects.Length);

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // The type of the child element should be ReplyToItemType, ForwardItemType, ReplyAllToItemType or SuppressReadReceiptType.
            foreach (ResponseObjectType responseObject in responseObjects)
            {
                if (responseObject.GetType() == typeof(ReplyToItemType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1371");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1371
                    // The schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirement(
                        1371,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] The type of ReplyToItem is t:ReplyToItemType ([MS-OXWSCDATA] section 2.2.4.64).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R131");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R131
                    // The response object with ReplyToItemType type is not null and the schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirementIfIsNotNull(
                        responseObject,
                        131,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] [The element ""ReplyToItem""] Specifies a reply to the sender of an item.");
                }
                else if (responseObject.GetType() == typeof(ForwardItemType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1372");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1372
                    // The schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirement(
                        1372,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] The type of ForwardItem is t:ForwardItemType ([MS-OXWSCDATA] section 2.2.4.37).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R132");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R132
                    // The response object with ForwardItemType type is not null and the schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirementIfIsNotNull(
                        responseObject,
                        132,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] [The element ""ForwardItem""] Specifies a server store item to be forwarded to recipients.");
                }
                else if (responseObject.GetType() == typeof(ReplyAllToItemType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1373");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1373
                    // The schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirement(
                        1373,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] The type of ReplyAllToItem is t:ReplyAllToItemType ([MS-OXWSCDATA] section 2.2.4.62).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R133");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R133
                    // The response object with ReplyAllToItemType type is not null and the schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirementIfIsNotNull(
                        responseObject,
                        133,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] [The element ""ReplyAllToItem""] Specifies a reply to the sender and all identified recipients of an item in the server store.");
                }
                else if (responseObject.GetType() == typeof(SuppressReadReceiptType))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1376");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1376
                    // The schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirement(
                        1376,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] The type of SuppressReadReceipt is t:SuppressReadReceiptType ([MS-OXWSCDATA] section 2.2.4.71).");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R136");

                    // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R136
                    // The response object with SuppressReadReceiptType type is not null and the schema is validated, so this requirement can be captured.
                    this.Site.CaptureRequirementIfIsNotNull(
                        responseObject,
                        136,
                        @"[In t:NonEmptyArrayOfResponseObjectsType Complex Type] [The element ""SuppressReadReceipt""] Specifies that read receipts are to be suppressed.");
                }
                else
                {
                    Site.Assume.Fail(
                        string.Format(
                        "The type of responseObject should be one of the following types: ReplyToItemType, ForwardItemType, ReplyAllToItemType or SuppressReadReceiptType, actual {0}",
                        responseObject.GetType()));
                }
            }
            #endregion

            #region Step 3: Delete messages in inbox.
            // Delete the created item.
            this.COREAdapter.DeleteItem(new DeleteItemType() { DeleteType = DisposalType.HardDelete, ItemIds = foundItems });
            this.ExistItemIds.Clear();
            this.FindNewItemsInFolder(DistinguishedFolderIdNameType.inbox);
            #endregion
        }

        /// <summary>
        /// This case is intended to validate the successful responses returned by MarkAsJunk operation.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC18_MarkAsJunkSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1787, this.Site), "Exchange 2007 and Exchange 2010 do not use the MarkAsJunk operation");

            #region Create the successful message
            MessageType[] items = new MessageType[]
            {
                new MessageType
                {
                    Sender = new SingleRecipientType
                    {
                        Item = new EmailAddressType
                        {
                            EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site)
                        }
                    },

                    Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem),
                }
            };

            string itemSubject = items[0].Subject;
            string itemSender = items[0].Sender.Item.EmailAddress;

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            MarkAsJunkType markAsJunkRequest = new MarkAsJunkType();
            markAsJunkRequest.ItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);
            markAsJunkRequest.IsJunk = true;
            markAsJunkRequest.MoveItem = true;

            MarkAsJunkResponseType markAsJunkResponse = this.COREAdapter.MarkAsJunk(markAsJunkRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(markAsJunkResponse, 1, this.Site);

            MarkAsJunkResponseMessageType markAsJunkResponseMessage = (MarkAsJunkResponseMessageType)markAsJunkResponse.ResponseMessages.Items[0];

            if (Common.IsRequirementEnabled(1787, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1787");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1787
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    markAsJunkResponseMessage.ResponseClass,
                    1787,
                    @"[In Appendix C: Product Behavior] Implementation does use [The operation ""MarkAsJunk""] which Marks an item as junk. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1790, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1790");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1790
                this.Site.CaptureRequirementIfAreEqual<ResponseClassType>(
                    ResponseClassType.Success,
                    markAsJunkResponseMessage.ResponseClass,
                    1790,
                    @"[In Appendix C: Product Behavior] Implementation does use the MarkAsJunk operation which marks an item as junk. (Exchange 2013 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1844");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1844
            this.Site.CaptureRequirementIfIsNotNull(
                markAsJunkResponse,
                1844,
                @"[In m:MarkAsJunkResponseType Complex Type] This type [MarkAsJunkResponseType Complex Type] extends the BaseResponseMessageType, as specified in [MS-OXWSCDATA] section 2.2.4.16.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1847");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1847
            this.Site.CaptureRequirementIfIsNotNull(
                markAsJunkResponseMessage,
                1847,
                @"[In m:MarkAsJunkResponseMessageType Complex Type] This type [MarkAsJunkResponseMessageType Complex Type] extends the ResponseMessageType complex type, as specified in [MS-OXWSCDATA] section 2.2.4.65.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R3060");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R3060
            this.Site.CaptureRequirementIfIsNotNull(
                markAsJunkResponseMessage,
                "MS-OXWSCDATA",
                3060,
                @"[In m:ArrayOfResponseMessagesType Complex Type] The element ""MarkAsJunkResponseMessage"" with type ""m:MarkAsJunkResponseMessageType ([MS-OXWSCORE] section 3.1.4.6.3.3)"" specifies the response message for the MarkAsJunk operation ([MS-OXWSCORE] section 3.1.4.6).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1850");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1850
            this.Site.CaptureRequirementIfIsInstanceOfType(
                markAsJunkResponseMessage.MovedItemId,
                typeof(ItemIdType),
                1850,
                @"[In m:MarkAsJunkResponseMessageType Complex Type] The type of MovedItemId is t:ItemIdType (section 2.2.4.19).");

            ItemIdType[] getItemIds = new ItemIdType[] { markAsJunkResponseMessage.MovedItemId };
            GetItemResponseType getItemResponse = this.CallGetItemOperation(getItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1851");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1851
            // When use the MovedItemId to get item and only one item is returned in the response, it means that the MovedItemId specifies an identifier of the moved item, thus this requirement can be verified.
            this.Site.CaptureRequirement(
                1851,
                @"[In m:MarkAsJunkResponseMessageType Complex Type] [The child element ""MovedItemId""] Specifies the item identifier of the moved item. ");

            // Find the item in the Junk Email folder.
            ItemIdType[] foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.junkemail, itemSubject, "User1");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1920");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1920
            // Since the value of IsJunk is true and the value of MoveItem is true, the item is moved to the Junk Email folder. Thus the response of FindItem should not be null, then this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                foundItems,
                1920,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] True and [the value of ""MoveItem"" is] True, The operation moves the email item to the Junk Email folder. ");

            string blockedSender = null;
            string userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
            bool isInBlockedSender = false;
            if (Common.IsRequirementEnabled(1839, this.Site))
            {
                blockedSender = this.CORESUTControlAdapter.GetMailboxJunkEmailConfiguration(userName);

                isInBlockedSender = blockedSender.Contains(itemSender);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1839");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1839
                this.Site.CaptureRequirementIfIsTrue(
                    isInBlockedSender,
                    1839,
                    @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] True and [the value of ""MoveItem"" is] True, the operation adds the sender of the email to the blocked sender list and moves the email item to the Junk Email folder.");
            }
            markAsJunkRequest.ItemIds = foundItems;
            markAsJunkRequest.IsJunk = true;
            markAsJunkRequest.MoveItem = false;
            markAsJunkResponse = this.COREAdapter.MarkAsJunk(markAsJunkRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(markAsJunkResponse, 1, this.Site);

            // Find the item in the Junk Email folder.
            foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.junkemail, itemSubject, "User1");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1921");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1921
            // Since the value of IsJunk is true and the value of MoveItem is false, the item is still in the Junk Email folder. Thus the response of FindItem should not be null, then this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                foundItems,
                1921,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] True and [the value of ""MoveItem"" is] False,The email item is not moved.");

            if (Common.IsRequirementEnabled(1839, this.Site))
            {
                blockedSender = this.CORESUTControlAdapter.GetMailboxJunkEmailConfiguration(userName);
                isInBlockedSender = blockedSender.Contains(itemSender);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1840");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1840
                this.Site.CaptureRequirementIfIsTrue(
                    isInBlockedSender,
                    1840,
                    @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] True and [the value of ""MoveItem"" is] False, the operation adds the sender of the email to the blocked sender list.");
            }
            markAsJunkRequest.ItemIds = foundItems;
            markAsJunkRequest.IsJunk = false;
            markAsJunkRequest.MoveItem = true;
            markAsJunkResponse = this.COREAdapter.MarkAsJunk(markAsJunkRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(markAsJunkResponse, 1, this.Site);

            // Find the item in the Inbox folder.
            foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, itemSubject, "User1");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1922");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1922
            // Since the value of IsJunk is false and the value of MoveItem is true, the item is moved to the inbox folder. Thus the response of FindItem should not be null, then this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                foundItems,
                1922,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] False and [the value of ""MoveItem"" is] True, The operation moves the email item back to the Inbox folder.");

            blockedSender = this.CORESUTControlAdapter.GetMailboxJunkEmailConfiguration(userName);
            isInBlockedSender = blockedSender.Contains(itemSender);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1841");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1841
            this.Site.CaptureRequirementIfIsFalse(
                isInBlockedSender,
                1841,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] False and [the value of ""MoveItem"" is] True, the operation removes the sender from the blocked sender list.");

            markAsJunkRequest.ItemIds = foundItems;
            markAsJunkRequest.IsJunk = false;
            markAsJunkRequest.MoveItem = false;
            markAsJunkResponse = this.COREAdapter.MarkAsJunk(markAsJunkRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(markAsJunkResponse, 1, this.Site);

            // Find the item in the Inbox folder.
            foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, itemSubject, "User1");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1923");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1923
            // Since the value of IsJunk is false and the value of MoveItem is false, the item is still in the Inbox folder. Thus the response of FindItem should not be null, then this requirement can be verified.
            this.Site.CaptureRequirementIfIsNotNull(
                foundItems,
                1923,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] False and [the value of ""MoveItem"" is] False, The email item is not moved.");

            blockedSender = this.CORESUTControlAdapter.GetMailboxJunkEmailConfiguration(userName);
            isInBlockedSender = blockedSender.Contains(itemSender);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1842");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1842
            this.Site.CaptureRequirementIfIsFalse(
                isInBlockedSender,
                1842,
                @"[In m:MarkAsJunkType Complex Type] [When the value of ""IsJunk"" is] False and [the value of ""MoveItem"" is] False, the operation removes the sender from the blocked sender list.");

            this.ExistItemIds.Clear();
            this.ExistItemIds.Add(foundItems[0]);
            #endregion
        }

        /// <summary>
        /// This case is intended to validate the response returned by SendItem operation with the SaveItemToFolder element.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC19_SendItemWithSaveItemToFolder()
        {
            #region Step 1: Create the item.
            string itemSubject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            ItemIdType[] createdItemIds = this.CreateItemWithRecipient(itemSubject);
            #endregion

            #region Step 2: Send the item when SaveItemToFolder is false.
            // Call SendItem to send the created message item from Drafts folder to the recipient identified by the ToRecipients element, by using createdItemIds in CreateItem response.
            SendItemResponseType sendResponse = this.CallSendItemOperation(
                createdItemIds,
                DistinguishedFolderIdNameType.sentitems,
                false);

            // Check the operation response.
            Common.CheckOperationSuccess(sendResponse, 1, this.Site);

            // Clear ExistItemIds for SendItem, since the item has been sent out and no copy remains.
            this.InitializeCollection();
            #endregion

            #region Step 3: Find item in sent items folder.
            // Loop to check whether the sent item has been received by the recepient.
            ItemIdType[] receivedItemId = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, itemSubject, "User2");
            Site.Assert.IsNotNull(receivedItemId, "The email with subject {0} should exist in the inbox folder of User2.", itemSubject);

            // Check whether the sent item has been saved to the SentItems folder.
            this.SwitchUser("User1");
            ItemType[] savedItem = this.FindItemWithRestriction(DistinguishedFolderIdNameType.sentitems, itemSubject);

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1649");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1649
            // The savedItemId is null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNull(
                savedItem,
                1649,
                @"[In m:SendItemType Complex Type] otherwise [SaveItemToFolder is] false, specifies [a copy of a sent item is not saved].");
            #endregion

            #region Step 4: Clean the item sent out which has not been saved.
            this.CleanItemsSentOut(new string[] { itemSubject });
            #endregion

            #region Step 5: Create another item.
            createdItemIds = this.CreateItemWithRecipient(itemSubject);
            #endregion

            #region Step 6: Send the item when SaveItemToFolder is true.
            // Call SendItem to send the created message item from Inbox folder to the recipient identified by the ToRecipients element, by using createdItemIds in CreateItem response.
            sendResponse = this.CallSendItemOperation(
                createdItemIds,
                DistinguishedFolderIdNameType.sentitems,
                true);

            // Check the operation response.
            Common.CheckOperationSuccess(sendResponse, 1, this.Site);

            // Clear ExistItemIds for SendItem, since the item has been sent out and no copy remains.
            this.InitializeCollection();
            #endregion

            #region Step 7: Find item in sent items folder.
            ItemIdType[] savedItemId = this.FindItemsInFolder(DistinguishedFolderIdNameType.sentitems, itemSubject, "User1");

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1648");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1648
            // The savedItemId is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                savedItemId,
                1648,
                @"[In m:SendItemType Complex Type] [SaveItemToFolder is] True, specifies a copy of a sent item is saved.");

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 savedItemId.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 savedItemId.GetLength(0));

            this.ExistItemIds.Add(savedItemId[0]);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R463");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R463
            // If the sent item was found in the sentitems folder, R463 can be captured.
            this.Site.CaptureRequirement(
                463,
                @"[In m:SendItemType Complex Type] [The element ""SavedItemFolderId""] Specifies the identity of the folder that contains a saved version of a sent item.");
            #endregion

            #region Step 8: Clean the item sent out which has been saved.
            this.CleanItemsSentOut(new string[] { itemSubject });
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the SOAP headers for base item in success request.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC20_VerifyOperationsWithSOAPHeaderSuccessful()
        {
            #region Step 1: Create one item with recipient.
            // Clear the soap header.
            this.ClearSoapHeaders();

            // Configure the SOAP headers for CreateItem and UpdateItem operations.
            Dictionary<string, object> headerValues = new Dictionary<string, object>();
            headerValues = this.ConfigureSOAPHeader();

            // Configure the TimeZoneContext SOAP Header.
            TimeZoneContextType timeZoneContext = new TimeZoneContextType();
            timeZoneContext.TimeZoneDefinition = new TimeZoneDefinitionType();
            timeZoneContext.TimeZoneDefinition.Id = TestSuiteHelper.TimeZoneID;

            headerValues.Add("TimeZoneContext", timeZoneContext);
            this.COREAdapter.ConfigureSOAPHeader(headerValues);

            // Call the CreateItem operation and save the item to Drafts folder.
            MessageType[] items = new MessageType[] { this.CreateItemWithOneRecipient() };

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            MessageType[] createdItems = Common.GetItemsFromInfoResponse<MessageType>(createItemResponse);

            // One created items should be returned.
            Site.Assert.AreEqual<int>(
                    1,
                    createdItems.GetLength(0),
                    "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    1,
                    createdItems.GetLength(0));
            #endregion

            #region Step 2: Update the item which is created in Step 1.
            ItemChangeType[] itemChanges = new ItemChangeType[]
            {
                TestSuiteHelper.CreateItemChangeItem(createdItems[0], 1)
            };

            // Clear ExistItemIds list for MoveItem.
            this.InitializeCollection();

            // Call UpdateItem operation to update the subject of the created item, by using the ItemId in CreateItem response.
            UpdateItemResponseType updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);

            // Clear the soap header.
            this.ClearSoapHeaders();
            #endregion

            #region Step 3: Move the item updated in Step 2 from Drafts folder to Inbox folder.
            // Configure the SOAP headers for MoveItem and MarkAllItemsAsRead operations.
            headerValues.Remove("TimeZoneContext");
            this.COREAdapter.ConfigureSOAPHeader(headerValues);

            // Call the MoveItem operation, by using the ItemId in UpdateItem response.
            ItemIdType[] draftsItem = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(draftsItem, 0);
            this.InitializeCollection();
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, draftsItem);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 1, this.Site);
            #endregion

            #region Step 4: Mark all items in Inbox folder as read.
            // Exchange 2007 and Exchange 2010 do not support the MarkAllItemsAsRead operation.
            if (Common.IsRequirementEnabled(1290, this.Site))
            {
                // Configure Inbox folder as the target folder.
                BaseFolderIdType[] folderIds = new BaseFolderIdType[1];
                DistinguishedFolderIdType distinguishedFolder = new DistinguishedFolderIdType();
                distinguishedFolder.Id = DistinguishedFolderIdNameType.inbox;
                folderIds[0] = distinguishedFolder;

                // Mark all items in Inbox folder as unread, and suppress the receive receipts.
                MarkAllItemsAsReadResponseType markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(true, true, folderIds);

                // Check the operation response.
                Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1290");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1290
                // The MarkAllItemsAsRead operation was executed successfully, so this requirement can be captured.
                this.Site.CaptureRequirementIfIsTrue(
                    this.IsSchemaValidated,
                    1290,
                    @"[In Appendix C: Product Behavior] Implementation does support the MarkAllItemsAsRead operation which marks all items in a folder as read. (Exchange 2013  and above follow this behavior.)");
            }

            // Clear the soap header.
            this.ClearSoapHeaders();
            #endregion

            #region Step 5: Get the item in Inbox folder.
            // Configure the SOAP headers for GetItem.
            headerValues.Add("TimeZoneContext", timeZoneContext);
            this.COREAdapter.ConfigureSOAPHeader(headerValues);

            // Call the GetItem operation, by using the ItemId in MoveItem response.
            ItemIdType[] itemArray = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(itemArray, 0);
            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            // Clear the soap header.
            this.ClearSoapHeaders();
            #endregion

            #region Step 6: Copy the message type item from Inbox folder to Drafts folder.
            // Configure the SOAP headers for CopyItem, SendItem and DeleteItem.
            headerValues.Remove("TimeZoneContext");
            this.COREAdapter.ConfigureSOAPHeader(headerValues);

            // Save the ID of the item in Inbox folder and call the CopyItem operation.
            itemArray = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(itemArray, 0);
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, itemArray);

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
            #endregion

            #region Step 7: Send the item in Drafts folder.
            // Call SendItem to send the copied item in Drafts folder, by using copiedItemIds in CopyItem response.
            SendItemResponseType sendResponse = this.CallSendItemOperation(
                copiedItemIds,
                DistinguishedFolderIdNameType.sentitems,
                false);

            // Check the operation response.
            Common.CheckOperationSuccess(sendResponse, 1, this.Site);

            // Remove the sent itemId from ExistItemIds list, since the item has been sent out and no copy remains.
            this.ExistItemIds.Remove(copiedItemIds[0] as ItemIdType);
            #endregion

            #region Step 8: Delete the item in Inbox folder and clean all items sent out.
            // Delete the item in Inbox folder and clear the ExistItemIds list.
            DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);

            this.InitializeCollection();

            // Clear the soap header.
            this.ClearSoapHeaders();

            // Clean the items sent out.
            this.CleanItemsSentOut(new string[] { items[0].Subject });
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the IsDraft and IsUnmodified Boolean values for base item with successful responses.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC21_VerifyItemWithIsDraftAndIsUnmodified()
        {
            #region Step 1: Create the item which will be saved.
            // Define the MessageType item to create.
            EmailAddressType addressTo = new EmailAddressType();
            addressTo.EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            EmailAddressType addressCc = new EmailAddressType();
            addressCc.EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            MessageType[] message = new MessageType[1]
            {
                new MessageType
                {
                    ToRecipients = new EmailAddressType[] { addressTo, addressCc },
                    CcRecipients = new EmailAddressType[] { addressCc, addressTo },
                    Subject = Common.GenerateResourceName(this.Site, "ItemSaveOnly")
                }
            };

            // Call the CreateItem operation.
            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.inbox, message);

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

            #region Step 2: Get the created item.
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

            // If the item is saved, the IsDraft element should be true and the IsUnmodified element should be false.
            ItemInfoResponseMessageType itemInfoResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            ArrayOfRealItemsType arrayOfRealItemsType = itemInfoResponseMessage.Items;

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1608, Expected result:{0}, Actual result:{1}", true, arrayOfRealItemsType.Items[0].IsDraft);

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1608
            bool isVerifiedR1608 = arrayOfRealItemsType.Items[0].IsDraftSpecified && arrayOfRealItemsType.Items[0].IsDraft;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1608,
                1608,
                @"[In t:ItemType Complex Type] [IsDraft is] True, indicates an item has not been sent.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1615, Expected result:{0}, Actual result:{1}", false, arrayOfRealItemsType.Items[0].IsUnmodified);

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1615
            bool isVerifiedR1615 = arrayOfRealItemsType.Items[0].IsUnmodifiedSpecified && !arrayOfRealItemsType.Items[0].IsUnmodified;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1615,
                1615,
                @"[In t:ItemType Complex Type] otherwise [IsUnmodified is] false, indicates [an item has been modified].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R96");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R96
            // The schema is validated, the DisplayCc element is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                arrayOfRealItemsType.Items[0].DisplayCc,
                96,
                @"[In t:ItemType Complex Type] [The element ""DisplayCc""] Specifies the display string that is used for the contents of the Cc box.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R97");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R97
            // The schema is validated, the DisplayCc element is concatenated by the display names of all Cc recipients, so this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<string>(
                (Common.GetConfigurationPropertyValue("User2Name", this.Site) + "; " + Common.GetConfigurationPropertyValue("User1Name", this.Site)).ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                arrayOfRealItemsType.Items[0].DisplayCc.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                97,
                @"[In t:ItemType Complex Type] This [DisplayCc] is the concatenated string of all Cc recipient display names.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R98");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R98
            // The schema is validated, the DisplayTo element is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                arrayOfRealItemsType.Items[0].DisplayTo,
                98,
                @"[In t:ItemType Complex Type] [The element ""DisplayTo""] Specifies the display string that is used for the contents of the To box.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R99");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R99
            // The schema is validated, the DisplayTo element is concatenated by the display names of all To recipients, so this requirement can be captured.
            this.Site.CaptureRequirementIfAreEqual<string>(
                (Common.GetConfigurationPropertyValue("User1Name", this.Site) + "; " + Common.GetConfigurationPropertyValue("User2Name", this.Site)).ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                arrayOfRealItemsType.Items[0].DisplayTo.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                99,
                @"[In t:ItemType Complex Type] This [DisplayTo] is the concatenated string of all To recipient display names.");
            #endregion

            #region Step 3: Create the item which will be sent.
            // Define the CreateItem request.
            CreateItemType requestItem = new CreateItemType();
            requestItem.MessageDispositionSpecified = true;
            requestItem.MessageDisposition = MessageDispositionType.SendOnly;
            requestItem.Items = new NonEmptyArrayOfAllItemsType();
            message = new MessageType[1]
            {
                new MessageType
                {
                    ToRecipients = new EmailAddressType[] { addressTo },
                    Subject = Common.GenerateResourceName(this.Site, "ItemSendOnly")
                }
            };
            requestItem.Items.Items = message;

            // Call the CreateItem operation.
            createItemResponse = this.COREAdapter.CreateItem(requestItem);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 4: Get the received item.
            // Find the received item in the Inbox folder.
            ItemIdType[] findItem = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, message[0].Subject, "User1");

            // The result of FindItemsInFolder should not be null.
            Site.Assert.IsNotNull(
                findItem,
                "The result of FindItemsInFolder should not be null.");

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 findItem.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 findItem.GetLength(0));

            // Get the found item.
            getItemResponse = this.CallGetItemOperation(findItem);

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

            // If the item is sent, the IsDraft element should be false and the IsUnmodified element should be true.
            itemInfoResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            arrayOfRealItemsType = itemInfoResponseMessage.Items;

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1609, Expected result:{0}, Actual result:{1}", false, arrayOfRealItemsType.Items[0].IsDraft);

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1609
            bool isVerifiedR1609 = arrayOfRealItemsType.Items[0].IsDraftSpecified && !arrayOfRealItemsType.Items[0].IsDraft;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1609,
                1609,
                @"[In t:ItemType Complex Type] otherwise [IsDraft is] false, indicates [an item has been sent].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1614, Expected result:{0}, Actual result:{1}", true, arrayOfRealItemsType.Items[0].IsUnmodified);

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1614
            bool isVerifiedR1614 = arrayOfRealItemsType.Items[0].IsUnmodifiedSpecified && arrayOfRealItemsType.Items[0].IsUnmodified;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1614,
                1614,
                @"[In t:ItemType Complex Type] [IsUnmodified is] True, indicates an item has not been modified.");
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate the EntityExtractionResult value for base item from successful response is consistent with request value.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC22_VerifyItemWithEntityExtractionResult()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(1708, this.Site), "Exchange 2007 and Exchange 2010 do not support the EntityExtractionResultType complex type.");

            #region Step 1: Create message with specific body.

            // Define the specific elements' values.
            string meetingSuggestions = "Let's meet for business discussion, from 2:00pm to 2:30pm, December 15th, 2012.";
            string address = "1234 Main Street, Redmond, WA 07722";
            string taskSuggestions = "Please update the spreadsheet by today.";
            string phoneNumbers = "(235) 555-0110";
            string phoneNumberType = "Home";
            string businessName = "Department of Revenue Services";
            string contactDisplayName = TestSuiteHelper.ContactString;
            string emailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            string url = Common.GetConfigurationPropertyValue("ServiceUrl", this.Site);
            string userEmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Define the MessageType item which will be created.
            MessageType[] messages = new MessageType[] { new MessageType() };
            messages[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem);
            EmailAddressType email = new EmailAddressType();
            email.EmailAddress = userEmailAddress;
            messages[0].ToRecipients = new EmailAddressType[1];
            messages[0].ToRecipients[0] = email;
            messages[0].Body = new BodyType();
            messages[0].Body.Value = string.Format(
                "{0} {1} Any problems, contact with {2} from {3}, his {4} phone number is {5}, his email is {6}, his blog is {7} and his address is {8}",
                meetingSuggestions,
                taskSuggestions,
                contactDisplayName,
                businessName,
                phoneNumberType,
                phoneNumbers,
                emailAddress,
                url,
                address);

            // Define the request of CreateItem operation.
            CreateItemType requestItem = new CreateItemType();
            requestItem.MessageDispositionSpecified = true;
            requestItem.MessageDisposition = MessageDispositionType.SendOnly;
            requestItem.Items = new NonEmptyArrayOfAllItemsType();
            requestItem.Items.Items = messages;

            // Call the CreateItem operation.
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(requestItem);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);
            #endregion

            #region Step 2: Find the message in the inbox folder of User1.
            // Find the received item in the Inbox folder.
            ItemIdType[] foundItems = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, messages[0].Subject, "User1");

            // The result of FindItemsInFolder should not be null.
            Site.Assert.IsNotNull(
                foundItems,
                "The result of FindItemsInFolder should not be null.");

            // One item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 foundItems.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 foundItems.GetLength(0));
            #endregion

            #region Step 3: Get the message.
            // Get information from the found items.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(foundItems);

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

            // Check whether the child elements of EntityExtractionResultType have been returned successfully.
            ItemInfoResponseMessageType getItems = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;

            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1325");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1325
            // The schema is validated and InternetMessageHeaders is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                getItems.Items.Items[0].InternetMessageHeaders,
                1325,
                @"[In t:ItemType Complex Type] The type of InternetMessageHeaders is t:NonEmptyArrayOfInternetHeadersType (section 2.2.4.12).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R122");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R122
            // The schema is validated and InternetMessageHeaders is not null, so this requirement can be captured.
            this.Site.CaptureRequirement(
                122,
                @"[In t:NonEmptyArrayOfInternetHeadersType Complex Type] The type [NonEmptyArrayOfInternetHeadersType] is defined as follow:
<xs:complexType name=""NonEmptyArrayOfInternetHeadersType"">
  <xs:sequence>
    <xs:element name=""InternetMessageHeader""
      type=""t:InternetHeaderType""
      maxOccurs=""unbounded""
     />
  </xs:sequence>
</xs:complexType>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R20301");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R20301
            // MS-OXWSCORE_R1325 is captured, this requirement can be captured directly.
            this.Site.CaptureRequirement(
                20301,
                @"[In t:ItemType Complex Type] It [InternetMessageHeaders] can be retrieved by GetItem (section 3.1.4.4) operation.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1367");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1367
            // The schema is validated and InternetMessageHeaders is not null, so this requirement can be captured.
            this.Site.CaptureRequirement(
                1367,
                @"[In t:NonEmptyArrayOfInternetHeadersType Complex Type] The type of InternetMessageHeader is t:InternetHeaderType([MS-OXWSCDATA] section 2.2.4.35).");

            foreach (InternetHeaderType internetMessageHeader in getItems.Items.Items[0].InternetMessageHeaders)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1674");

                // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1674
                // The schema is validated and the internetMessageHeader is not null, so this requirement can be captured.
                this.Site.CaptureRequirementIfIsNotNull(
                    internetMessageHeader,
                    "MS-OXWSCDATA",
                    1674,
                    @"[In t:InternetHeaderType Complex Type] The attribute ""HeaderName"" is ""xs:string"" type.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R88");
                this.Site.Log.Add(LogEntryKind.Debug, "The HeaderName should not be null, actual {0}.", internetMessageHeader.HeaderName);
                this.Site.Log.Add(LogEntryKind.Debug, "The Value should not be null, actual {0}.", internetMessageHeader.Value);

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R88
                // The schema is validated and the child elements in internetMessageHeader are not null, so this requirement can be captured.
                bool isVerifiedR88 = internetMessageHeader.HeaderName != null && internetMessageHeader.Value != null;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR88,
                    88,
                    @"[In t:ItemType Complex Type] [The element ""InternetMessageHeaders""] Specifies an array of the type InternetHeaderType that represents the collection of all Internet message headers that are contained in an item in a mailbox.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R124");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R124
                this.Site.CaptureRequirement(
                    124,
                    @"[In t:NonEmptyArrayOfInternetHeadersType Complex Type] [The element ""InternetMessageHeader""] Specifies a single Internet message header.");
            }

            EntityExtractionResultType entityExtractionResult = getItems.Items.Items[0].EntityExtractionResult;

            // Verify EntityExtractionResultType structure.
            this.VerifyEntityExtractionResultType(entityExtractionResult);

            // Verify ArrayOfAddressEntitiesType structure.
            this.VerifyArrayOfAddressEntitiesType(entityExtractionResult.Addresses, address);

            // Verify ArrayOfMeetingSuggestionsType structure.
            // "2:00pm, December 15th, 2012" is supposed to be "2012/12/15 14:00:00" in dataTime type.
            // "2:30pm, December 15th, 2012" is supposed to be "2012/12/15 14:30:00" in dataTime type.
            this.VerifyArrayOfMeetingSuggestionsType(entityExtractionResult.MeetingSuggestions, meetingSuggestions, DateTime.Parse("2012/12/15 14:00:00"), DateTime.Parse("2012/12/15 14:30:00"), Common.GetConfigurationPropertyValue("User1Name", this.Site), userEmailAddress);

            // Verify ArrayOfTaskSuggestionsType structure.
            this.VerifyArrayOfTaskSuggestionsType(entityExtractionResult.TaskSuggestions, taskSuggestions, Common.GetConfigurationPropertyValue("User1Name", this.Site), userEmailAddress);

            // Verify ArrayOfEmailAddressEntitiesType structure.
            this.VerifyArrayOfEmailAddressEntitiesType(entityExtractionResult.EmailAddresses, emailAddress);

            // Verify ArrayOfContactsType structure.
            Uri uri = new Uri(url);
            this.VerifyArrayOfContactsType(entityExtractionResult.Contacts, contactDisplayName, businessName, uri, phoneNumbers, phoneNumberType, emailAddress, address);

            // Verify ArrayOfUrlEntitiesType structure.
            this.VerifyArrayOfUrlEntitiesType(entityExtractionResult.Urls, uri);

            // Verify ArrayOfPhoneEntitiesType structure.
            this.VerifyArrayOfPhoneEntitiesType(entityExtractionResult.PhoneNumbers, phoneNumbers, phoneNumberType);
            #endregion
        }

        /// <summary>
        /// This test case is intended to validate PublicFolderItem of IdStorageType with the successful response returned by CreateItem operations for a message.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC23_VerifyPublicFolderItem()
        {
            FolderIdType folderId = null;

            FindFolderType findRequest = new FindFolderType();
            findRequest.FolderShape = new FolderResponseShapeType();
            findRequest.FolderShape.BaseShape = DefaultShapeNamesType.AllProperties;
            DistinguishedFolderIdType id = new DistinguishedFolderIdType();
            id.Id = DistinguishedFolderIdNameType.publicfoldersroot;
            findRequest.ParentFolderIds = new DistinguishedFolderIdType[] { id };
            findRequest.Traversal = FolderQueryTraversalType.Shallow;

            BaseFolderType[] folders = null;
            FindFolderResponseType findFolderResponse = this.SRCHAdapter.FindFolder(findRequest);
            FindFolderResponseMessageType findFolderResponseMessageType = new FindFolderResponseMessageType();
            if (findFolderResponse != null && findFolderResponse.ResponseMessages.Items[0].ResponseClass == ResponseClassType.Success)
            {
                findFolderResponseMessageType = findFolderResponse.ResponseMessages.Items[0] as FindFolderResponseMessageType;
                folders = findFolderResponseMessageType.RootFolder.Folders;
                foreach (BaseFolderType folder in folders)
                {
                    if (folder.DisplayName.Equals(Common.GetConfigurationPropertyValue("PublicFolderName", this.Site)))
                    {
                        folderId = folder.FolderId;
                    }
                }
            }

            Site.Assert.IsNotNull(
                folderId,
                "The destination public folder {0} in should exist!",
                Common.GetConfigurationPropertyValue("PublicFolderName", this.Site));

            MessageType message = new MessageType();
            MessageType[] items = new MessageType[] { message };
            CreateItemType requestItem = new CreateItemType();
            requestItem.MessageDispositionSpecified = true;
            requestItem.MessageDisposition = MessageDispositionType.SaveOnly;
            requestItem.Items = new NonEmptyArrayOfAllItemsType();
            requestItem.Items.Items = items;

            requestItem.SavedItemFolderId = new TargetFolderIdType();

            requestItem.SavedItemFolderId.Item = folderId;

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(requestItem);

            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemInfoResponseMessageType itemInfo = (ItemInfoResponseMessageType)createItemResponse.ResponseMessages.Items[0];
            ItemIdId itemId = this.ITEMIDAdapter.ParseItemId(itemInfo.Items.Items[0].ItemId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R62");

            // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R62
            Site.CaptureRequirementIfAreEqual<IdStorageType>(
                IdStorageType.PublicFolderItem,
                itemId.StorageType,
                "MS-OXWSITEMID",
                62,
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
                        PublicFolder = 1,]

                        /// <summary>
                        /// The Id represents an item in a PublicFolder store.
                        /// </summary>
                        PublicFolderItem = 2,
                [
                        /// <summary>
                        /// The Id represents an item or folder in a mailbox and contains a mailbox GUID.
                        /// </summary>
                        MailboxItemMailboxGuidBased = 3,

                        /// <summary>
                        /// The Id represents a conversation in a mailbox and contains a mailbox GUID.
                        /// </summary>
                        ConversationIdMailboxGuidBased = 4,

                        /// <summary>
                        /// The Id represents (by objectGuid) an object in the Active Directory.
                        /// </summary>
                        ActiveDirectoryObject = 5,]
                }");
        }

        /// <summary>
        /// This case is intended to validate to delete an item with setting SuppressReadReceipts successfully.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC24_DeleteItemWithSuppressReadReceipts()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2311, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the SuppressReadReceipts attribute.");

            #region Send an email with setting IsReadReceiptRequested to true.

            MessageType message = new MessageType();
            message.IsReadReceiptRequestedSpecified = true;
            message.IsReadReceiptRequested = true;
            message.ToRecipients = new EmailAddressType[1];
            EmailAddressType recipient = new EmailAddressType();
            recipient.EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            message.ToRecipients[0] = recipient;
            message.From = new SingleRecipientType
            {
                Item = new EmailAddressType
                {
                    EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site)
                }
            };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { message };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.MessageDisposition = MessageDispositionType.SendOnly;
            createItemRequest.MessageDispositionSpecified = true;
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            #endregion

            #region Find the email in receiver's inbox.

            ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User2");
            Site.Assert.IsNotNull(findItemIds, "The receiver should receive the email.");

            #endregion

            #region Delete the found email with setting SuppressReadReceipts to true.

            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.ItemIds = findItemIds;
            deleteItemRequest.DeleteType = DisposalType.HardDelete;
            deleteItemRequest.SuppressReadReceiptsSpecified = true;
            deleteItemRequest.SuppressReadReceipts = true;
            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);
            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User1");
            Site.Assert.IsNull(findItemIds, "The read receipt email should not be received if receiver delete the email with setting SuppressReadReceipts to true.");

            #endregion

            #region Send an email with setting IsReadReceiptRequested to true.

            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            #endregion

            #region Find the email in receiver's inbox.

            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User2");
            Site.Assert.IsNotNull(findItemIds, "The receiver should receive the email.");

            #endregion

            #region Delete the found email with setting SuppressReadReceipts to false.

            deleteItemRequest.ItemIds = findItemIds;
            deleteItemRequest.SuppressReadReceipts = false;
            deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);
            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User1");
            Site.Assert.AreEqual<int>(1, findItemIds.Length, "The read receipt email should be received if receiver delete the email with setting SuppressReadReceipts to false.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2311");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2311
            // This requirement can be captured directly after above steps.
            this.Site.CaptureRequirement(
                2311,
                @"[In Appendix C: Product Behavior] Implementation does support the SuppressReadReceipts attribute which specifies whether read receipts are suppressed. (Exchange 2013 SP1 and above follow this behavior.)");

            #endregion
        }

        /// <summary>
        /// This case is intended to validate to update an item with setting SuppressReadReceipts successfully.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S04_TC25_UpdateItemWithSuppressReadReceipts()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(2315, this.Site), "Exchange 2007, Exchange 2010, and the initial release of Exchange 2013 do not support the SuppressReadReceipts attribute.");

            #region Send an email with setting IsReadReceiptRequested to true.

            MessageType message = new MessageType();
            message.IsReadReceiptRequestedSpecified = true;
            message.IsReadReceiptRequested = true;
            message.ToRecipients = new EmailAddressType[1];
            EmailAddressType recipient = new EmailAddressType();
            recipient.EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            message.ToRecipients[0] = recipient;
            message.From = new SingleRecipientType
            {
                Item = new EmailAddressType
                {
                    EmailAddress = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site)
                }
            };
            CreateItemType createItemRequest = new CreateItemType();
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { message };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 1);
            createItemRequest.MessageDisposition = MessageDispositionType.SendOnly;
            createItemRequest.MessageDispositionSpecified = true;
            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            #endregion

            #region Find the email in receiver's inbox.

            ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User2");
            Site.Assert.IsNotNull(findItemIds, "The receiver should receive the email.");

            #endregion

            #region Update the found email with setting SuppressReadReceipts to true.

            ItemChangeType[] itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();
            itemChanges[0].Item = findItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItemFiled = new SetItemFieldType();
            setItemFiled.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.messageIsRead
            };
            setItemFiled.Item1 = new MessageType()
            {
                IsRead = true,
                IsReadSpecified = true
            };
            itemChanges[0].Updates[0] = setItemFiled;
            UpdateItemType updateItemType = new UpdateItemType();
            updateItemType.ItemChanges = itemChanges;
            updateItemType.ConflictResolution = ConflictResolutionType.AutoResolve;
            updateItemType.MessageDisposition = MessageDispositionType.SaveOnly;
            updateItemType.MessageDispositionSpecified = true;
            updateItemType.SuppressReadReceipts = true;
            updateItemType.SuppressReadReceiptsSpecified = true;
            UpdateItemResponseType updateItemResponse = this.COREAdapter.UpdateItem(updateItemType);
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);
            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User1");
            Site.Assert.IsNull(findItemIds, "The read receipt email should not be received if receiver update the email with setting SuppressReadReceipts to true.");

            List<string> subjects = new List<string>();
            subjects.Add(createItemRequest.Items.Items[0].Subject);
            this.ExistItemIds.Clear();
            #endregion

            #region Send an email with setting IsReadReceiptRequested to true.

            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem, 2);
            createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            #endregion

            #region Find the email in receiver's inbox.

            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User2");
            Site.Assert.IsNotNull(findItemIds, "The receiver should receive the email.");

            #endregion

            #region Update the found email with setting SuppressReadReceipts to false.

            updateItemType.ItemChanges[0].Item = findItemIds[0];
            updateItemType.SuppressReadReceipts = false;
            updateItemResponse = this.COREAdapter.UpdateItem(updateItemType);
            Common.CheckOperationSuccess(updateItemResponse, 1, this.Site);
            findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, createItemRequest.Items.Items[0].Subject, "User1");
            Site.Assert.AreEqual<int>(1, findItemIds.Length, "The read receipt email should not be received if receiver update the email with setting SuppressReadReceipts to true.");
            subjects.Add(createItemRequest.Items.Items[0].Subject);
            this.ExistItemIds.Clear();
            this.ExistItemIds.Add(findItemIds[0]);
            this.CleanItemsSentOut(subjects.ToArray());

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2315");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2315
            // This requirement can be captured directly after above steps.
            this.Site.CaptureRequirement(
                2315,
                @"[In Appendix C: Product Behavior] Implementation does  support the SuppressReadReceipts attribute specifies whether read receipts are suppressed. (<113> Section 3.1.4.9.3.2:  This attribute [SuppressReadReceipts] was introduced in Exchange 2013 SP1.)");
            #endregion
        }

        #endregion
    }
}