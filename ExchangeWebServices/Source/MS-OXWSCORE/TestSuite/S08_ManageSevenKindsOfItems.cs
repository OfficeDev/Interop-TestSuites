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
    using System.Collections.Generic;
    using System.Reflection;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is designed to test operations related to creation, retrieving, updating, movement, copy, and deletion of base, contact, distribution list, email, meeting, post, and task items on the server.
    /// </summary>
    [TestClass]
    public class S08_ManageSevenKindsOfItems : TestSuiteBase
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
        /// This test case is intended to validate the successful response returned by CreateItem, GetItem and DeleteItem operations for multiple types of items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S08_TC01_CreateGetDeleteTypesOfItemsSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create Items.
            ItemIdType[] createdItemIds = CreateAllTypesItems();
            #endregion

            #region Step 2: Get items

            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 7, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            #endregion

            #region Step3: Delete the item
            DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 7, this.Site);

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 4: Get deleted items

            getItemResponse = this.CallGetItemOperation(getItemIds);

            Site.Assert.AreEqual<int>(
                 7,
                 getItemResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 7,
                 getItemResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the GetItem operation is executed failed with ErrorItemNotFound response code.
            foreach (ResponseMessageType responseMessage in getItemResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Error,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Get each types of items should succeed! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.ErrorItemNotFound,
                            responseMessage.ResponseCode));
            }

            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem and CopyItem operations for multiple types of items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S08_TC02_CopyTypesOfItemsSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create Items.

            // Initialize items data.
            object obj;
            List<ItemType> items = new List<ItemType>();
            BaseItemIdType[] createdItemIds;

            // Get the ItemType and six extend types which base on ItemType.
            Assembly assembly = Assembly.GetAssembly(typeof(ItemType));
            Type[] types = assembly.GetTypes();

            // Initialize the public properties (Subject and Body) which the seven kinds of operation both have.
            PropertyInfo subjectField;
            PropertyInfo bodyField;
            string subject = TestSuiteHelper.SubjectForCreateItem;
            BodyType body = new BodyType()
            {
                Value = TestSuiteHelper.BodyForBaseItem,
                BodyType1 = BodyTypeType.Text
            };

            // Set the Subject and Body properties for each type.
            foreach (Type type in types)
            {
                if ((type.BaseType == typeof(ItemType) || type == typeof(ItemType)) && !type.IsAbstract)
                {
                    string typeName = type.ToString();
                    obj = assembly.CreateInstance(typeName);
                    subjectField = type.GetProperty("Subject");
                    if (subjectField != null)
                    {
                        subjectField.SetValue(obj, Common.GenerateResourceName(this.Site, subject + type.Name), null);
                    }

                    bodyField = type.GetProperty("Body");
                    if (bodyField != null)
                    {
                        bodyField.SetValue(obj, body, null);
                    }

                    items.Add((ItemType)obj);
                }
            }

            ItemType[] itemTypes = items.ToArray();
            CreateItemType createRequest = new CreateItemType()
            {
                Items = new NonEmptyArrayOfAllItemsType()
                {
                    Items = itemTypes
                }
            };

            createRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            createRequest.MessageDispositionSpecified = true;
            createRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
            createRequest.SendMeetingInvitationsSpecified = true;

            // Call CreateItem to create seven items that contains Subject and Body public elements in the Inbox folder on the server.
            CreateItemResponseType createResponse = this.COREAdapter.CreateItem(createRequest);

            // Get the create item Ids.
            createdItemIds = Common.GetItemIdsFromInfoResponse(createResponse);

            // Check the operation response.
            Common.CheckOperationSuccess(createResponse, 7, this.Site);

            #endregion

            #region Step 2: Copy items.

            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 7, this.Site);
            #endregion 

            this.FindNewItemsInFolder(DistinguishedFolderIdNameType.drafts);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem and MoveItem operations for multiple types of items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S08_TC03_MoveTypesOfItemsSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create Items.
            ItemIdType[] createdItemIds = CreateAllTypesItems();
            #endregion

            #region Step 2: Move items.
            // Clear ExistItemIds for MoveItem
            this.InitializeCollection();

            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.deleteditems, createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 7, this.Site);

            #endregion 

            this.FindNewItemsInFolder(DistinguishedFolderIdNameType.deleteditems);
        }

        /// <summary>
        /// This test case is intended to validate the successful response returned by CreateItem and UpdateItem operations for multiple types of items with required elements.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S08_TC04_UpdateTypesOfItemsSuccessfully()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create Items.
            ItemIdType[] createdItemIds = CreateAllTypesItems();
            #endregion

            #region Step 2: Update items.

            ItemChangeType[] itemChanges = new ItemChangeType[createdItemIds.Length];

            // Set the public properties (Subject) which all the seven kinds of operation have.
            for (int i = 0; i < createdItemIds.Length; i++)
            {
                itemChanges[i] = new ItemChangeType();
                itemChanges[i].Item = createdItemIds[i];
                itemChanges[i].Updates = new ItemChangeDescriptionType[]
                    {
                        new SetItemFieldType()
                        {
                            Item = new PathToUnindexedFieldType()
                            {
                                FieldURI = UnindexedFieldURIType.itemSubject
                            },

                            Item1 = new ItemType()
                            {
                               Subject = Common.GenerateResourceName(
                                            this.Site,
                                            TestSuiteHelper.SubjectForUpdateItem)
                            }
                        }
                    };
            }

            UpdateItemResponseType updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 7, this.Site);

            #endregion 
        }

        /// <summary>
        /// This test case is intended to validate the failed response returned by UpdateItem operation with ErrorIncorrectUpdatePropertyCount response code for multiple types of items.
        /// </summary>
        [TestCategory("MSOXWSCORE"), TestMethod()]
        public void MSOXWSCORE_S08_TC05_UpdateTypesOfItemsFailed()
        {
            Site.Assume.IsTrue(Common.IsRequirementEnabled(19241, this.Site), "Exchange 2007 doesn't support MS-OXWSDLIST");

            #region Step 1: Create Items.
            ItemIdType[] createdItemIds = CreateAllTypesItems();
            #endregion

            #region Step 2: Update items.
            // Initialize the change item to update.
            UpdateItemType updateRequest = new UpdateItemType();
            ItemChangeType[] itemChanges = new ItemChangeType[createdItemIds.Length];

            // Set two properties (Subject and ReminderMinutesBeforeStart) to update, in order to return an error "ErrorIncorrectUpdatePropertyCount".
            for (int i = 0; i < createdItemIds.Length; i++)
            {
                itemChanges[i] = new ItemChangeType();
                itemChanges[i].Item = createdItemIds[i];
                itemChanges[i].Updates = new ItemChangeDescriptionType[1];
                SetItemFieldType setItem1 = new SetItemFieldType();
                setItem1.Item = new PathToUnindexedFieldType()
                {
                    FieldURI = UnindexedFieldURIType.itemSubject
                };
                setItem1.Item1 = new ContactItemType()
                {
                    Subject = Common.GenerateResourceName(
                        this.Site,
                        TestSuiteHelper.SubjectForUpdateItem),
                    ReminderMinutesBeforeStart = TestSuiteHelper.ReminderMinutesBeforeStart
                };
                itemChanges[i].Updates[0] = setItem1;
            }

            updateRequest.ItemChanges = itemChanges;
            updateRequest.MessageDispositionSpecified = true;
            updateRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            updateRequest.SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType.SendToAllAndSaveCopy;
            updateRequest.SendMeetingInvitationsOrCancellationsSpecified = true;

            // Call UpdateItem to update the Subject and the ReminderMinutesBeforeStart of the created item simultaneously.
            UpdateItemResponseType updateItemResponse = this.COREAdapter.UpdateItem(updateRequest);

            foreach (ResponseMessageType responseMessage in updateItemResponse.ResponseMessages.Items)
            {
                // Verify ResponseCode is ErrorIncorrectUpdatePropertyCount.
                this.VerifyErrorIncorrectUpdatePropertyCount(responseMessage.ResponseCode);
            }
            #endregion
        }
        #endregion
    }
}