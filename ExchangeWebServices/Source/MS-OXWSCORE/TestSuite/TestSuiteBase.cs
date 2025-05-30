namespace Microsoft.Protocols.TestSuites.MS_OXWSCORE
{
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.Reflection;
    using System.Text.RegularExpressions;
    using System.Xml;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestSuites.MS_OXWSITEMID;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The bass class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// Define a collection used to store existing items.
        /// </summary>
        private Collection<ItemIdType> existItemIds;

        /// <summary>
        /// Define a collection used to store copied items.
        /// </summary>
        private Collection<ItemIdType> copiedItemIds;
        #endregion

        #region Properties
        /// <summary>
        /// Gets the list to store and access the existing itemIds.
        /// The adding of new itemIds is called in ExchangeServiceBinding_ResponseEvent method,
        /// The removal of obsolete itemIds is manually done after deleteItem, SendItem and MoveItem operation.
        /// </summary>
        protected Collection<ItemIdType> ExistItemIds
        {
            get { return this.existItemIds; }
        }

        /// <summary>
        /// Gets the collection to store and access the copied itemIds.
        /// </summary>
        protected Collection<ItemIdType> CopiedItemIds
        {
            get { return this.copiedItemIds; }
        }

        /// <summary>
        /// Gets the MS-OXWSCORE Adapter.
        /// </summary>
        protected IMS_OXWSCOREAdapter COREAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSCORE SUT Control Adapter.
        /// </summary>
        protected IMS_OXWSCORESUTControlAdapter CORESUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSITEMID Adapter.
        /// </summary>
        protected IMS_OXWSITEMIDAdapter ITEMIDAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSSRCH Adapter Instance
        /// </summary>
        protected IMS_OXWSSRCHAdapter SRCHAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSUSRCFG SUT control adapter.
        /// </summary>
        protected IMS_OXWSUSRCFGSUTControlAdapter USRCFGSUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets the last response get from server.
        /// </summary>
        protected BaseResponseMessageType LastResponse { get; private set; }

        /// <summary>
        /// Gets a value indicating whether the schema validation is successful.
        /// </summary>
        protected bool IsSchemaValidated { get; private set; }

        #endregion

        #region Test case initialize and clean up

        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();
            ExchangeServiceBinding.ServiceResponseEvent += new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);
            this.InitializeCollection();
            this.COREAdapter = Site.GetAdapter<IMS_OXWSCOREAdapter>();
            this.CORESUTControlAdapter = this.Site.GetAdapter<IMS_OXWSCORESUTControlAdapter>();
            this.SRCHAdapter = Site.GetAdapter<IMS_OXWSSRCHAdapter>();
            this.ITEMIDAdapter = Site.GetAdapter<IMS_OXWSITEMIDAdapter>();
            this.USRCFGSUTControlAdapter = Site.GetAdapter<IMS_OXWSUSRCFGSUTControlAdapter>();
            this.ClearSoapHeaders();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            ExchangeServiceBinding.ServiceResponseEvent -= new ExchangeServiceBinding.ServiceResponseDelegate(this.ExchangeServiceBinding_ResponseEvent);

            if (this.ExistItemIds != null && this.ExistItemIds.Count > 0)
            {
                DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

                foreach (ResponseMessageType messageType in deleteItemResponse.ResponseMessages.Items)
                {
                    Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        messageType.ResponseClass,
                        string.Format(
                            "Delete item should succeed! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            messageType.ResponseCode));
                }
            }

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();

            base.TestCleanup();
        }
        #endregion

        #region Test case base methods
        /// <summary>
        /// Initialize the collections.
        /// </summary>
        protected void InitializeCollection()
        {
            // Initialize the collection of existing items.
            if (this.existItemIds != null)
            {
                this.existItemIds.Clear();
            }
            else
            {
                this.existItemIds = new Collection<ItemIdType>();
            }

            // Initialize the collection of copied items.
            if (this.copiedItemIds != null)
            {
                this.copiedItemIds.Clear();
            }
            else
            {
                this.copiedItemIds = new Collection<ItemIdType>();
            }
        }

        /// <summary>
        /// Create an item with all properties.
        /// </summary>
        /// <returns>The item object</returns>
        protected ItemType CreateFullPropertiesItem()
        {
            ItemType item = new ItemType();

            // Set the ItemClass as "IPM.Note".
            item.ItemClass = "IPM.Note";
            item.Subject = Common.GenerateResourceName(
               this.Site,
               TestSuiteHelper.SubjectForCreateItem);
            item.SensitivitySpecified = true;
            item.Sensitivity = SensitivityChoicesType.Normal;
            item.Body = new BodyType();
            item.Body.BodyType1 = BodyTypeType.Text;
            item.Body.Value = TestSuiteHelper.BodyForBaseItem;

            item.Categories = new string[1];
            item.Categories[0] = TestSuiteHelper.CategoryName;
            item.ImportanceSpecified = true;
            item.Importance = ImportanceChoicesType.Normal;
            item.InReplyTo = TestSuiteHelper.InReplyTo;
            item.ReminderDueBySpecified = true;
            item.ReminderDueBy = DateTime.Now.AddMinutes(15);
            item.ReminderIsSetSpecified = true;
            item.ReminderIsSet = true;
            item.ReminderMinutesBeforeStart = TestSuiteHelper.ReminderMinutesBeforeStart;
            item.ExtendedProperty = new ExtendedPropertyType[1];
            item.ExtendedProperty[0] = new ExtendedPropertyType();
            item.ExtendedProperty[0].ExtendedFieldURI = new PathToExtendedFieldType();

            // Set the extend properties for the element.
            DistinguishedPropertySetType distinguishedPropertySetId = DistinguishedPropertySetType.Common;

            int propertyId = Convert.ToInt32(TestSuiteHelper.PropertyId);
            MapiPropertyTypeType propertyType = MapiPropertyTypeType.String;

            item.ExtendedProperty[0].ExtendedFieldURI.DistinguishedPropertySetId = distinguishedPropertySetId;
            item.ExtendedProperty[0].ExtendedFieldURI.DistinguishedPropertySetIdSpecified = true;
            item.ExtendedProperty[0].ExtendedFieldURI.PropertyId = propertyId;
            item.ExtendedProperty[0].ExtendedFieldURI.PropertyIdSpecified = true;
            item.ExtendedProperty[0].ExtendedFieldURI.PropertyType = propertyType;
            item.ExtendedProperty[0].Item = TestSuiteHelper.ElementValue;
            item.Culture = TestSuiteHelper.Culture;

            // Add Exchange 2013 elements.
            if (Common.IsRequirementEnabled(4003, this.Site))
            {
                item.Body.IsTruncatedSpecified = true;
                item.Body.IsTruncated = true;
            }

            if (Common.IsRequirementEnabled(1271, this.Site))
            {
                item.Flag = new FlagType();
                item.Flag.FlagStatus = FlagStatusType.Flagged;
                item.Flag.StartDateSpecified = true;
                item.Flag.StartDate = DateTime.Now;
                item.Flag.DueDateSpecified = true;
                item.Flag.DueDate = DateTime.Now.AddDays(1);
            }

            if (Common.IsRequirementEnabled(1353, this.Site))
            {
                item.RetentionDateSpecified = true;
                item.RetentionDate = DateTime.Now.AddDays(1);
            }

            if (Common.IsRequirementEnabled(2281, this.Site))
            {
                FileAttachmentType fileAttachment = new FileAttachmentType();
                fileAttachment.Name = Common.GenerateResourceName(this.Site, "File attachment name");
                fileAttachment.Content = Convert.FromBase64String("/9j/4AAQSkZJRgABAQEAYABgAAD/7AARRHVja3kAAQAEAAAARgAA/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgADAAUAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A5j9hnXPH/gy003RdI+KejWVjciO7muPF86y2GiWFu3nahceUJUyEtfOcKroplMZeSJAxH1j+2p8f9L8TfBTw1qvwo8YSSWGr6FB4p1bStWsLi01TVPDl4ZLe21CwkeOJ9yXf2cSQlCfKukdjDvhM34beLf2n/G3wz+B2g+ItA1qbS9Y0bxBm1uIlDFA9vPG6kNkFWQlSp4welZfxG/4K1fHj9ozQj4T8V+N73U9J1u7e7uxLmWU+YYWaGJ5Cxt7fzLeF/s8Hlxbox8mABXNQjVjQnQlVk+zurrRbWSXnbbuedTyvD0ISoQvZ333WnyPqLxp8PG8feIrjU5dG8O60ZWMYuDeW1tjyyUKeWzALtZWGFG3Oe+aK/OP4ja3c+CfiZ4m07TJTbWdvq90iIQJSAsrKMs+WJwo6miuFZRbT2svvX+QoZTQirXf4f5H/2Q==");

                ItemAttachmentType itemAttachment = new ItemAttachmentType();
                itemAttachment.Name = Common.GenerateResourceName(this.Site, "Item attachment name");
                itemAttachment.Item = new ItemType();
                itemAttachment.Item.Subject = Common.GenerateResourceName(this.Site, "Item attachment subject");
                item.Attachments = new AttachmentType[2];

                item.Attachments[0] = fileAttachment;
                item.Attachments[1] = itemAttachment;
            }

            if (Common.IsRequirementEnabled(2283, this.Site))
            {
                item.ReminderNextTimeSpecified = true;
                item.ReminderNextTime = DateTime.Now.AddMinutes(30);
            }

            return item;
        }

        /// <summary>
        /// Create an item with one recipient.
        /// </summary>
        /// <returns>The item object.</returns>
        protected MessageType CreateItemWithOneRecipient()
        {
            EmailAddressType address;

            address = new EmailAddressType()
            {
                EmailAddress = Common.GetConfigurationPropertyValue("User2Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site),
            };

            MessageType message = new MessageType()
            {
                Subject = Common.GenerateResourceName(
                    this.Site,
                    TestSuiteHelper.SubjectForCreateItem),

                ToRecipients = new EmailAddressType[]
                {
                    address
                }
            };

            return message;
        }

        /// <summary>
        /// Create an item by using the type parameter.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be created.</param>
        /// <returns>The ItemId of the created item.</returns>
        protected ItemIdType[] CreateItemForSpecificItemType<T>(T item)
            where T : ItemType, new()
        {
            CreateItemType createItemRequest = new CreateItemType();

            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = new ItemType[] { item };

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e" and set PropertyId to "123" with Int32 type.
            createItemRequest.Items.Items[0] = this.SetPathToExtendedFieldTypeProperties<T>(item, false, false, true, true, false, false);

            if (createItemRequest.Items.Items[0] is MessageType)
            {
                DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
                folderIdForCreateItems.Id = DistinguishedFolderIdNameType.drafts;
                createItemRequest.SavedItemFolderId = new TargetFolderIdType();
                createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
                createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
                createItemRequest.MessageDispositionSpecified = true;
            }

            if (createItemRequest.Items.Items[0] is CalendarItemType)
            {
                DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
                folderIdForCreateItems.Id = DistinguishedFolderIdNameType.calendar;
                createItemRequest.SavedItemFolderId = new TargetFolderIdType();
                createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
                createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
                createItemRequest.SendMeetingInvitationsSpecified = true;
            }

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created contact item should be returned.
            Site.Assert.AreEqual<int>(
                1,
                 createdItemIds.GetLength(0),
                 "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdItemIds.GetLength(0));

            return createdItemIds;
        }

        /// <summary>
        /// Create an item for SendItem operation.
        /// </summary>
        /// <param name="itemSubject">Subject of the created item.</param>
        /// <returns>The ItemId of the created item.</returns>
        protected ItemIdType[] CreateItemWithRecipient(string itemSubject)
        {
            #region Config the item
            MessageType[] items = new MessageType[] { this.CreateItemWithOneRecipient() };
            items[0].Subject = itemSubject;
            #endregion

            CreateItemResponseType createItemResponse = this.CallCreateItemOperation(DistinguishedFolderIdNameType.drafts, items);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 1, this.Site);

            ItemIdType[] createdItems = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // One created item should be returned.
            Site.Assert.AreEqual<int>(
                 1,
                 createdItems.GetLength(0),
                 "One created item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 createdItems.GetLength(0));

            return createdItems;
        }

        /// <summary>
        /// Create item with minimum elements which are needed.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be created.</param>
        /// <returns>The ItemId of the created item.</returns>
        protected ItemIdType[] CreateItemWithMinimumElements<T>(T item)
            where T : ItemType, new()
        {
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the item
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();

            createItemRequest.Items.Items = new ItemType[] { item };
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem);

            if (createItemRequest.Items.Items[0] is MessageType)
            {
                DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
                folderIdForCreateItems.Id = DistinguishedFolderIdNameType.drafts;
                createItemRequest.SavedItemFolderId = new TargetFolderIdType();
                createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
                createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
                createItemRequest.MessageDispositionSpecified = true;
            }

            if (createItemRequest.Items.Items[0] is CalendarItemType)
            {
                DistinguishedFolderIdType folderIdForCreateItems = new DistinguishedFolderIdType();
                folderIdForCreateItems.Id = DistinguishedFolderIdNameType.calendar;
                createItemRequest.SavedItemFolderId = new TargetFolderIdType();
                createItemRequest.SavedItemFolderId.Item = folderIdForCreateItems;
                createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
                createItemRequest.SendMeetingInvitationsSpecified = true;
            }
            #endregion

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

            return createdItemIds;
        }

        /// <summary>
        /// Create items for all item types.
        /// </summary>
        /// <returns>The ItemIds of the created items.</returns>
        protected ItemIdType[] CreateAllTypesItems()
        {
            // Initialize items data.
            object obj;
            List<ItemType> items = new List<ItemType>();
            ItemIdType[] createdItemIds;

            // Get the ItemType and six extend types which base on ItemType.
            Assembly assembly = Assembly.GetAssembly(typeof(ItemType));
            Type[] types = assembly.GetTypes();

            // Initialize the public properties (Subject and Body) which the seven kinds of operation both have.
            PropertyInfo subjectField;
            PropertyInfo bodyField;
            string subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem);
            BodyType body = new BodyType()
            {
                Value = "Body For Base Item",
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

                    // RoleMemberItemType and NetworkItemType are for internal use only.
                    // AbchPersonItemType is covered in MS-OXWSCONT
                    if (type != typeof(RoleMemberItemType)
                        && type != typeof(NetworkItemType)
                        && type != typeof(AbchPersonItemType))
                    {
                        items.Add((ItemType)obj);
                    }
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

            // Check whether the CreateItem operation is executed successfully.
            foreach (ResponseMessageType responseMessage in createResponse.ResponseMessages.Items)
            {
                Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "Create each types of items should succeed! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }

            return createdItemIds;
        }

        /// <summary>
        /// The operation searches the parent folder and returns items that meet a specified search restriction.
        /// </summary>
        /// <param name="parentFolder">A enumeration value that specifies the distinguished folder to search.</param>
        /// <param name="searchRestriction">A string that specifies the value for a search restriction.</param>
        /// <param name="role">A string that specifies the identity with which to search mailbox.</param>
        /// <returns>If the operation succeeds, return the IDs of the found items.</returns>
        protected ItemIdType[] FindItemsInFolder(DistinguishedFolderIdNameType parentFolder, string searchRestriction, string role)
        {
            #region Switch role
            this.SwitchUser(role);
            #endregion

            #region Loop to find the items.
            int counter = 0;
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            while (counter < upperBound)
            {
                // Wait for the item received.
                System.Threading.Thread.Sleep(waitTime);

                // Find the item.
                ItemType[] foundItems = this.FindItemWithRestriction(parentFolder, searchRestriction);

                if (foundItems != null)
                {
                    List<ItemIdType> foundItemIds = new List<ItemIdType>();
                    foreach (ItemType foundItem in foundItems)
                    {
                        if (searchRestriction == null || foundItem.Subject.Contains(searchRestriction))
                        {
                            foundItemIds.Add(foundItem.ItemId);
                        }
                    }

                    if (foundItemIds.Count != 0)
                    {
                        // Log the retry count when the item has been found.
                        Site.Log.Add(LogEntryKind.Debug, string.Format("The retry count of FindItem operation is {0}.", counter));

                        return foundItemIds.ToArray();
                    }
                }

                counter++;
            }

            // Log the retry count when the item has not been found.
            Site.Log.Add(LogEntryKind.Debug, string.Format("The retry count of FindItem operation is {0}.", counter));

            #endregion

            return null;
        }

        /// <summary>
        /// The operation searches the mailbox and returns items that meet a specified search restriction.
        /// </summary>
        /// <param name="folder">A enumeration value that specifies the distinguished folder to search.</param>
        /// <param name="searchRestriction">A string that specifies the value for a search restriction.</param>
        /// <returns>If the operation succeeds, return an array of all matched items; otherwise, return null.</returns>
        protected ItemType[] FindItemWithRestriction(DistinguishedFolderIdNameType folder, string searchRestriction)
        {
            #region Construct FindItem request
            FindItemType findRequest = new FindItemType();

            #region Specify all properties to return in FindItem response.
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            #endregion

            #region Specify a folder that is the root of the search.
            // Specifies searching items in the distinguished folder represents by input parameter folder
            DistinguishedFolderIdType disFolderId = new DistinguishedFolderIdType();
            disFolderId.Id = folder;
            findRequest.ParentFolderIds = new BaseFolderIdType[1];
            findRequest.ParentFolderIds[0] = disFolderId;
            #endregion

            #region Specifies a search restriction or query
            PathToUnindexedFieldType itemSubject = new PathToUnindexedFieldType();
            itemSubject.FieldURI = UnindexedFieldURIType.itemSubject;
            ContainsExpressionType expressionType = new ContainsExpressionType();
            expressionType.Item = itemSubject;

            // Specifies that the comparison is between the substring of the property value and the constant
            expressionType.ContainmentMode = ContainmentModeType.Substring;

            // Indicates the ContainmentMode property is serialized in the SOAP message.
            expressionType.ContainmentModeSpecified = true;

            // Specifies that the comparison ignores casing and non-spacing characters
            expressionType.ContainmentComparison = ContainmentComparisonType.IgnoreCaseAndNonSpacingCharacters;

            // Indicates the ContainmentComparison property is serialized in the SOAP message.
            expressionType.ContainmentComparisonSpecified = true;
            expressionType.Constant = new ConstantValueType();
            expressionType.Constant.Value = searchRestriction;

            RestrictionType restriction = new RestrictionType();
            restriction.Item = expressionType;
            if (!string.IsNullOrEmpty(searchRestriction))
            {
                findRequest.Restriction = restriction;
            }
            #endregion
            #endregion

            #region Call FindItem operation
            FindItemResponseType findResponse = this.SRCHAdapter.FindItem(findRequest);

            if (findResponse != null
                && findResponse.ResponseMessages != null
                && findResponse.ResponseMessages.Items != null
                && findResponse.ResponseMessages.Items.Length > 0)
            {
                // The count of items in ResponseMessages should be one.
                Site.Assert.AreEqual<int>(
                1,
                findResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count in ResponseMessages: {0}, Actual: {1}",
                1,
                findResponse.ResponseMessages.Items.GetLength(0));

                ArrayOfRealItemsType items = ((FindItemResponseMessageType)findResponse.ResponseMessages.Items[0]).RootFolder.Item as ArrayOfRealItemsType;

                if (items != null && items.Items != null && items.Items.Length > 0)
                {
                    List<ItemType> foundItems = new List<ItemType>();
                    foreach (ItemType item in items.Items)
                    {
                        if (item.ItemId != null && !string.IsNullOrEmpty(item.ItemId.Id))
                        {
                            foundItems.Add(item);
                        }
                    }

                    if (foundItems.Count != 0)
                    {
                        return foundItems.ToArray();
                    }
                }
            }

            return null;
            #endregion
        }

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        /// <returns>A dictionary of configured SOAP headers.</returns>
        protected Dictionary<string, object> ConfigureSOAPHeader()
        {
            // Initialize a dictionary to store common SOAP headers.
            Dictionary<string, object> soapHeaders = new Dictionary<string, object>();

            // Configure the Impersonation SOAP header and add it to the soapHeaders list.
            ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
            impersonation.ConnectingSID = new ConnectingSIDType();
            PrimarySmtpAddressType smtpAddress = new PrimarySmtpAddressType();
            smtpAddress.Value = Common.GetConfigurationPropertyValue("User1Name", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);
            impersonation.ConnectingSID.Item = smtpAddress;
            soapHeaders.Add("ExchangeImpersonation", impersonation);

            // Configure the MailboxCulture SOAP header and add it to the soapHeaders list.
            MailboxCultureType mailboxCulture = new MailboxCultureType();
            mailboxCulture.Value = TestSuiteHelper.Culture;
            soapHeaders.Add("MailboxCulture", mailboxCulture);

            // Return the list.
            return soapHeaders;
        }

        /// <summary>
        /// Clear the soap headers.
        /// </summary>
        protected void ClearSoapHeaders()
        {
            // Initialize a dictionary to store common SOAP headers.
            Dictionary<string, object> soapHeaders = new Dictionary<string, object>();

            // Set the value of the soap headers to null.
            soapHeaders.Add("ExchangeImpersonation", null);
            soapHeaders.Add("MailboxCulture", null);
            soapHeaders.Add("TimeZoneContext", null);
            soapHeaders.Add("DateTimePrecision", null);

            // Configure the SOAP header.
            this.COREAdapter.ConfigureSOAPHeader(soapHeaders);
        }

        /// <summary>
        /// Clean all items which have been sent out to User2.
        /// </summary>
        /// <param name="itemSubjects">The subjects of items which have been sent out to User2.</param>
        /// <param name="isCalendarChecked">Indicates whether need to check the Calendar folder.</param>
        protected void CleanItemsSentOut(string[] itemSubjects, bool isCalendarChecked = false)
        {
            List<ItemIdType> receivedItemIds = new List<ItemIdType>();

            for (int i = 0; i < itemSubjects.Length; i++)
            {
                // Find items in the Inbox folder of User2 which were sent out by User1 with subject.
                ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.inbox, itemSubjects[i], "User2");

                Site.Assert.IsNotNull(findItemIds, "Mail item should be available in User2 mailbox.");
                Site.Assert.AreEqual<int>(1, findItemIds.Length, "There should be only one item with subject {0}.", itemSubjects[i]);

                receivedItemIds.Add(findItemIds[0]);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R424");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R424
                // The sent item is found in receiver's mailbox, this requirement can be captured.
                this.Site.CaptureRequirement(
                    424,
                    @"[In SendItem Operation] The SendItem operation sends message items on the server.");
            }

            if (isCalendarChecked)
            {
                for (int i = 0; i < itemSubjects.Length; i++)
                {
                    // Find items in the Calendar folder of User2 which were sent out by User1 with subject.
                    ItemIdType[] findItemIds = this.FindItemsInFolder(DistinguishedFolderIdNameType.calendar, itemSubjects[i], "User2");

                    if (findItemIds != null)
                    {
                        receivedItemIds.Add(findItemIds[0]);
                    }
                }
            }

            // Delete the found items.
            DeleteItemType deleteItemRequest = new DeleteItemType();
            deleteItemRequest.ItemIds = receivedItemIds.ToArray();
            deleteItemRequest.DeleteType = DisposalType.HardDelete;

            // AffectedTaskOccurrences indicates whether a task instance or a task master is to be deleted.
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;

            if (Common.IsRequirementEnabled(2311, this.Site))
            {
                deleteItemRequest.SuppressReadReceipts = true;
                deleteItemRequest.SuppressReadReceiptsSpecified = true;
            }

            // SendMeetingCancellations describes how cancellations are to be handled for deleted meetings.
            deleteItemRequest.SendMeetingCancellationsSpecified = true;
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);

            this.Site.Assert.AreEqual<int>(
                receivedItemIds.Count,
                deleteItemResponse.ResponseMessages.Items.GetLength(0),
                "Expected Item Count: {0}, Actual Item Count: {1}",
                receivedItemIds.Count,
                deleteItemResponse.ResponseMessages.Items.GetLength(0));

            foreach (ResponseMessageType responseMessage in deleteItemResponse.ResponseMessages.Items)
            {
                this.Site.Assert.AreEqual<ResponseClassType>(
                        ResponseClassType.Success,
                        responseMessage.ResponseClass,
                        string.Format(
                            "The operation should be successful! Expected response code: {0}, actual response code: {1}",
                            ResponseCodeType.NoError,
                            responseMessage.ResponseCode));
            }

            // Switch the credentials to User1.
            this.SwitchUser("User1");
        }

        /// <summary>
        /// Switch the credentials for different users.
        /// </summary>
        /// <param name="user">The user which will be used.</param>
        protected void SwitchUser(string user)
        {
            if (user == "User1" || user == "User2")
            {
                // Initialize user name and password.
                string userName = null;
                string password = null;

                // Get user name and password from ptfconfig.
                if (user == "User1")
                {
                    userName = Common.GetConfigurationPropertyValue("User1Name", this.Site);
                    password = Common.GetConfigurationPropertyValue("User1Password", this.Site);
                }
                else
                {
                    userName = Common.GetConfigurationPropertyValue("User2Name", this.Site);
                    password = Common.GetConfigurationPropertyValue("User2Password", this.Site);
                }

                string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
                this.COREAdapter.SwitchUser(userName, password, domain);
                this.SRCHAdapter.SwitchUser(userName, password, domain);
            }
            else
            {
                Site.Assert.Fail("The value of the argument user of SwitchUser is invalid. It should be one of the possible values: 'User1' and 'User2'");
            }
        }

        /// <summary>
        /// Find new items in a specific folder.
        /// </summary>
        /// <param name="folder">The folder in which items will be found.</param>
        protected void FindNewItemsInFolder(DistinguishedFolderIdNameType folder)
        {
            // Find items in the mailbox of User1 without restriction.
            ItemIdType[] findItemId = this.FindItemsInFolder(folder, null, "User1");

            Site.Assert.IsNotNull(findItemId, "There should be at least one item in the mailbox of User1.");

            // Add the found items into the ExistItemIds collection.
            foreach (ItemIdType itemId in findItemId)
            {
                if (!this.IsIdExisted(itemId))
                {
                    this.ExistItemIds.Add(itemId);
                }
            }
        }

        /// <summary>
        /// Verify whether the precision of DateTime elements in response are expected precision.
        /// </summary>
        /// <param name="responseRawXml">The raw XML response received from protocol SUT.</param>
        /// <param name="expectedPrecision">The expected DateTime precision.</param>
        /// <returns>If the precision of DateTime elements in response are expected precision, return true. Otherwise, return false.</returns>
        protected bool IsExpectedDateTimePrecision(XmlElement responseRawXml, string expectedPrecision)
        {
            XmlNodeList nodes = responseRawXml.GetElementsByTagName("t:Message");
            foreach (XmlNode node in nodes)
            {
                if (node.HasChildNodes && node.ChildNodes != null && node.ChildNodes.Count > 0)
                {
                    foreach (XmlNode child in node.ChildNodes)
                    {
                        DateTime result;
                        if (DateTime.TryParse(child.InnerText, out result))
                        {
                            // When the expected precision is Seconds, "." should not exist in the value of DateTime element.
                            if (expectedPrecision == "Seconds" && child.InnerText.Contains("."))
                            {
                                this.Site.Log.Add(LogEntryKind.Debug, "The value of DateTime element {0} is {1}.", child.Name, child.InnerText);
                                return false;
                            }

                            // When the expected precision is Milliseconds, "." should exist in the value of DateTime element.
                            if (expectedPrecision == "Milliseconds" && !child.InnerText.Contains("."))
                            {
                                this.Site.Log.Add(LogEntryKind.Debug, "The value of DateTime element {0} is {1}.", child.Name, child.InnerText);
                                return false;
                            }
                        }
                    }
                }
            }

            // If all the DateTime elements are expected precision, return true.
            return true;
        }

        #region The method of setting the PathToExtendedFieldType property
        /// <summary>
        /// Set the properties of the type "PathToExtendedFieldType".
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be updated.</param>
        /// <param name="distinguishedPropertySetIdSpec">A boolean value which indicates whether enable the DistinguishedPropertySetId property.</param>
        /// <param name="propertyTagSpec">A boolean value which indicates whether enable the PropertyTag property.</param>
        /// <param name="propertySetIdSpec">A boolean value which indicates whether enable the PropertySetId property.</param>
        /// <param name="propertyIdSpec">A boolean value which indicates whether enable the PropertyId property.</param>
        /// <param name="propertyNameSpec">A boolean value which indicates whether enable the PropertyName property.</param>
        /// <param name="isArrayString">A boolean value which indicates whether the property type value is string or array of string.</param>
        /// <returns>Return an ItemType or its child class object.</returns>
        protected T SetPathToExtendedFieldTypeProperties<T>(
                             T item,
                             bool distinguishedPropertySetIdSpec,
                             bool propertyTagSpec,
                             bool propertySetIdSpec,
                             bool propertyIdSpec,
                             bool propertyNameSpec,
                             bool isArrayString)
        where T : ItemType, new()
        {
            item.Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForCreateItem);
            item.Body = new BodyType();
            item.Body.BodyType1 = BodyTypeType.HTML;
            item.Body.Value = @"<html><head><title>Test</title></head><body><script type=""text/javascript"">It's in script block.</script>It's out of script block.<a href = "" http://www.microsoft.com"" target=""_top""><img src=""http://www.microsoft.com/global/ImageStore/PublishingImages/logos/hp/logo-lg-1x.png"" /></a></body></html>";
            ExtendedPropertyType extendedPropertyType = new ExtendedPropertyType();
            extendedPropertyType.ExtendedFieldURI = new PathToExtendedFieldType();

            // Set the extend properties for the element.
            DistinguishedPropertySetType distinguishedPropertySetIdValue = DistinguishedPropertySetType.Common;

            // The property tag. The PropertyTag attribute can be represented as either a hexadecimal value or a short integer.
            // The hexadecimal value range: 0x8000< hexadecimal value <0xFFFE, it represents the custom range of properties.
            string propertyTagValue = "0x3a45";

            // The GUID value.
            string propertySetIdValue = "c11ff724-aa03-4555-9952-8fa248a11c3e";

            int propertyIdValue = Convert.ToInt32(TestSuiteHelper.PropertyId);
            string propertyNameValue = TestSuiteHelper.PropertyName;
            MapiPropertyTypeType propertyTypeValue = new MapiPropertyTypeType();
            object propertyValue = null;
            if (isArrayString)
            {
                propertyTypeValue = MapiPropertyTypeType.StringArray;
                NonEmptyArrayOfPropertyValuesType propertyValues = new NonEmptyArrayOfPropertyValuesType();
                propertyValues.Items = new string[] { TestSuiteHelper.ElementValue };
                propertyValue = propertyValues;
            }
            else
            {
                propertyTypeValue = MapiPropertyTypeType.String;
                propertyValue = TestSuiteHelper.ElementValue;
            }

            // Set the properties for the element "ExtendedFieldURI" with the type "PathToExtendedFieldType".
            if (distinguishedPropertySetIdSpec)
            {
                extendedPropertyType.ExtendedFieldURI.DistinguishedPropertySetId = distinguishedPropertySetIdValue;
                extendedPropertyType.ExtendedFieldURI.DistinguishedPropertySetIdSpecified = true;
            }

            if (propertyTagSpec)
            {
                extendedPropertyType.ExtendedFieldURI.PropertyTag = propertyTagValue;
            }

            if (propertySetIdSpec)
            {
                extendedPropertyType.ExtendedFieldURI.PropertySetId = propertySetIdValue;
            }

            if (propertyIdSpec)
            {
                extendedPropertyType.ExtendedFieldURI.PropertyId = propertyIdValue;
                extendedPropertyType.ExtendedFieldURI.PropertyIdSpecified = true;
            }

            if (propertyNameSpec)
            {
                extendedPropertyType.ExtendedFieldURI.PropertyName = propertyNameValue;
            }

            extendedPropertyType.ExtendedFieldURI.PropertyType = propertyTypeValue;
            extendedPropertyType.Item = propertyValue;
            item.ExtendedProperty = new ExtendedPropertyType[]
            {
                extendedPropertyType
            };
            return item;
        }
        #endregion

        #region MS-OXWSCORE adapter method.
        /// <summary>
        /// Copy items and puts the items in a different folder.
        /// </summary>
        /// <param name="folderId">The destination folder.</param>
        /// <param name="itemIds">Contain the unique identities of items to be copy.</param>
        /// <returns>A response to this operation request.</returns>
        protected CopyItemResponseType CallCopyItemOperation(DistinguishedFolderIdNameType folderId, BaseItemIdType[] itemIds)
        {
            CopyItemType requestItem = new CopyItemType();
            requestItem.ToFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = folderId;
            requestItem.ToFolderId.Item = distinguishedFolderId;
            requestItem.ItemIds = itemIds;

            CopyItemResponseType response = this.COREAdapter.CopyItem(requestItem);
            return response;
        }

        /// <summary>
        /// Create items in the Exchange store.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="folderId">The folder in which the items are created.</param>
        /// <param name="items">Items to be created.</param>
        /// <returns>A response to this operation request.</returns>
        protected CreateItemResponseType CallCreateItemOperation<T>(
            DistinguishedFolderIdNameType folderId,
            T[] items)
        where T : ItemType
        {
            CreateItemType requestItem = new CreateItemType();

            // If items are message items, below properties need to be specified
            if (items is MessageType[])
            {
                requestItem.MessageDispositionSpecified = true;
                requestItem.MessageDisposition = MessageDispositionType.SaveOnly;
            }

            // If items are calendar items, below properties need to be specified
            if (items is CalendarItemType[])
            {
                requestItem.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
                requestItem.SendMeetingInvitationsSpecified = true;
            }

            requestItem.SavedItemFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType distinguishedFolderId = new DistinguishedFolderIdType();
            distinguishedFolderId.Id = folderId;
            requestItem.SavedItemFolderId.Item = distinguishedFolderId;
            requestItem.Items = new NonEmptyArrayOfAllItemsType();
            requestItem.Items.Items = items;

            CreateItemResponseType response = this.COREAdapter.CreateItem(requestItem);

            return response;
        }

        /// <summary>
        /// Get items on the server.
        /// </summary>
        /// <param name="itemIds">The items id to be gotten.</param>
        /// <returns>A response to this operation request.</returns>
        protected GetItemResponseType CallGetItemOperation(BaseItemIdType[] itemIds)
        {
            GetItemType requestItem = new GetItemType();
            requestItem.ItemShape = new ItemResponseShapeType()
            {
                BaseShape = DefaultShapeNamesType.AllProperties
            };

            requestItem.ItemIds = itemIds;

            GetItemResponseType response = this.COREAdapter.GetItem(requestItem);

            #region Steps to obtain ConversationIdMailboxGuidBased storage type
            foreach (ResponseMessageType responseItem in response.ResponseMessages.Items)
            {
                ItemInfoResponseMessageType item = responseItem as ItemInfoResponseMessageType;
                if (item.Items != null && item.Items.Items != null)
                {
                    if (item.Items.Items[0] as MessageType != null)
                    {
                        MessageType messageItem = item.Items.Items[0] as MessageType;
                        if (messageItem.ConversationId != null)
                        {
                            ItemIdId conversationIdId = this.ITEMIDAdapter.ParseItemId(messageItem.ConversationId);

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R64");

                            // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R64
                            Site.CaptureRequirementIfAreEqual<IdStorageType>(
                                IdStorageType.ConversationIdMailboxGuidBased,
                                conversationIdId.StorageType,
                                "MS-OXWSITEMID",
                                64,
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

                                        /// <summary>
                                        /// The Id represents an item or folder in a mailbox and contains a mailbox GUID.
                                        /// </summary>
                                        MailboxItemMailboxGuidBased = 3,]

                                        /// <summary>
                                        /// The Id represents a conversation in a mailbox and contains a mailbox GUID.
                                        /// </summary>
                                        ConversationIdMailboxGuidBased = 4,
                                [
                                        /// <summary>
                                        /// The Id represents (by objectGuid) an object in the Active Directory.
                                        /// </summary>
                                        ActiveDirectoryObject = 5,]
                                }");
                        }
                    }
                }
            }
            #endregion Steps to obtain ConversationIdMailboxGuidBased storage type

            return response;
        }

        /// <summary>
        /// Get the extended property of the item return by GetItem operation.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        /// <param name="pathToExtendedFieldType">The extended property field to be returned.</param>
        /// <returns>The extended property of the item in the GetItem response.</returns>
        protected ExtendedPropertyType CallGetItemOperationWithAdditionalProperties<T>(T item, PathToExtendedFieldType pathToExtendedFieldType)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();
            getItem.ItemIds = new BaseItemIdType[] { item.ItemId };
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItem.ItemShape.AdditionalProperties = new BasePathToElementType[] { pathToExtendedFieldType };

            // Call GetItem to get the created item by using getItem.
            GetItemResponseType getItemResponse = new GetItemResponseType();
            getItemResponse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            ItemInfoResponseMessageType itemInfoGetResponseMessage = new ItemInfoResponseMessageType();
            itemInfoGetResponseMessage = getItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            item = itemInfoGetResponseMessage.Items.Items[0] as T;
            ExtendedPropertyType extendedProperty = TestSuiteHelper.Copy<ExtendedPropertyType>(item.ExtendedProperty[0]);

            // Check whether the extended property is not null
            Site.Assert.IsNotNull(
                extendedProperty,
                "The extended property should not be null.");

            return extendedProperty;
        }

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <returns>A response to this operation request.</returns>
        protected DeleteItemResponseType CallDeleteItemOperation()
        {
            return this.CallDeleteItemOperation(DisposalType.HardDelete);
        }

        /// <summary>
        /// Delete items on the server.
        /// </summary>
        /// <param name="disposalType">The disposal type of item to be delete.</param>
        /// <returns>A response to this operation request.</returns>
        protected DeleteItemResponseType CallDeleteItemOperation(DisposalType disposalType)
        {
            // Get ItemIds.
            DeleteItemType deleteItemRequest = new DeleteItemType();
            ItemIdType[] itemArray = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(itemArray, 0);
            deleteItemRequest.ItemIds = itemArray;

            // Enumeration value to describe how an item is to be deleted.
            deleteItemRequest.DeleteType = disposalType;

            // AffectedTaskOccurrences indicates whether a task instance or a task master is to be deleted.
            deleteItemRequest.AffectedTaskOccurrencesSpecified = true;
            deleteItemRequest.AffectedTaskOccurrences = AffectedTaskOccurrencesType.AllOccurrences;

            // SendMeetingCancellations describes how cancellations are to be handled for deleted meetings.
            deleteItemRequest.SendMeetingCancellationsSpecified = true;
            deleteItemRequest.SendMeetingCancellations = CalendarItemCreateOrDeleteOperationType.SendToNone;
            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);

            return deleteItemResponse;
        }

        /// <summary>
        /// Move items on the server.
        /// </summary>
        /// <param name="folderId">The folder the items to be moved to.</param>
        /// <param name="itemIds">Items' id to be moved.</param>
        /// <returns>A response to this operation request.</returns>
        protected MoveItemResponseType CallMoveItemOperation(
            DistinguishedFolderIdNameType folderId,
            BaseItemIdType[] itemIds)
        {
            MoveItemType requestItem = new MoveItemType();

            requestItem.ToFolderId = new TargetFolderIdType();
            DistinguishedFolderIdType distinguishedFolerId = new DistinguishedFolderIdType();
            distinguishedFolerId.Id = folderId;
            requestItem.ToFolderId.Item = distinguishedFolerId;
            requestItem.ItemIds = itemIds;

            MoveItemResponseType response = this.COREAdapter.MoveItem(requestItem);
            return response;
        }

        /// <summary>
        /// Send messages and post items on the server.
        /// </summary>
        /// <param name="itemIds">Items' id to be moved.</param>
        /// <param name="folderId">Specify the identity of the folder that contains a saved version of a sent item.</param>
        /// <param name="saveItemToFolder">Specify a boolean value that indicates whether a copy of a sent item is saved.</param>
        /// <returns>A response to this operation request.</returns>
        protected SendItemResponseType CallSendItemOperation(
            BaseItemIdType[] itemIds,
            DistinguishedFolderIdNameType folderId,
            bool saveItemToFolder)
        {
            SendItemType requestItem = new SendItemType();
            requestItem.SaveItemToFolder = saveItemToFolder;
            if (saveItemToFolder)
            {
                requestItem.SavedItemFolderId = new TargetFolderIdType();
                DistinguishedFolderIdType distinguishedFol = new DistinguishedFolderIdType();
                distinguishedFol.Id = folderId;
                requestItem.SavedItemFolderId.Item = distinguishedFol;
            }

            requestItem.ItemIds = itemIds;

            SendItemResponseType response = this.COREAdapter.SendItem(requestItem);
            return response;
        }

        /// <summary>
        /// Mark all items as read or unread on the server.
        /// </summary>
        /// <param name="readFlag">Specify a Boolean value that indicates whether to mark the items as read or not.</param>
        /// <param name="suppressReadReceipts">Specify a boolean value that indicates the read receipts are suppressed or not.</param>
        /// <param name="folderIds">Specify the folders' id.</param>
        /// <returns>A response to this operation request.</returns>
        protected MarkAllItemsAsReadResponseType CallMarkAllItemsAsReadOperation(
            bool readFlag,
            bool suppressReadReceipts,
            BaseFolderIdType[] folderIds)
        {
            MarkAllItemsAsReadType markAllItemsAsReadRequest = new MarkAllItemsAsReadType();

            // The request properties.
            markAllItemsAsReadRequest.ReadFlag = readFlag;
            markAllItemsAsReadRequest.SuppressReadReceipts = suppressReadReceipts;
            markAllItemsAsReadRequest.FolderIds = folderIds;

            MarkAllItemsAsReadResponseType markAllItemsAsReadResponse = this.COREAdapter.MarkAllItemsAsRead(markAllItemsAsReadRequest);

            return markAllItemsAsReadResponse;
        }

        /// <summary>
        /// Update items on the server.
        /// </summary>
        /// <param name="folderId">Specify the target folder identifier for saved items.</param>
        /// <param name="saveItemToFolder">Specify a boolean value that indicates whether a folder specified for saved items.</param>
        /// <param name="itemChanges">Specify an array of item changes.</param>
        /// <returns>A response to this operation request.</returns>
        protected UpdateItemResponseType CallUpdateItemOperation(
            DistinguishedFolderIdNameType folderId,
            bool saveItemToFolder,
            ItemChangeType[] itemChanges)
        {
            UpdateItemType requestItem = new UpdateItemType();
            if (saveItemToFolder)
            {
                requestItem.SavedItemFolderId = new TargetFolderIdType();
                DistinguishedFolderIdType distinguishedFol = new DistinguishedFolderIdType();
                distinguishedFol.Id = folderId;
                requestItem.SavedItemFolderId.Item = distinguishedFol;
            }

            requestItem.ItemChanges = itemChanges;
            requestItem.MessageDisposition = MessageDispositionType.SaveOnly;
            requestItem.MessageDispositionSpecified = true;
            requestItem.SendMeetingInvitationsOrCancellations = CalendarItemUpdateOperationType.SendToNone;
            requestItem.SendMeetingInvitationsOrCancellationsSpecified = true;
            requestItem.ConflictResolution = ConflictResolutionType.AutoResolve;
            UpdateItemResponseType response = this.COREAdapter.UpdateItem(requestItem);
            return response;
        }
        #endregion

        /// <summary>
        /// Handle the server response.
        /// </summary>
        /// <param name="request">The request messages.</param>
        /// <param name="response">The response messages.</param>
        /// <param name="isSchemaValidated">Verify the schema.</param>
        protected void ExchangeServiceBinding_ResponseEvent(
            BaseRequestType request,
            BaseResponseMessageType response,
            bool isSchemaValidated)
        {
            this.IsSchemaValidated = isSchemaValidated;
            this.LastResponse = response;

            bool hasItemInfo = false;
            foreach (ResponseMessageType responseMessage in response.ResponseMessages.Items)
            {
                if (responseMessage is ItemInfoResponseMessageType)
                {
                    hasItemInfo = true;
                    break;
                }
            }

            BaseItemIdType[] itemIds;
            if (hasItemInfo)
            {
                itemIds = Common.GetItemIdsFromInfoResponse(response);
            }
            else
            {
                itemIds = new BaseItemIdType[0];
            }

            foreach (ItemIdType itemId in itemIds)
            {
                bool notExist = true;
                foreach (ItemIdType exist in this.ExistItemIds)
                {
                    if (exist.Id == itemId.Id)
                    {
                        notExist = false;
                        break;
                    }
                }

                if (notExist)
                {
                    this.ExistItemIds.Add(itemId);
                    this.VerifyRLECompress(itemId);
                    this.VerifyRLEDecompress(itemId);
                }
            }
        }

        #region common test steps
        /// <summary>
        /// Verify the response returned by GetItem operation with the ItemShape element.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the created item.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.IdOnly;
            getItem.ItemShape.IncludeMimeContent = false;
            getItem.ItemShape.IncludeMimeContentSpecified = true;
            getItem.ItemShape.BodyType = BodyTypeResponseType.Best;
            getItem.ItemShape.BodyTypeSpecified = true;
            getItem.ItemShape.AdditionalProperties = new PathToUnindexedFieldType[] { new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.itemBody } };

            if (Common.IsRequirementEnabled(21498, this.Site))
            {
                getItem.ItemShape.ConvertHtmlCodePageToUTF8 = false;
                getItem.ItemShape.ConvertHtmlCodePageToUTF8Specified = true;
            }

            if (Common.IsRequirementEnabled(2149904, this.Site))
            {
                getItem.ItemShape.InlineImageUrlTemplate = "InlineImageUrlTemplate";
            }

            if (Common.IsRequirementEnabled(2149905, this.Site))
            {
                getItem.ItemShape.BlockExternalImages = false;
                getItem.ItemShape.BlockExternalImagesSpecified = true;
            }

            if (Common.IsRequirementEnabled(2149908, this.Site))
            {
                getItem.ItemShape.AddBlankTargetToLinks = false;
                getItem.ItemShape.AddBlankTargetToLinksSpecified = true;
            }

            // Get the created item.
            GetItemResponseType getItemResponse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            item = Common.GetItemsFromInfoResponse<T>(getItemResponse)[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1197");

            // The request set the itemBody in AdditionalProperties element of request
            // If the Body element of the item in response is not null, which represents requested additional property is returned in a response,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.Body,
                "MS-OXWSCDATA",
                1197,
                @"[In t:ItemResponseShapeType Complex Type] The element ""AdditionalProperties"" with type ""t:NonEmptyArrayOfPathsToElementType"" Specifies a set of requested additional properties to return in a response.");

            if (Common.IsRequirementEnabled(2149904, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2149904");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R2149904
                // The server response with a successful response when including InlineImageUrlTemplate element in request
                // this requirement can be verified.
                Site.CaptureRequirement(
                    "MS-OXWSCDATA",
                    2149904,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""InlineImageUrlTemplate"" with type ""xs:string ([XMLSCHEMA2])"" which specifies the name of the template for the inline image URL. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with different DefaultShapeNamesType enumeration values in ItemShape element.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_DefaultShapeNamesTypeEnum<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the created item with BaseShape set to AllProperties.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            GetItemResponseType getItemResponse_AllProperties = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_AllProperties, 1, this.Site);

            T[] item_AllProperties = Common.GetItemsFromInfoResponse<T>(getItemResponse_AllProperties);

            Site.Assert.AreEqual<int>(
                 1,
                 item_AllProperties.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_AllProperties.GetLength(0));

            Site.Assert.IsNotNull(
                item_AllProperties[0].Subject,
                "The subject element in returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R59");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R59
            // The request have get item all properties,
            // and the responses are successfully,
            // this requirement can be verified.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                59,
                @"[In t:DefaultShapeNamesType Simple Type] The value ""AllProperties"" specifies all the properties that are defined for the AllProperties shape.");
            
            #endregion

            #region Step 3: Get the created item with BaseShape set to IdOnly.
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.IdOnly;

            GetItemResponseType getItemResponse_IdOnly = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_IdOnly, 1, this.Site);

            T[] item_IdOnly = Common.GetItemsFromInfoResponse<T>(getItemResponse_IdOnly);

            Site.Assert.AreEqual<int>(
                 1,
                 item_IdOnly.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_IdOnly.GetLength(0));

            // Get the tag name according to the type of item.
            string tagName = null;

            if (item is ContactItemType)
            {
                tagName = "t:Contact";
            }
            else if (item is CalendarItemType)
            {
                tagName = "t:CalendarItem";
            }
            else if (item is TaskType)
            {
                tagName = "t:Task";
            }
            else if (item is PostItemType)
            {
                tagName = "t:PostItem";
            }
            else if (item is MessageType)
            {
                tagName = "t:Message";
            }
            else if (item is DistributionListType)
            {
                tagName = "t:DistributionList";
            }
            else
            {
                tagName = "t:Message";
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R62");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R62
            // The request set BaseShape element to IdOnly,
            // If only ItemId element presents, and all other child elements of the item is null or not specified,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                Common.IsIdOnly((XmlElement)this.COREAdapter.LastRawResponseXml, tagName, "t:ItemId"),
                "MS-OXWSCDATA",
                62,
                @"[In t:DefaultShapeNamesType Simple Type] The value of ""IdOnly"" specifies only the item or folder ID.");
            #endregion

            #region Step 4: Get the created item with BaseShape set to Default.
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.Default;

            GetItemResponseType getItemResponse_Default = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_Default, 1, this.Site);

            T[] item_Default = Common.GetItemsFromInfoResponse<T>(getItemResponse_Default);

            Site.Assert.AreEqual<int>(
                 1,
                 item_Default.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_Default.GetLength(0));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R61");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R61
            // The request have get item by Default,
            // and the responses are successfully,
            // this requirement can be verified.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                61,
                @"[In t:DefaultShapeNamesType Simple Type] The value ""Default"" specifies a set of properties that are defined as the default for the item or folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1185");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1185
            // The request have set BaseShape element to different value,
            // and the responses are in different shape,
            // this requirement can be verified.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1185,
                @"[In t:ItemResponseShapeType Complex Type] The element ""BaseShape"" with type ""t:DefaultShapeNamesType(section 2.2.3.7)"" Specifies the requested base properties to return in a response.");
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with different BodyTypeResponseType enumeration values in ItemShape element.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_BodyTypeResponseTypeEnum<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the created item with BodyType set to HTML.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItem.ItemShape.BodyType = BodyTypeResponseType.HTML;
            getItem.ItemShape.BodyTypeSpecified = true;

            GetItemResponseType getItemResponse_HTML = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_HTML, 1, this.Site);

            T[] item_HTML = Common.GetItemsFromInfoResponse<T>(getItemResponse_HTML);

            Site.Assert.AreEqual<int>(
                 1,
                 item_HTML.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_HTML.GetLength(0));

            Site.Assert.IsNotNull(
                item_HTML[0].Body,
                "The body element in returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R17");
            Site.Log.Add(LogEntryKind.Debug, "The BodyType of the Body should be HTML, actual {0}, the Value of the Body is {1}.", item_HTML[0].Body.BodyType1, item_HTML[0].Body.Value);

            bool isVerifyR17 = BodyTypeType.HTML == item_HTML[0].Body.BodyType1 && TestSuiteHelper.IsHTML(item_HTML[0].Body.Value);

            // The request set BodyType element to HTML value,
            // if the BodyType of the Body element in response is HTML format,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR17,
                "MS-OXWSCDATA",
                17,
                @"[In t:BodyTypeResponseType Simple Type] The value ""HTML"" specifies that the response returns an item body as HTML.");
            #endregion

            #region Step 3: Get the created item with BodyType set to Text.
            getItem.ItemShape.BodyType = BodyTypeResponseType.Text;

            GetItemResponseType getItemResponse_Text = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_Text, 1, this.Site);

            T[] item_Text = Common.GetItemsFromInfoResponse<T>(getItemResponse_Text);

            Site.Assert.AreEqual<int>(
                 1,
                 item_Text.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_Text.GetLength(0));

            Site.Assert.IsNotNull(
                item_Text[0].Body,
                "The body element in returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R18");
            Site.Log.Add(LogEntryKind.Debug, "The BodyType of the Body should be Text, actual {0}, the Value of the Body is {1}.", item_Text[0].Body.BodyType1, item_Text[0].Body.Value);

            bool isVerifyR18 = BodyTypeType.Text == item_Text[0].Body.BodyType1 && !TestSuiteHelper.IsHTML(item_Text[0].Body.Value);

            // The request set BodyType element to Text value,
            // if the BodyType of the Body element in response is Text format,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR18,
                "MS-OXWSCDATA",
                18,
                @"[In t:BodyTypeResponseType Simple Type] The value ""Text"" specifies that the response returns an item body as plain text.");
            #endregion

            #region Step 4: Get the created item with BodyType set to Best.
            getItem.ItemShape.BodyType = BodyTypeResponseType.Best;

            GetItemResponseType getItemResponse_Best = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_Best, 1, this.Site);

            T[] item_Best = Common.GetItemsFromInfoResponse<T>(getItemResponse_Best);

            Site.Assert.AreEqual<int>(
                 1,
                 item_Best.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_Best.GetLength(0));

            Site.Assert.IsNotNull(
                item_Best[0].Body,
                "The body element in returned item should not be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R15");

            // The request set BodyType element to Best value, and the value of the body is html format,
            // if the BodyType of the Body element in response is according to the html content of the Body element,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsTrue(
                TestSuiteHelper.IsHTML(item_Best[0].Body.Value),
                "MS-OXWSCDATA",
                15,
                @"[In t:BodyTypeResponseType Simple Type] The value ""Best"" specifies that the response returns the richest available body content.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1190");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1190
            // The requests set BodyType element to different values, and the values of the body in responses are in accordingly format,
            // this requirement can be verified.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                1190,
                @"[In t:ItemResponseShapeType Complex Type] The element ""BodyType"" with type ""t:BodyTypeResponseType(section 2.2.3.1)"" Specifies the requested body text format for the Body property that is returned in a response.");
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with the ItemShape element in which IncludeMimeContent element exists.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be got.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_IncludeMimeContentBoolean<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2: Get the created item with IncludeMimeContent set to true.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItem.ItemShape.IncludeMimeContent = true;
            getItem.ItemShape.IncludeMimeContentSpecified = true;

            if (Common.IsRequirementEnabled(2919, this.Site))
            {
                // Return Additional Property 'itemMimeContentUTF8'
                List<PathToUnindexedFieldType> additionalProperties = new List<PathToUnindexedFieldType>();
                additionalProperties.Add(new PathToUnindexedFieldType() { FieldURI = UnindexedFieldURIType.itemMimeContentUTF8 });
                getItem.ItemShape.AdditionalProperties = additionalProperties.ToArray();
            }

            GetItemResponseType getItemResponse_IncludeMimeContentTrue = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_IncludeMimeContentTrue, 1, this.Site);

            // Check whether the schema is validated
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Get item from GetItemResponse.
            item = Common.GetItemsFromInfoResponse<T>(getItemResponse_IncludeMimeContentTrue)[0];          

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1307");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1307
            // If the MimeContent element of the item is not null, and the schema is validated,
            // which represents the MIME content of an item is returned in a response as MimeContentType,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.MimeContent,
                1307,
                @"[In t:ItemType Complex Type] The type of MimeContent is t:MimeContentType (section 2.2.4.10).");
           
            if (Common.IsRequirementEnabled(2919, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2919");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1348
                // if the element is not null, the schema is validated, and the instanceKey is base64 binary data
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    item.MimeContentUTF8.Value,
                    2919,
                    @"[In Appendix C: Product Behavior] Implementation does support the MimeContentUTF8 element which specifies an instance of the MimeContentUTF8Type complex type that contains the native MIME stream of an object that is represented in UTF-8. (<79> Section 2.2.4.24:  Exchange 2016 and above follow this behavior.)");
            }
            
            if (Common.IsRequirementEnabled(23091, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R23091");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R23091
                // If the MimeContent element of the item is not null, and the schema is validated,
                // which represents the MIME content of an item is returned in a response as MimeContentType,
                // this requirement can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    item.MimeContent,
                    23091,
                    @"[In Appendix C: Product Behavior] This element [MimeContent] is applicable for PostItemType, MessageType, CalendarItemType, ContactType, TaskType and DistributionListType item when retrieving MIME content.(<52> Section 2.2.4.24: Exchange 2010SP3 and above follow this behavior.)");
            }
            if (Common.IsRequirementEnabled(23093, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R23093");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R23093
                this.Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                    ResponseCodeType.ErrorUnsupportedMimeConversion,
                    getItemResponse_IncludeMimeContentTrue.ResponseMessages.Items[0].ResponseCode,
                    23093,
                    @"[In Appendix C: Product Behavior] An ErrorUnsupportedMimeConversion will be returned. (<52> Section 2.2.4.24:  In Exchange 2007, Exchange 2010, Exchange 2010 SP1 and Microsoft Exchange Server 2010 Service Pack 2 (SP2), when retrieving MIME content for an item other than a PostItemType, MessageType, or CalendarItemType object, an ErrorUnsupportedMimeConversion will be returned.)");
            }
            
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R69");

            // If the value of MimeContent element is a base64 string format,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                TestSuiteHelper.IsBase64String(item.MimeContent.Value),
                69,
                @"[In t:ItemType Complex Type] [The element 'MimeContent'] Specifies an instance of the MimeContentType complex type that contains the native MIME stream of an object that is represented in base64encoding.");           

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1362");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1362
            // If the CharacterSet element of the MimeContent is not null, and the schema is validated,
            // which represents the CharacterSet of a MimeContent is returned in a response as string,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.MimeContent.CharacterSet,
                1362,
                @"[In t:MimeContentType Complex Type] The type of CharacterSet is xs:string [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R114");

            // The CharacterSet attribute use the International Standards Organization (ISO) name, but the specific name should not be checked, just validate the name is not null or empty here.
            Site.CaptureRequirementIfIsTrue(
                !string.IsNullOrEmpty(item.MimeContent.CharacterSet),
                114,
                @"[In t:MimeContentType Complex Type] [The attribute ""CharacterSet""] Specifies the International Standards Organization (ISO) name of the character set that is used in a MIME message.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21188");

            // The request set IncludeMimeContent element to true,
            // If the MimeContent element of the item is not null, which represents the MIME content of an item is returned in a response,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                item.MimeContent,
                "MS-OXWSCDATA",
                21188,
                @"[In t:ItemResponseShapeType Complex Type] [IncludeMimeContent is] True, specifies the MIME content of an item is returned in a response.");

            if (item is MessageType
                || item is PostItemType
                || item is CalendarItemType)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2012");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2012
                Site.CaptureRequirementIfIsNotNull(
                    item.MimeContent,
                    2012,
                    @"[In t:ItemType Complex Type] This element [MimeContent] is only applicable to PostItemType, MessageType, and CalendarItemType object when setting MIME content for an item. ");
            }
            #endregion

            #region Step 3: Get the created item with IncludeMimeContent set to false.
            getItem.ItemShape.IncludeMimeContent = false;
            GetItemResponseType getItemResponse_IncludeMimeContentFalse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_IncludeMimeContentFalse, 1, this.Site);

            item = Common.GetItemsFromInfoResponse<T>(getItemResponse_IncludeMimeContentFalse)[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21189");

            // The request set IncludeMimeContent element to false,
            // If the MimeContent element of the item is null, which represents the MIME content of an item is not returned in a response,
            // this requirement can be verified.
            Site.CaptureRequirementIfIsNull(
                item.MimeContent,
                "MS-OXWSCDATA",
                21189,
                @"[In t:ItemResponseShapeType Complex Type] otherwise [IncludeMimeContent is] false, specifies [the MIME content of an item is not returned in a response].");
            #endregion
        }
        
        /// <summary>
        /// Verify the responses returned by GetItem operation with the ItemShape element in which ConvertHtmlCodePageToUTF8 element exists or is not specified.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_ConvertHtmlCodePageToUTF8Boolean<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the created item without ConvertHtmlCodePageToUTF8.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            GetItemResponseType getItemResponse_ConvertHtmlCodePageToUTF8NotSpecified = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_ConvertHtmlCodePageToUTF8NotSpecified, 1, this.Site);

            T[] item_ConvertHtmlCodePageToUTF8NotSpecified = Common.GetItemsFromInfoResponse<T>(getItemResponse_ConvertHtmlCodePageToUTF8NotSpecified);

            Site.Assert.AreEqual<int>(
                 1,
                 item_ConvertHtmlCodePageToUTF8NotSpecified.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_ConvertHtmlCodePageToUTF8NotSpecified.GetLength(0));

            Site.Assert.IsNotNull(
                item_ConvertHtmlCodePageToUTF8NotSpecified[0].Body,
                "The body element in returned item should not be null.");

            string charSet_ConvertHtmlCodePageToUTF8NotSpecified = TestSuiteHelper.GetCharsetOfHTML(item_ConvertHtmlCodePageToUTF8NotSpecified[0].Body.Value);
            #endregion

            #region Step 3: Get the created item with ConvertHtmlCodePageToUTF8 set to true.
            // Set the ConvertHtmlCodePageToUTF8 property.
            getItem.ItemShape.ConvertHtmlCodePageToUTF8 = true;
            getItem.ItemShape.ConvertHtmlCodePageToUTF8Specified = true;

            GetItemResponseType getItemResponse_ConvertHtmlCodePageToUTF8True = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_ConvertHtmlCodePageToUTF8True, 1, this.Site);

            T[] item_ConvertHtmlCodePageToUTF8True = Common.GetItemsFromInfoResponse<T>(getItemResponse_ConvertHtmlCodePageToUTF8True);

            Site.Assert.AreEqual<int>(
                 1,
                 item_ConvertHtmlCodePageToUTF8True.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_ConvertHtmlCodePageToUTF8True.GetLength(0));

            Site.Assert.IsNotNull(
                item_ConvertHtmlCodePageToUTF8True[0].Body,
                "The body element in returned item should not be null.");

            string charSet_ConvertHtmlCodePageToUTF8True = TestSuiteHelper.GetCharsetOfHTML(item_ConvertHtmlCodePageToUTF8True[0].Body.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21195");

            // The request set ConvertHtmlCodePageToUTF8 element to true,
            // if the charset of the html in response body equals to "utf-8", which represents the item HTML body is converted to UTF8,
            // this requirement can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                "utf-8",
                charSet_ConvertHtmlCodePageToUTF8True,
                "MS-OXWSCDATA",
                21195,
                @"[In t:ItemResponseShapeType Complex Type] [ConvertHtmlCodePageToUTF8 is] True, specifies the item HTML body is converted to UTF8.");

            if (Common.IsRequirementEnabled(21498, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21498, Expected value: utf-8, Actual value of not specifying the element: {0}, Actual value of setting the element to true: {1}", charSet_ConvertHtmlCodePageToUTF8NotSpecified, charSet_ConvertHtmlCodePageToUTF8True.ToLower(new CultureInfo(TestSuiteHelper.Culture, false)));

                // If the responses of the request set ConvertHtmlCodePageToUTF8 element to true, and the request not specify ConvertHtmlCodePageToUTF8 element,
                // both return the "utf-8" charset,
                // this requirement can be verified.
                bool isVerifyR21498 = string.Equals("utf-8", charSet_ConvertHtmlCodePageToUTF8NotSpecified, StringComparison.OrdinalIgnoreCase)
                    && string.Equals("utf-8", charSet_ConvertHtmlCodePageToUTF8True, StringComparison.OrdinalIgnoreCase);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR21498,
                    "MS-OXWSCDATA",
                    21498,
                    @"[In Appendix C: Product Behavior] Implementation does include ConvertHtmlCodePageToUTF8 element whose value MUST be set to true or the element MUST NOT be specified to indicate to the implementation to convert the HTML code page to UTF8. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Step 4: Get the created item with ConvertHtmlCodePageToUTF8 set to false.
            getItem.ItemShape.ConvertHtmlCodePageToUTF8 = false;
            getItem.ItemShape.ConvertHtmlCodePageToUTF8Specified = true;
            GetItemResponseType getItemResponse_ConvertHtmlCodePageToUTF8False = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_ConvertHtmlCodePageToUTF8False, 1, this.Site);

            T[] item_ConvertHtmlCodePageToUTF8False = Common.GetItemsFromInfoResponse<T>(getItemResponse_ConvertHtmlCodePageToUTF8False);

            Site.Assert.AreEqual<int>(
                 1,
                 item_ConvertHtmlCodePageToUTF8False.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_ConvertHtmlCodePageToUTF8False.GetLength(0));

            Site.Assert.IsNotNull(
                item_ConvertHtmlCodePageToUTF8False[0].Body,
                "The body element in returned item should not be null.");

            string charSet_ConvertHtmlCodePageToUTF8False = TestSuiteHelper.GetCharsetOfHTML(item_ConvertHtmlCodePageToUTF8False[0].Body.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21196");

            // The request set ConvertHtmlCodePageToUTF8 element to false,
            // if the charset of the html in response body does not equal to "utf-8", which represents the item HTML body is not converted to UTF8,
            // this requirement can be verified.
            Site.CaptureRequirementIfAreNotEqual<string>(
                "utf-8",
                charSet_ConvertHtmlCodePageToUTF8False,
                "MS-OXWSCDATA",
                21196,
                @"[In t:ItemResponseShapeType Complex Type] otherwise [ConvertHtmlCodePageToUTF8 is] false, specifies [the item HTML body is not converted to UTF8].");
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with the ItemShape element in which FilterHtmlContent element exists or is not specified.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_FilterHtmlContentBoolean<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 3: Get the created item with FilterHtmlContent set to true.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            // Set the FilterHtmlContent property.
            getItem.ItemShape.FilterHtmlContent = true;
            getItem.ItemShape.FilterHtmlContentSpecified = true;

            GetItemResponseType getItemResponse_FilterHtmlContentTrue = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_FilterHtmlContentTrue, 1, this.Site);

            T[] item_FilterHtmlContentTrue = Common.GetItemsFromInfoResponse<T>(getItemResponse_FilterHtmlContentTrue);

            Site.Assert.AreEqual<int>(
                 1,
                 item_FilterHtmlContentTrue.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_FilterHtmlContentTrue.GetLength(0));

            Site.Assert.IsNotNull(
                item_FilterHtmlContentTrue[0].Body,
                "The body element in returned item should not be null.");

            bool filterHtmlContent = item_FilterHtmlContentTrue[0].Body.Value.Contains("</script>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21193");

            // Verify MS-OXWSCDATA_R21193.
            Site.CaptureRequirementIfIsFalse(
                filterHtmlContent,
                "MS-OXWSCDATA",
                21193,
                @"[In t:ItemResponseShapeType Complex Type] [FilterHtmlContent is] True, specifies HTML content filtering is enabled.");

            #endregion

            #region Step 4: Get the created item with FilterHtmlContent set to false.
            getItem.ItemShape.FilterHtmlContent = false;
            getItem.ItemShape.FilterHtmlContentSpecified = true;
            GetItemResponseType getItemResponse_FilterHtmlContentFalse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_FilterHtmlContentFalse, 1, this.Site);

            T[] item_FilterHtmlContentFalse = Common.GetItemsFromInfoResponse<T>(getItemResponse_FilterHtmlContentFalse);

            Site.Assert.AreEqual<int>(
                 1,
                 item_FilterHtmlContentFalse.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_FilterHtmlContentFalse.GetLength(0));

            Site.Assert.IsNotNull(
                item_FilterHtmlContentFalse[0].Body,
                "The body element in returned item should not be null.");

            filterHtmlContent = item_FilterHtmlContentFalse[0].Body.Value.Contains("</script>");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21194");

            // Verify MS-OXWSCDATA_R21193.
            Site.CaptureRequirementIfIsTrue(
                filterHtmlContent,
                "MS-OXWSCDATA",
                21194,
                @"[In t:ItemResponseShapeType Complex Type] otherwise [FilterHtmlContent is] false, specifies [HTML content filtering is not enabled].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2119413");

            // Verify MS-OXWSCDATA_R2119413.
            // This requirement can be captured after above steps.
            Site.CaptureRequirement(
                "MS-OXWSCDATA",
                2119413,
                @"[In Appendix C: Product Behavior] Implementation does support the FilterHtmlContent element. (Exchange 2010 and above follow this behavior.)");
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with the ItemShape element in which AddBlankTargetToLinks element exists.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_AddBlankTargetToLinksBoolean<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the created item with AddBlankTargetToLinks set to true.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;
            getItem.ItemShape.AddBlankTargetToLinks = true;
            getItem.ItemShape.AddBlankTargetToLinksSpecified = true;

            GetItemResponseType getItemResponse_AddBlankTargetToLinksTrue = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_AddBlankTargetToLinksTrue, 1, this.Site);

            T[] item_AddBlankTargetToLinksTrue = Common.GetItemsFromInfoResponse<T>(getItemResponse_AddBlankTargetToLinksTrue);

            Site.Assert.AreEqual<int>(
                 1,
                 item_AddBlankTargetToLinksTrue.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_AddBlankTargetToLinksTrue.GetLength(0));

            Site.Assert.IsNotNull(
                item_AddBlankTargetToLinksTrue[0].Body,
                "The body element in returned item should not be null.");

            string target_AddBlankTargetToLinksTrue = TestSuiteHelper.GetTargetAttribute(item_AddBlankTargetToLinksTrue[0].Body.Value);

            if (Common.IsRequirementEnabled(2149908, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2149908");

                // The request set AddBlankTargetToLinks element to true,
                // if the target attribute of the link in the html string of response contains "blank", which represents target attribute is set to a value of blank,
                // this requirement can be verified.
                Site.CaptureRequirementIfIsTrue(
                    target_AddBlankTargetToLinksTrue.Contains("blank"),
                    "MS-OXWSCDATA",
                    2149908,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""AddBlankTargetToLinks"" with type ""xs:boolean""  which is true specifying the target attribute is set to a value of blank. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Step 3: Get the created item with AddBlankTargetToLinks set to false.
            getItem.ItemShape.AddBlankTargetToLinks = false;
            getItem.ItemShape.AddBlankTargetToLinksSpecified = true;
            GetItemResponseType getItemResponse_AddBlankTargetToLinksFalse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_AddBlankTargetToLinksFalse, 1, this.Site);

            T[] item_AddBlankTargetToLinksFalse = Common.GetItemsFromInfoResponse<T>(getItemResponse_AddBlankTargetToLinksFalse);

            Site.Assert.AreEqual<int>(
                 1,
                 item_AddBlankTargetToLinksFalse.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_AddBlankTargetToLinksFalse.GetLength(0));

            Site.Assert.IsNotNull(
                item_AddBlankTargetToLinksFalse[0].Body,
                "The body element in returned item should not be null.");

            string target_AddBlankTargetToLinksFalse = TestSuiteHelper.GetTargetAttribute(item_AddBlankTargetToLinksFalse[0].Body.Value);

            if (Common.IsRequirementEnabled(2149909, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2149909");

                // The request set AddBlankTargetToLinks element to false,
                // if the target attribute of the link in the html string of response does not contain "blank", which represents target attribute is not set to a value of blank,
                // this requirement can be verified.
                Site.CaptureRequirementIfIsFalse(
                    target_AddBlankTargetToLinksFalse.Contains("blank"),
                    "MS-OXWSCDATA",
                    2149909,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""AddBlankTargetToLinks"" with type ""xs:boolean""  which is false specifying the target attribute is not set to a value of blank. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by GetItem operation with the ItemShape element in which BlockExternalImages element exists.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be gotten.</param>
        protected void TestSteps_VerifyGetItemWithItemResponseShapeType_BlockExternalImagesBoolean<T>(T item)
            where T : ItemType, new()
        {
            GetItemType getItem = new GetItemType();

            #region Step 1: Create an item.
            // Create item and return the item id.
            getItem.ItemIds = this.CreateItemForSpecificItemType(item);
            #endregion

            #region Step 2: Get the create item with BlockExternalImages set to true.
            // Set the item shape's elements.
            getItem.ItemShape = new ItemResponseShapeType();
            getItem.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            List<PathToUnindexedFieldType> pathToUnindexedFields = new List<PathToUnindexedFieldType>();

            if (Common.IsRequirementEnabled(1357, this.Site))
            {
                PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemBlockStatus;
                pathToUnindexedFields.Add(pathToUnindexedField);
            }

            if (Common.IsRequirementEnabled(1358, this.Site))
            {
                PathToUnindexedFieldType pathToUnindexedField = new PathToUnindexedFieldType();
                pathToUnindexedField.FieldURI = UnindexedFieldURIType.itemHasBlockedImages;
                pathToUnindexedFields.Add(pathToUnindexedField);
            }

            getItem.ItemShape.AdditionalProperties = pathToUnindexedFields.ToArray();

            getItem.ItemShape.BlockExternalImages = true;
            getItem.ItemShape.BlockExternalImagesSpecified = true;

            GetItemResponseType getItemResponse_BlockExternalImagesTrue = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_BlockExternalImagesTrue, 1, this.Site);

            T[] item_BlockExternalImagesTrue = Common.GetItemsFromInfoResponse<T>(getItemResponse_BlockExternalImagesTrue);

            Site.Assert.AreEqual<int>(
                 1,
                 item_BlockExternalImagesTrue.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_BlockExternalImagesTrue.GetLength(0));

            Site.Assert.IsNotNull(
                item_BlockExternalImagesTrue[0].Body,
                "The body element in returned item should not be null.");

            bool containImgSrc_BlockExternalImagesTrue = TestSuiteHelper.ContainImageSrcOfHTML(item_BlockExternalImagesTrue[0].Body.Value);

            if (Common.IsRequirementEnabled(2149905, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2149905");

                // The body of the request contains the image,
                // if the containImgSrc_BlockExternalImagesTrue is false, which represents the body of item in the response does not contain any image,
                // this requirement can be verified.
                Site.CaptureRequirementIfIsFalse(
                    containImgSrc_BlockExternalImagesTrue,
                    "MS-OXWSCDATA",
                    2149905,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""BlockExternalImages"" with type ""xs:boolean"" which is true specifying external images are blocked. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1357, this.Site))
            {
                Site.Assert.IsTrue(
                    item_BlockExternalImagesTrue[0].BlockStatusSpecified,
                    "The BlockStatus element in returned item should be specified.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1357");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1357
                Site.CaptureRequirementIfIsTrue(
                    item_BlockExternalImagesTrue[0].BlockStatus,
                    1357,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""BlockStatus"" with type ""xs:boolean"" which is true indicating images are blocked. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1358, this.Site))
            {
                Site.Assert.IsTrue(
                    item_BlockExternalImagesTrue[0].HasBlockedImagesSpecified,
                    "The HasBlockedImages element in returned item should be specified.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1358");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1358
                Site.CaptureRequirementIfIsTrue(
                    item_BlockExternalImagesTrue[0].HasBlockedImages,
                    1358,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""HasBlockedImages"" with type ""xs:boolean"" which is true indicating the item has blocked images. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion

            #region Step 3: Get the created item with BlockExternalImages set to false.
            getItem.ItemShape.BlockExternalImages = false;
            getItem.ItemShape.BlockExternalImagesSpecified = true;
            GetItemResponseType getItemResponse_BlockExternalImagesFalse = this.COREAdapter.GetItem(getItem);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse_BlockExternalImagesFalse, 1, this.Site);

            T[] item_BlockExternalImagesFalse = Common.GetItemsFromInfoResponse<T>(getItemResponse_BlockExternalImagesFalse);

            Site.Assert.AreEqual<int>(
                 1,
                 item_BlockExternalImagesFalse.GetLength(0),
                 "One item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 item_BlockExternalImagesFalse.GetLength(0));

            Site.Assert.IsNotNull(
                item_BlockExternalImagesFalse[0].Body,
                "The body element in returned item should not be null.");

            bool containImgSrc_BlockExternalImagesFalse = TestSuiteHelper.ContainImageSrcOfHTML(item_BlockExternalImagesFalse[0].Body.Value);

            if (Common.IsRequirementEnabled(2149906, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R2149906");

                // The body of the request contains the image,
                // if the containImgSrc_BlockExternalImagesFalse is true, which represents the body of item in the response contains the image,
                // this requirement can be verified.
                Site.CaptureRequirementIfIsTrue(
                    containImgSrc_BlockExternalImagesFalse,
                    "MS-OXWSCDATA",
                    2149906,
                    @"[In Appendix C: Product Behavior] Implementation does use the element ""BlockExternalImages"" with type ""xs:boolean"" which is false specifying external images are not blocked. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1622, this.Site))
            {
                Site.Assert.IsTrue(
                    item_BlockExternalImagesFalse[0].BlockStatusSpecified,
                    "The BlockStatus element in returned item should be specified.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1622");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1622
                Site.CaptureRequirementIfIsFalse(
                    item_BlockExternalImagesFalse[0].BlockStatus,
                    1622,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""BlockStatus"" with type ""xs:boolean"" which is false indicating images are not blocked. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1623, this.Site))
            {
                Site.Assert.IsTrue(
                    item_BlockExternalImagesFalse[0].HasBlockedImagesSpecified,
                    "The HasBlockedImages element in returned item should be specified.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1623");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1623
                Site.CaptureRequirementIfIsFalse(
                    item_BlockExternalImagesFalse[0].BlockStatus,
                    1623,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""HasBlockedImages"" with type ""xs:boolean"" which is false indicating the item has not blocked images. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// Verify the responses returned by MarkAllItemsAsRead operation.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="items">The items to be updated.</param>
        protected void TestSteps_VerifyMarkAllItemsAsRead<T>(T[] items)
            where T : ItemType, new()
        {
            #region Step 1:Create two items.
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the items
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();
            createItemRequest.Items.Items = items;
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                1);
            createItemRequest.Items.Items[1].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                2);

            if (createItemRequest.Items.Items[0] is CalendarItemType)
            {
                createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
                createItemRequest.SendMeetingInvitationsSpecified = true;
            }

            if (createItemRequest.Items.Items[0] is MessageType)
            {
                createItemRequest.MessageDispositionSpecified = true;
                createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;
            }
            #endregion

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

            // Check the operation response.
            Common.CheckOperationSuccess(createItemResponse, 2, this.Site);

            ItemIdType[] createdItemIds = Common.GetItemIdsFromInfoResponse(createItemResponse);

            // Two created items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    createdItemIds.GetLength(0),
                    "Two created contact item should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    createdItemIds.GetLength(0));
            #endregion

            #region Step 2:Get two items.
            // Call GetItem operation using the created item IDs.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            ItemIdType[] getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two contact items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two items should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));
            #endregion

            #region Step 3:Mark all items as unread, and suppress the receive receipts.
            BaseFolderIdType[] folderIds = new BaseFolderIdType[1];
            DistinguishedFolderIdType distinguishedFol = new DistinguishedFolderIdType();
            distinguishedFol.Id = DistinguishedFolderIdNameType.drafts;
            folderIds[0] = distinguishedFol;

            // Mark all items in drafts folder as unread, and suppress the receive receipts.
            MarkAllItemsAsReadResponseType markAllItemsAsReadResponse = this.CallMarkAllItemsAsReadOperation(false, true, folderIds);

            // Check the operation response.
            Common.CheckOperationSuccess(markAllItemsAsReadResponse, 1, this.Site);
            #endregion

            #region Step 4:Get two items.
            // Call GetItem operation using the created item IDs.
            getItemResponse = this.CallGetItemOperation(createdItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // Two items should be returned.
            Site.Assert.AreEqual<int>(
                    2,
                    getItemIds.GetLength(0),
                    "Two items should be returned! Expected Item Count: {0}, Actual Item Count: {1}",
                    2,
                    getItemIds.GetLength(0));
            #endregion
        }

        /// <summary>
        /// Verify the failed responses returned by UpdateItem operation.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be updated.</param>
        protected void TestSteps_VerifyUpdateItemFailedResponse<T>(T item)
            where T : ItemType, new()
        {
            #region Step 1:Create an item
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2:Update the item, using SetItemField element.
            // Change the Item data.
            ItemChangeType[] itemChanges = new ItemChangeType[1];
            itemChanges[0] = new ItemChangeType();
            itemChanges[0].Item = createdItemIds[0];
            itemChanges[0].Updates = new ItemChangeDescriptionType[1];
            SetItemFieldType setItem = new SetItemFieldType();
            setItem.Item = new PathToUnindexedFieldType()
            {
                FieldURI = UnindexedFieldURIType.itemSubject
            };
            setItem.Item1 = new T()
            {
                Subject = Common.GenerateResourceName(this.Site, TestSuiteHelper.SubjectForUpdateItem),
                Body = new BodyType()
            };
            itemChanges[0].Updates[0] = setItem;

            // Call UpdateItem to update the Subject and the Body of the created item simultaneously, by using ItemId in CreateItem response.
            UpdateItemResponseType updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.deleteditems,
                false,
                itemChanges);

            #endregion

            // Verify that Items element is present in the UpdateItem response and the response code is "ErrorIncorrectUpdatePropertyCount".
            foreach (ResponseMessageType responseMessage in updateItemResponse.ResponseMessages.Items)
            {
                // Verify ResponseCode is ErrorIncorrectUpdatePropertyCount.
                this.VerifyErrorIncorrectUpdatePropertyCount(responseMessage.ResponseCode);
            }
        }

        /// <summary>
        /// Verify the successful responses returned by UpdateItem operation.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be updated.</param>
        protected void TestSteps_VerifyUpdateItemSuccessfulResponse<T>(T item)
            where T : ItemType, new()
        {
            #region Step 1:Create an item
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2:Get the item
            // Call GetItem operation using the created item IDs.
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

            #region Step 3:Update the item
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
            append.Item1 = new T()
            {
                Body = new BodyType()
                {
                    BodyType1 = BodyTypeType.Text,
                    Value = TestSuiteHelper.BodyForBaseItem
                }
            };
            itemChanges[0].Updates[0] = append;

            // Call UpdateItem to update the email address of the created item, by using ItemId in CreateItem response.
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
            #endregion

            #region Step 4:Get the updated item
            // Call GetItem operation using the updated item IDs.
            getItemResponse = this.CallGetItemOperation(updatedItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 1, this.Site);

            getItemIds = Common.GetItemIdsFromInfoResponse(getItemResponse);

            // One contact item should be returned.
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
            #endregion
        }

        /// <summary>
        /// Verify the successful responses returned by CreateItem, GetItem and DeleteItem operations.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="item">The item to be operated.</param>
        protected void TestSteps_VerifyCreateGetDeleteItem<T>(T item)
            where T : ItemType, new()
        {
            #region Step 1:Create an item
            ItemIdType[] createdItemIds = this.CreateItemWithMinimumElements(item);
            #endregion

            #region Step 2:Get the item
            // Call GetItem operation using the created item IDs.
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

            #region Step3:Delete the item
            DeleteItemResponseType deleteItemResponse = this.CallDeleteItemOperation();

            // Check the operation response.
            Common.CheckOperationSuccess(deleteItemResponse, 1, this.Site);            

            // Clear ExistItemIds for DeleteItem.
            this.InitializeCollection();
            #endregion

            #region Step 4:Get the deleted contact item
            // Call GetItem operation using the deleted item IDs.
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
        }

        /// <summary>
        /// Verify the successful responses returned by CreateItem, UpdateItem, MoveItem, GetItem, CopyItem and SendItem when there are multiple distribution list, meeting, contact, post, or task items.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="items">The items to be operated.</param>
        protected void TestSteps_VerifyOperateMultipleItems<T>(T[] items)
            where T : ItemType, new()
        {
            #region Step 1:Create multiple items
            CreateItemType createItemRequest = new CreateItemType();

            #region Config the items
            createItemRequest.Items = new NonEmptyArrayOfAllItemsType();

            createItemRequest.Items.Items = items;
            createItemRequest.Items.Items[0].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                1);
            createItemRequest.Items.Items[1].Subject = Common.GenerateResourceName(
                this.Site,
                TestSuiteHelper.SubjectForCreateItem,
                2);
            createItemRequest.MessageDispositionSpecified = true;
            createItemRequest.MessageDisposition = MessageDispositionType.SaveOnly;

            if (createItemRequest.Items.Items[0] is CalendarItemType)
            {
                createItemRequest.SendMeetingInvitations = CalendarItemCreateOrDeleteOperationType.SendToAllAndSaveCopy;
                createItemRequest.SendMeetingInvitationsSpecified = true;
            }

            #endregion

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);

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

            #region Step 2 - 5: Update, move, get and copy the items.
            for (int i = 0; i < items.Length; i++)
            {
                items[i].ItemId = createdItemIds[i];
            }

            this.OperateMultipleItems(items);
            #endregion
        }

        /// <summary>
        /// Verify update, move, get and copy the multiple items.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object.</typeparam>
        /// <param name="items">The items to be updated.</param>
        protected void OperateMultipleItems<T>(T[] items)
            where T : ItemType, new()
        {
            #region Update the items
            ItemChangeType[] itemChanges = new ItemChangeType[]
            {
                TestSuiteHelper.CreateItemChangeItem(items[0], 1),
                TestSuiteHelper.CreateItemChangeItem(items[1], 2)
            };

            // Call UpdateItem to update the email address of the created item, by using ItemId in CreateItem response.
            UpdateItemResponseType updateItemResponse = this.CallUpdateItemOperation(
                DistinguishedFolderIdNameType.drafts,
                true,
                itemChanges);

            // Check the operation response.
            Common.CheckOperationSuccess(updateItemResponse, 2, this.Site);
            #endregion

            #region Move the items
            // Configure ItemIds.
            ItemIdType[] moveItemIds = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(moveItemIds, 0);

            // Clear ExistItemIds for MoveItem.
            this.InitializeCollection();

            // Call MoveItem operation.
            MoveItemResponseType moveItemResponse = this.CallMoveItemOperation(DistinguishedFolderIdNameType.inbox, moveItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(moveItemResponse, 2, this.Site);
            #endregion

            #region Get the items
            // The items to get.
            ItemIdType[] itemArray = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(itemArray, 0);

            // Call GetItem operation.
            GetItemResponseType getItemResponse = this.CallGetItemOperation(itemArray);

            // Check the operation response.
            Common.CheckOperationSuccess(getItemResponse, 2, this.Site);

            #endregion

            #region Copy the items
            // Configure ItemIds.
            ItemIdType[] copyItemIds = new ItemIdType[this.ExistItemIds.Count];
            this.ExistItemIds.CopyTo(copyItemIds, 0);
            foreach (ItemIdType copyItemId in copyItemIds)
            {
                this.CopiedItemIds.Add(copyItemId);
            }

            // Call CopyItem operation.
            CopyItemResponseType copyItemResponse = this.CallCopyItemOperation(DistinguishedFolderIdNameType.drafts, copyItemIds);

            // Check the operation response.
            Common.CheckOperationSuccess(copyItemResponse, 2, this.Site);

            #endregion
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R42 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (DistinguishedPropertySetId, PropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common, set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to StringArray.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                true,
                true,
                false,
                true);

            // Initialize request Items.
            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with four extended attributes:  DistinguishedPropertySetId, PropertySetId, PropertyId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The DistinguishedPropertySetId should not be used with PropertySetId, but DistinguishedPropertySetId is used with PropertySetId, so it violates this rule and it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (DistinguishedPropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common,
            // set PropertyId to "123" with Int32 type and set PropertyType to StringArray.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                true,
                false,
                true);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Check whether the schema is validated
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Get the Item Ids.
            T[] createItemOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemOne[0],
                "The item in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R101");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R101
            // The ExtendedProperty element is gotten in CallGetItemOperationWithAdditionalProperties method, and the schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirement(
                101,
                @"[In t:ItemType Complex Type] [The element ""ExtendedProperty""] Specifies an array of ExtendedPropertyType elements that identify extended MAPI properties.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1336");

            // The extendedPropertyOne element is not null, and the schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                extendedPropertyOne,
                1336,
                @"[In t:ItemType Complex Type] The type of ExtendedProperty is t:ExtendedPropertyType ([MS-OXWSXPROP] section 2.1.5).");

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R42 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertySetId does not exist.
            Site.Assert.IsNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "The value of PropertySetId property should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R42");
            Site.Log.Add(LogEntryKind.Debug, "The value of DistinguishedPropertySetIdSpecified should be true, actual {0}.", extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertySetId should be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            bool isVerifyR42 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified && extendedPropertyOne.ExtendedFieldURI.PropertySetId == null;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R42
            // Verify if DistinguishedPropertySetId is used, the PropertySetId attributes cannot be used.
            // Because the passed condition that DistinguishedPropertySetId exists and PropertySetId does not exist 
            // and the failed condition that both DistinguishedPropertySetId and PropertySetId exist are built in above, 
            // so MS-OXWSXPROP_R42 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR42,
                "MS-OXWSXPROP",
                42,
                @"[In t:PathToExtendedFieldType Complex Type] If this attribute [DistinguishedPropertySetId] is used, the PropertySetId attributes cannot be used.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R147");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R147
            Site.CaptureRequirementIfAreEqual(
                typeof(NonEmptyArrayOfPropertyValuesType),
                extendedPropertyOne.Item.GetType(),
                "MS-OXWSXPROP",
                147,
                @"[In t:ExtendedPropertyType Complex Type] The type of element Values is t:NonEmptyArrayOfPropertyValuesType (section 2.1.3).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R144");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R144
            // As defined in schema, the type of element Value is xs:string,
            // if the Item element is NonEmptyArrayOfPropertyValuesType type, and schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirement(
                "MS-OXWSXPROP",
                144,
                @"[In t:NonEmptyArrayOfPropertyValuesType Complex Type]The type of element Value is xs:string [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R23");
            Site.Log.Add(LogEntryKind.Debug, "The count of items in extendedPropertyOne should be at least one, actual {0}.", ((NonEmptyArrayOfPropertyValuesType)extendedPropertyOne.Item).Items.GetLength(0));

            bool isVerifyR23 = ((NonEmptyArrayOfPropertyValuesType)extendedPropertyOne.Item).Items.GetLength(0) >= 1;

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R23
            Site.CaptureRequirementIfIsTrue(
                isVerifyR23,
                "MS-OXWSXPROP",
                23,
                @"[In t:NonEmptyArrayOfPropertyValuesType Complex Type] This array [the collection of values for an extended property which is represented by NonEmptyArrayOfPropertyValuesType complex type] has at least one member.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R148");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R148
            // If the element is specified in response and the schema is validated, this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "MS-OXWSXPROP",
                148,
                @"[In t:PathToExtendedFieldType Complex Type] The type of attribute DistinguishedPropertySetId is  t:DistinguishedPropertySetType ([MS-OXWSCDATA] section 2.2.2.10).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R152");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R152
            // If the element is specified in response and the schema is validated, this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified,
                "MS-OXWSXPROP",
                152,
                @"[In t:PathToExtendedFieldType Complex Type]The type of attribute PropertyId is xs:int [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R153");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R153
            // The attribute PropertyType is a required one in schema and if the schema is validated, this requirement can be validated.
            Site.CaptureRequirement(
                "MS-OXWSXPROP",
                153,
                @"[In t:PathToExtendedFieldType Complex Type] The type of attribute PropertyType is t:MapiPropertyTypeType (section 2.1.7).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R37");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R37
            // As defined in schema, the PathToExtendedFieldType extends BasePathToElementType complex type,
            // If schema is validated, this requirement can be validated.
            Site.CaptureRequirement(
                "MS-OXWSXPROP",
                37,
                @"[In t:PathToExtendedFieldType Complex Type] The PathToExtendedFieldType complex type extends the BasePathToElementType complex type ([MS-OXWSCDATA] section 2.2.3.13).");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R161 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyDistinguishedPropertySetIdConflictsWithPropertyTag<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (DistinguishedPropertySetId, PropertyTag, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common, set PropertyTag to "0x3a45",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                true,
                false,
                true,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with four extended attributes:  DistinguishedPropertySetId, PropertyTag, PropertyId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The DistinguishedPropertySetId should not be used with PropertyTag, but DistinguishedPropertySetId is used with PropertyTag, so it violates this rule and it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (DistinguishedPropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common,
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R161 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishedPropertySetId exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetId property should be specified, its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetId.ToString());

            // Check PropertySetId does not exist.
            Site.Assert.IsNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag value should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R161");
            Site.Log.Add(LogEntryKind.Debug, "The value of DistinguishedPropertySetIdSpecified should be true, actual {0}.", extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyTag should be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            bool isVerifyR161 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified && extendedPropertyOne.ExtendedFieldURI.PropertyTag == null;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R161
            // Verify if DistinguishedPropertySetId is used, the PropertyTag attributes cannot be used.
            // If the passed condition that DistinguishedPropertySetId exists and PropertyTag does not exist
            // and the failed condition that both DistinguishedPropertySetId and PropertySetId  exist are built in above,
            // MS-OXWSXPROP_R161 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR161,
                "MS-OXWSXPROP",
                161,
                @"[In t:PathToExtendedFieldType Complex Type] If this attribute [DistinguishedPropertySetId] is used, the PropertyTag attributes cannot be used.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R43 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (DistinguishedPropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                false,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  DistinguishedPropertySetId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The DistinguishedPropertySetId is used with PropertyType, but it is not used with either PropertyId or PropertyName, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (DistinguishedPropertySetId, PropertyType, PropertyId)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common,
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R43 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishedPropertySetId exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetId property should be specified, its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetId.ToString());

            // Check PropertyName exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified,
                "The PropertyId property should be specified, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyId);

            bool isVerifyRS43 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified && this.IsSchemaValidated && (extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified == true || extendedPropertyOne.ExtendedFieldURI.PropertyName != null);

            // Verify DistinguishedPropertySetId is used with PropertyType and either PropertyId or PropertyName.
            // If the passed condition that DistinguishedPropertySetId exists and used with PropertyType and either PropertyId or PropertyName
            // and the failed condition that DistinguishedPropertySetId exists and used with PropertyType but not used with PropertyId or PropertyName are built in above,
            // MS-OXWSXPROP_R43 can be verified directly. 
            this.VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(isVerifyRS43);

            #region Step 3: Create item successfully: Set the properties (DistingushedPropertySetId, PropertyType, PropertyName)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common,
            // set PropertyName to  "Classification" and set PropertyType to String.
            T itemTwo = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                false,
                true,
                false);

            items = new T[] { itemTwo };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyName and PropertyType.
            CreateItemResponseType successResponseTwo = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponseTwo, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsTwo = Common.GetItemsFromInfoResponse<T>(successResponseTwo);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsTwo[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyTwo = this.CallGetItemOperationWithAdditionalProperties(createItemsTwo[0], itemTwo.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R43 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishedPropertySetId exists.
            Site.Assert.IsTrue(
                extendedPropertyTwo.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetId property should be specified, its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.DistinguishedPropertySetId.ToString());

            // Check PropertyName exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertyName,
                "The PropertyName property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertyName);

            // IsSchemaValidated is true means PropertyType is present, since it is a required attribute of PathToExtendedFieldType.
            isVerifyRS43 = extendedPropertyTwo.ExtendedFieldURI.DistinguishedPropertySetIdSpecified && this.IsSchemaValidated && (extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified == true || extendedPropertyOne.ExtendedFieldURI.PropertyName != null);

            // Verify DistinguishedPropertySetId is used with PropertyType and either PropertyId or PropertyName.
            // If the passed condition that DistinguishedPropertySetId exists and used with PropertyType and either PropertyId or PropertyName
            // and the failed condition that DistinguishedPropertySetId exists and used with PropertyType but not used with PropertyId or PropertyName are built in above,
            // MS-OXWSXPROP_R43 can be verified directly. 
            this.VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(isVerifyRS43);
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R157 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertySetIdConflictsWithDistinguishedPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (PropertySetId, DistinguishedPropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common, set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                true,
                true,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  DistinguishedPropertySetId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The DistinguishedPropertySetId should not be used with PropertySetId, but DistinguishedPropertySetId is used with PropertySetId, so it violates this rule and it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Check whether the schema is validated
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R157 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishedPropertySetId does not exist.
            Site.Assert.IsFalse(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The value of DistinguishedPropertySetId should be false.");

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R157");
            Site.Log.Add(LogEntryKind.Debug, "The value of DistinguishedPropertySetIdSpecified should be false, actual {0}.", extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertySetId should be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            bool isVerifyR157 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified == false && extendedPropertyOne.ExtendedFieldURI.PropertySetId != null;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R157
            // Verify PropertySetId is not used with DistinguishedPropertySetId.
            // If the passed condition that PropertySetId is not used with DistinguishedPropertySetId is verified with successfully response here,
            // and the failed condition that PropertySetId is used with DistinguishedPropertySetId is verified with error response in above,
            // MS-OXWSXPROP_R157 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR157,
                "MS-OXWSXPROP",
                157,
                @"[In t:PathToExtendedFieldType Complex Type] If this attribute [PropertySetId] is used, the DistinguishedPropertySetId  attributes cannot be used.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R149");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R149
            // If the element is not null in response and the schema is validated, this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "MS-OXWSXPROP",
                149,
                @"[In t:PathToExtendedFieldType Complex Type]  The type of PropretySetId is t:GuidType (section 2.1.6).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R124");

            // The regular expression of variable length properties.
            Regex isPropertySetId = new Regex("[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R124
            Site.CaptureRequirementIfIsTrue(
                isPropertySetId.IsMatch(extendedPropertyOne.ExtendedFieldURI.PropertySetId),
                "MS-OXWSXPROP",
                124,
                @"[In t:GuidType Simple Type] The following pattern is defined by the GuidType simple type: 
                [0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R158 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertySetIdConflictsWithPropertyTag<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (PropertySetId, PropertyTag, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common, set PropertyTag to "0x3a45",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                true,
                true,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  DistinguishedPropertySetId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyTag should not be used with PropertySetId, but PropertyTag is used with PropertySetId, so it violates this rule and it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertySetId, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  DistinguishedPropertySetId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R158 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag does not exist.
            Site.Assert.IsNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag value should be false.");

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R158");
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyTag should be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyTag);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertySetId should not be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            bool isVerifyR158 = extendedPropertyOne.ExtendedFieldURI.PropertyTag == null && extendedPropertyOne.ExtendedFieldURI.PropertySetId != null;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R158
            // Verify PropertySetId is not used with PropertyTag.
            // If the passed condition that PropertySetId is not used with PropertyTag 
            // and the failed condition that PropertySetId is used with PropertyTag are both built successfully in above,
            // MS-OXWSXPROP_R158 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR158,
                "MS-OXWSXPROP",
                158,
                @"[In t:PathToExtendedFieldType Complex Type] If this attribute [PropertySetId] is used, the PropertyTag attributes cannot be used.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R47 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertySetIdWithPropertyTypeOrPropertyName<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: Set the properties (PropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                false,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertySetId, and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertySetId is used with PropertyType, but it is not used with either PropertyId or PropertyName, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertySetId, PropertyType, PropertyId)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  PropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R47 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            // Check PropertyName exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified,
                "The PropertyId property should be specified, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyId);

            bool isVerifyRS47 = extendedPropertyOne.ExtendedFieldURI.PropertySetId != null && this.IsSchemaValidated && extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified;

            // Verify PropertySetId is used with PropertyType and either PropertyId or PropertyName.
            // If the passed condition that PropertySetId exists and used with PropertyType and either PropertyId or PropertyName
            // and the failed condition that PropertySetId exists and used with PropertyType but not used with PropertyId or PropertyName are both built successfully in above,
            // MS-OXWSXPROP_R47 can be verified directly. 
            this.VerifyPropertySetIdWithPropertyTypeOrPropertyName(isVerifyRS47);

            #region Step 3: Create item successfully: Set the properties (PropertySetId, PropertyType, PropertyName)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyName to "Classification" and set PropertyType to String.
            T itemTwo = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                false,
                true,
                false);

            items = new T[] { itemTwo };

            // Call CreateItem to create an item with three extended attributes:  PropertySetId, PropertyName and PropertyType.
            CreateItemResponseType successResponseTwo = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponseTwo, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsTwo = Common.GetItemsFromInfoResponse<T>(successResponseTwo);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsTwo[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyTwo = this.CallGetItemOperationWithAdditionalProperties(createItemsTwo[0], itemTwo.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R47 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId);

            // Check PropertyName exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertyName,
                "The PropertyName property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertyName);

            isVerifyRS47 = extendedPropertyTwo.ExtendedFieldURI.PropertySetId != null && this.IsSchemaValidated && extendedPropertyTwo.ExtendedFieldURI.PropertyName != null;

            // Verify PropertySetId is used with PropertyType and either PropertyId or PropertyName.
            // If the passed condition that PropertySetId exists and used with PropertyType and either PropertyId or PropertyName
            // and the failed condition that PropertySetId exists and used with PropertyType but not used with PropertyId or PropertyName are both built successfully in above,
            // MS-OXWSXPROP_R47 can be verified directly. 
            this.VerifyPropertySetIdWithPropertyTypeOrPropertyName(isVerifyRS47);
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R50 and related requirements about the PathToExtendedFieldType complex type with successful response returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyTagRepresentation<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item successfully: Set the properties (PropertyTag, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyTag to "0x3a45" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertyTag and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R50 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            // Check whether the schema is verified.
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R50");

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R50
            // Verify PropertyTag exists.
            // If PropertyTag exists and the schema is verified, MS-OXWSXPROP_R50 can be verified directly. 
            // The schema contain the pattern of "(0x|0X)[0-9A-Fa-f]{1,4}", so if the schema can be verified, then this requirement can be verified.
            Site.CaptureRequirementIfIsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "MS-OXWSXPROP",
                50,
                @"[In t:PathToExtendedFieldType Complex Type] The PropertyTag attribute can be represented as either a hexadecimal value or a short integer.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R51 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyTagConflictsWithDistinguishedPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyTag, DistinguishedPropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishedPropertySetId to Common,
            // set PropertyTag to "0x3a45" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                true,
                false,
                false,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  DistinguishedPropertySetId, PropertyTag and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyTag is used with PropertyType, but it is also used with DistinguishedPropertySetId, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyTag, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyTag to "0x3a45" and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertyTag and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R51 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            // Check DistinguishedPropertySetId does not exist.
            Site.Assert.IsFalse(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetIdSpecified value should be false.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R51");
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyTag should not be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyTag);
            Site.Log.Add(LogEntryKind.Debug, "The value of DistinguishedPropertySetIdSpecified should be false, actual {0}.", extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified);

            bool isVerifyR51 = extendedPropertyOne.ExtendedFieldURI.PropertyTag != null && extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified == false;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R51
            // Verify PropertyTag is not used with DistinguishedPropertySetId.
            // If the passed condition that PropertyTag exists and the DistinguishedPropertySetId does not exist 
            // and the failed condition that PropertyTag and DistinguishedPropertySetId both exist are built in above, 
            // MS-OXWSXPROP_R51 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR51,
                "MS-OXWSXPROP",
                51,
                @"[In t:PathToExtendedFieldType Complex Type] If the PropertyTag attribute is used, the DistinguishedPropertySetId MUST NOT be used.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R138 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyTagConflictsWithPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyTag, PropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyTag to "0x3a45" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                true,
                false,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  PropertyTag, PropertySetId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyTag is used with PropertyType, but it is also used with PropertySetId, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyTag, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyTag to "0x3a45" and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertyTag and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Check whether the schema is validated
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R138 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            // Check PropertySetId does not exist.
            Site.Assert.IsNull(
                extendedPropertyOne.ExtendedFieldURI.PropertySetId,
                "The PropertySetId value should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R138");
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyTag should not be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyTag);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertySetId should be null or empty, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            bool isVerifyR138 = extendedPropertyOne.ExtendedFieldURI.PropertyTag != null && string.IsNullOrEmpty(extendedPropertyOne.ExtendedFieldURI.PropertySetId);

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R138
            // Verify PropertyTag is not used with PropertySetId.
            // If the passed condition that PropertyTag exists and the PropertySetId does not exist are verified with successful response.
            // and the failed condition that PropertyTag and PropertySetId both exist are verified with error response in above,
            // MS-OXWSXPROP_R138 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR138,
                "MS-OXWSXPROP",
                138,
                @"[In t:PathToExtendedFieldType Complex Type]If the PropertyTag attribute is used, the PropertySetId MUST NOT be used.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R122");

            // The regular expression of variable length properties.
            Regex isPropertyTag = new Regex("(0x|0X)[0-9A-Fa-f]{1,4}");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R122
            Site.CaptureRequirementIfIsTrue(
                isPropertyTag.IsMatch(extendedPropertyOne.ExtendedFieldURI.PropertyTag),
                "MS-OXWSXPROP",
                122,
                @"[In t:PropertyTagType Simple Type] The following pattern is defined by the PropertyTagType simple type: 
                (0x|0X)[0-9A-Fa-f]{1,4}.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R145");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R145
            Site.CaptureRequirementIfIsTrue(
                this.IsSchemaValidated,
                "MS-OXWSXPROP",
                145,
                @"[In t:ExtendedPropertyType Complex Type] The type of element ExtendedFieldURI is t:PathToExtendedFieldType (section 2.1.5).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R146");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R146

            // In schema, the ExtendedPropertyType type could have an element Value as string type or Values element as NonEmptyArrayOfPropertyValuesType type
            // in this test case it should be the element Value as string type.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(string),
                extendedPropertyOne.Item.GetType(),
                "MS-OXWSXPROP",
                146,
                @"[In t:ExtendedPropertyType Complex Type] The type of element Value is xs:string [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R150");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R150
            // If the element is not null in response and the schema is validated, this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "MS-OXWSXPROP",
                150,
                @"[In t:PathToExtendedFieldType Complex Type]The type of attribute PropertyTag is  t:PropertyTagType (section 2.1.8).");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R139 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyTagConflictsWithPropertyName<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyTag, PropertyName, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyName to "Classification",
            // set PropertyTag to "0x3a45" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                true,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  PropertyTag, PropertyName and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyTag is used with PropertyType, but it is also used with PropertyName, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyTag, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyTag to "0x3a45" and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertyTag and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R139 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            // Check PropertyName does not exist.
            Site.Assert.IsNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyName,
                "The PropertyName value should be null.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R139");
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyTag should not be null, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyTag);
            Site.Log.Add(LogEntryKind.Debug, "The value of PropertyName should be null or empty, actual {0}.", extendedPropertyOne.ExtendedFieldURI.PropertyName);

            bool isVerifyR139 = extendedPropertyOne.ExtendedFieldURI.PropertyTag != null && string.IsNullOrEmpty(extendedPropertyOne.ExtendedFieldURI.PropertyName);

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R139
            // Verify PropertyTag is not used with PropertyName.
            // If the passed condition that PropertyTag exists and the PropertyName does not exist
            // and the failed condition that PropertyTag and the PropertyName both exist are built in above,
            // MS-OXWSXPROP_R139 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR139,
                "MS-OXWSXPROP",
                139,
                @"[In t:PathToExtendedFieldType Complex Type] If the PropertyTag attribute is used, the PropertyName MUST NOT be used.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R140 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyTagConflictsWithPropertyId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyTag, PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyId to "123" with Int32 type,
            // set PropertyTag to "0x3a45" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                true,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes:  PropertyTag, PropertyId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyTag is used with PropertyType, but it is also used with PropertyId, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyTag, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyTag to "0x3a45" and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                true,
                false,
                false,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes:  PropertyTag and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids. 
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R140 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyTag exists.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyTag,
                "The PropertyTag property should not be null, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.PropertyTag);

            // Check PropertyId does not exist.
            Site.Assert.IsFalse(
                extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified,
                "The PropertyIdSpecified value should be false.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R140");

            bool isVerifyR140 = extendedPropertyOne.ExtendedFieldURI.PropertyTag != null && extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified == false;

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R140
            // Verify PropertyTag is not used with PropertyId.
            // If the passed condition that PropertyTag exists and the PropertyId does not exist
            // and the failed condition that PropertyTag and PropertyId both exist are built,
            // MS-OXWSXPROP_R140 can be verified directly. 
            Site.CaptureRequirementIfIsTrue(
                isVerifyR140,
                "MS-OXWSXPROP",
                140,
                @"[In t:PathToExtendedFieldType Complex Type]If the PropertyTag attribute is used, the PropertyId MUST NOT be used.");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R53 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyName, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyName to "Classification" and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                false,
                false,
                true,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes: PropertyName and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyName is not used with either DistinguishedPropertySetId or PropertySetId, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyName, DistinguishPropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishPropertySetId to Common, set PropertyName to "Classification" and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                false,
                true,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes: PropertyName, DistinguishPropertySetId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);
            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R53 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishPropertySetId exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetId property should be specified, and its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetId.ToString());

            // Check PropertyName exists and is equal with set in CreateItem.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI.PropertyName,
                "The PropertyName property should not be null, and its value is {0}",
                extendedPropertyOne.ExtendedFieldURI.PropertyName);

            bool isVerifyRS53 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified == true && extendedPropertyOne.ExtendedFieldURI.PropertyName != null;

            // Verify PropertyName is used with PropertyType and either DistinguishPropertySetId or PropertySetId.
            // If the passed condition that PropertyName exists and used with PropertyType and either DistinguishPropertySetId or PropertySetId
            // and the failed condition that PropertyName exists and used with PropertyType but not used with either DistinguishedPropertySetId or PropertySetId are built in above, 
            // MS-OXWSXPROP_R53 can be verified directly. 
            this.VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(isVerifyRS53);

            #region Step 3: Create item successfully: Set the properties (PropertyName, PropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e", 
            // set PropertyName to "Classification" and set PropertyType to String.
            T itemTwo = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                false,
                true,
                false);

            items = new T[] { itemTwo };

            // Call CreateItem to create an item with three extended attributes:  PropertySetId, PropertyName and PropertyType.
            CreateItemResponseType successResponseTwo = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponseTwo, 1, this.Site);

            // Check whether the schema is validated
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            // Get the Item Ids.
            T[] createItemsTwo = Common.GetItemsFromInfoResponse<T>(successResponseTwo);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsTwo[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyTwo = this.CallGetItemOperationWithAdditionalProperties(createItemsTwo[0], itemTwo.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R53 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId);

            // Check PropertyName exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertyName,
                "The PropertyName property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertyName);

            isVerifyRS53 = extendedPropertyTwo.ExtendedFieldURI.PropertySetId != null && extendedPropertyTwo.ExtendedFieldURI.PropertyName != null;

            // Verify PropertyName is used with PropertyType and either DistinguishPropertySetId or PropertySetId.
            // If the passed condition that PropertyName exists and used with PropertyType and either DistinguishPropertySetId or PropertySetId
            // and the failed condition that PropertyName exists and used with PropertyType but not used with either DistinguishedPropertySetId or PropertySetId are built in above, 
            // MS-OXWSXPROP_R53 can be verified directly. 
            this.VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(isVerifyRS53);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R151");

            // Verify MS-OXWSCORE requirement: MS-OXWSXPROP_R151
            // If the element is not null in response and the schema is validated, this requirement can be validated.
            Site.CaptureRequirementIfIsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertyName,
                "MS-OXWSXPROP",
                151,
                @"[In t:PathToExtendedFieldType Complex Type] The type of attribute PropertyName is  xs:string [XMLSCHEMA2].");
        }

        /// <summary>
        /// Test steps for validating MS-OXWSXPROP_R55 and related requirements about the PathToExtendedFieldType complex type with both failed and successful responses returned by CreateItem operation for generic item types.
        /// </summary>
        /// <typeparam name="T">The ItemType or its child class object</typeparam>
        /// <param name="folderId">The folder Id</param>
        /// <param name="item">The item to be verified.</param>
        protected void TestSteps_VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId<T>(DistinguishedFolderIdNameType folderId, T item)
            where T : ItemType, new()
        {
            #region Step 1: Create item failed: (PropertyId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertyId to "123" with Int32 type and set PropertyType to String.
            T itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                false,
                true,
                false,
                false);

            T[] items = new T[] { itemOne };

            // Call CreateItem to create an item with two extended attributes: PropertyId and PropertyType.
            CreateItemResponseType failedResponse = this.CallCreateItemOperation(folderId, items);

            Site.Assert.AreEqual<int>(
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0),
                 "Expected Item Count: {0}, Actual Item Count: {1}",
                 1,
                 failedResponse.ResponseMessages.Items.GetLength(0));

            // Check whether the CreateItem operation is executed in error.
            Site.Assert.AreEqual<ResponseClassType>(
                ResponseClassType.Error,
                failedResponse.ResponseMessages.Items[0].ResponseClass,
                "The PropertyId is not used with either DistinguishedPropertySetId or PropertySetId, so it should be failed.");

            #endregion

            #region Step 2: Create item successfully: Set the properties (PropertyId, DistinguishPropertySetId PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set DistinguishPropertySetId to Common,
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            itemOne = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                true,
                false,
                false,
                true,
                false,
                false);

            items = new T[] { itemOne };

            // Call CreateItem to create an item with three extended attributes: PropertyId, DistinguishPropertySetId and PropertyType.
            CreateItemResponseType successResponse = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponse, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsOne = Common.GetItemsFromInfoResponse<T>(successResponse);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsOne[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyOne = this.CallGetItemOperationWithAdditionalProperties(createItemsOne[0], itemOne.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R55 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check DistinguishedPropertySetId exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified,
                "The DistinguishedPropertySetId property should be specified, its value is {0}.",
                extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetId.ToString());

            // Check PropertyId exists.
            Site.Assert.IsTrue(
                extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified,
                "The PropertyId should be specified, its value is {0}",
                extendedPropertyOne.ExtendedFieldURI.PropertyId.ToString());

            bool isVerifyRS55 = extendedPropertyOne.ExtendedFieldURI.DistinguishedPropertySetIdSpecified && extendedPropertyOne.ExtendedFieldURI.PropertyIdSpecified;

            // Verify PropertyId is used with PropertyType and either DistinguishPropertySetId or PropertySetId.
            // If the passed condition that PropertyId exists and used with PropertyType and either DistinguishPropertySetId or PropertySetId
            // and the failed condition that PropertyId exists and used with PropertyType but not used either DistinguishPropertySetId or PropertySetId are built in above, 
            // MS-OXWSXPROP_R55 can be verified directly. 
            this.VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(isVerifyRS55);

            #region Step 3: Create item successfully: Set the properties (PropertyId, PropertySetId, PropertyType)

            // Set the properties of PathToExtendedFieldType.
            // False means the property will not be specified, otherwise, it will be set.
            // Set PropertySetId to "c11ff724-aa03-4555-9952-8fa248a11c3e",
            // set PropertyId to "123" with Int32 type and set PropertyType to String.
            T itemTwo = this.SetPathToExtendedFieldTypeProperties<T>(
                item,
                false,
                false,
                true,
                true,
                false,
                false);

            items = new T[] { itemTwo };

            // Call CreateItem to create an item with three extended attributes:  PropertySetId, PropertyId and PropertyType.
            CreateItemResponseType successResponseTwo = this.CallCreateItemOperation(folderId, items);

            // Check the operation response.
            Common.CheckOperationSuccess(successResponseTwo, 1, this.Site);

            // Get the Item Ids.
            T[] createItemsTwo = Common.GetItemsFromInfoResponse<T>(successResponseTwo);

            // Check whether the ItemId is not null
            Site.Assert.IsNotNull(
                createItemsTwo[0],
                "The ItemId in the successful response of CreateItem operation should not be null.");

            ExtendedPropertyType extendedPropertyTwo = this.CallGetItemOperationWithAdditionalProperties(createItemsTwo[0], itemTwo.ExtendedProperty[0].ExtendedFieldURI);

            #endregion

            // If the conditions above have been built successfully, the requirement of MS-OXWSXPROP_R55 can be captured.
            // Check ExtendedFieldURI is not null.
            Site.Assert.IsNotNull(
                extendedPropertyOne.ExtendedFieldURI,
                "The value of ExtendedFieldURI property should not be null.");

            // Check PropertyId exists
            Site.Assert.IsTrue(
                extendedPropertyTwo.ExtendedFieldURI.PropertyIdSpecified,
                "The PropertyId property should be specified, its value is {0}",
                extendedPropertyTwo.ExtendedFieldURI.PropertyId.ToString());

            // Check PropertySetId exists.
            Site.Assert.IsNotNull(
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId,
                "The PropertySetId property should not be null, and its value is {0}.",
                extendedPropertyTwo.ExtendedFieldURI.PropertySetId);

            isVerifyRS55 = extendedPropertyTwo.ExtendedFieldURI.PropertyIdSpecified && extendedPropertyTwo.ExtendedFieldURI.PropertySetId != null;

            // Verify PropertyId is used with PropertyType and either DistinguishPropertySetId or PropertySetId.
            // If the passed condition that PropertyId exists and used with PropertyType and either DistinguishPropertySetId or PropertySetId
            // and the failed condition that PropertyId exists and used with PropertyType but not used either DistinguishPropertySetId or PropertySetId are built in above, 
            // MS-OXWSXPROP_R55 can be verified directly. 
            this.VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(isVerifyRS55);
        }
        #endregion
        #endregion

        #region Capture methods

        /// <summary>
        /// Capture method for ErrorIncorrectUpdatePropertyCount in ResponseCodeType.
        /// </summary>
        /// <param name="responseCode">Update item response.</param>
        protected void VerifyErrorIncorrectUpdatePropertyCount(ResponseCodeType responseCode)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R335");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R335
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                ResponseCodeType.ErrorIncorrectUpdatePropertyCount,
                responseCode,
                "MS-OXWSCDATA",
                335,
                @"[In m:ResponseCodeType Simple Type] The value ""ErrorIncorrectUpdatePropertyCount"" specifies that each change description in an UpdateItem or UpdateFolder method call MUST list only one property to be updated.");
        }

        /// <summary>
        /// Verify DistinguishedPropertySetId must be used with PropertyType or PropertyName.
        /// </summary>
        /// <param name="isVerify">Indicate whether DistinguishedPropertySetId must be used with PropertyType or PropertyName.</param>
        protected void VerifyDistinguishedPropertySetIdWithPropertyTypeOrPropertyName(bool isVerify)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R43");

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R43
            Site.CaptureRequirementIfIsTrue(
                isVerify,
                "MS-OXWSXPROP",
                43,
                @"[In t:PathToExtendedFieldType Complex Type] This attribute [DistinguishedPropertySetId] MUST be used with the PropertyType attribute and either the PropertyId or PropertyName attribute.");
        }

        /// <summary>
        /// Verify PropertySetId must be used with PropertyType or PropertyName.
        /// </summary>
        /// <param name="isVerify">Indicate whether PropertySetId must be used with PropertyType or PropertyName.</param>
        protected void VerifyPropertySetIdWithPropertyTypeOrPropertyName(bool isVerify)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R47");

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R47
            Site.CaptureRequirementIfIsTrue(
                isVerify,
                "MS-OXWSXPROP",
                47,
                @"[In t:PathToExtendedFieldType Complex Type] This attribute [PropertySetId] MUST be used with the PropertyType attribute and either the PropertyId or PropertyName attribute.");
        }

        /// <summary>
        /// Verify PropertyName must be coupled with DistinguishedPropertySetId or PropertySetId.
        /// </summary>
        /// <param name="isVerify">Indicate whether PropertyName must be coupled with DistinguishedPropertySetId or PropertySetId.</param>
        protected void VerifyPropertyNameWithDistinguishedPropertySetIdOrPropertySetId(bool isVerify)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R53");

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R53
            Site.CaptureRequirementIfIsTrue(
                isVerify,
                "MS-OXWSXPROP",
                53,
                @"[In t:PathToExtendedFieldType Complex Type] This attribute [PropertyName] MUST be coupled with either the DistinguishedPropertySetId or PropertySetId attribute.");
        }

        /// <summary>
        /// Verify PropertyId must be coupled with DistinguishedPropertySetId or PropertySetId.
        /// </summary>
        /// <param name="isVerify">Indicate whether PropertyId must be coupled with DistinguishedPropertySetId or PropertySetId.</param>
        protected void VerifyPropertyIdWithDistinguishedPropertySetIdOrPropertySetId(bool isVerify)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSXPROP_R55");

            // Verify MS-OXWSXPROP requirement: MS-OXWSXPROP_R55
            Site.CaptureRequirementIfIsTrue(
                isVerify,
                "MS-OXWSXPROP",
                55,
                @"[In t:PathToExtendedFieldType Complex Type] PropertyId MUST be coupled with either the DistinguishedPropertySetId or PropertySetId attribute.");
        }

        /// <summary>
        /// Capture method for MS-OXWSCDATA_R619.
        /// </summary>
        /// <param name="responseCode">Indicate whether the Requirement is verified.</param>
        protected void VerifyErrorObjectTypeChanged(ResponseCodeType responseCode)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R619");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R619
            Site.CaptureRequirementIfAreEqual<ResponseCodeType>(
                   ResponseCodeType.ErrorObjectTypeChanged,
                   responseCode,
                   "MS-OXWSCDATA",
                   619,
                   @"[In m:ResponseCodeType Simple Type] [ErrorObjectTypeChanged:] For the CreateItem method, the ItemClass property MUST be consistent with the strongly typed item such as a Message or Contact.");
        }

        /// <summary>
        /// Verify the accuracy of MS-OXWSITEMID RLE decompressing method
        /// </summary>
        /// <param name="itemId">An ItemIdType object returned in response</param>
        protected void VerifyRLEDecompress(ItemIdType itemId)
        {
            ItemIdId itemIdId = this.ITEMIDAdapter.ParseItemId(itemId);
            if (itemIdId.CompressionByte == 1)
            {
                byte[] bytesId = Convert.FromBase64String(itemId.Id);
                byte[] decompressedBytes = this.ITEMIDAdapter.Decompress(bytesId, 383).ToArray();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R18, the returned compressed Id byte array length is {0}, while the decompressed byte array (excluding compress byte) length is {1}", bytesId.Length, decompressedBytes.Length);

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R18
                Site.CaptureRequirementIfIsTrue(
                    decompressedBytes.Length + 1 > bytesId.Length,
                    "MS-OXWSITEMID",
                    18,
                    "[In Compression Type (byte)] If the compressed Id is smaller than the uncompressed Id, the value of this byte [Compression Type] is 1, indicating that RLE compression was used.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R17");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R17
                // MS-OXWSITEMID_R18 is verified means MS-OXWSITEMID_R17 can be captured directly
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    17,
                    "[In Compression Type (byte)] If RLE compression is used, then for each Id generated the full Id is compressed (minus the compression byte) and compared with the size of the uncompressed Id.");

                byte[] newBytesId = new byte[decompressedBytes.Length + 1];
                newBytesId[0] = 0;
                Array.Copy(decompressedBytes, 0, newBytesId, 1, decompressedBytes.Length);
                ItemIdType newItemId = new ItemIdType();
                newItemId.ChangeKey = itemId.ChangeKey;
                newItemId.Id = Convert.ToBase64String(newBytesId);
                GetItemResponseType getItemResponse = this.CallGetItemOperation(new ItemIdType[] { newItemId });

                Common.CheckOperationSuccess(getItemResponse, 1, this.Site);
                this.ExistItemIds.Remove(newItemId);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R22");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R22
                // Successfully calling GetItem with decompressed Id indicates the algorithm for decompressing
                // is correct, so R22 can be captured.
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    22,
                    @"[In Id Decompression Algorithm] The following code describes the algorithm for decompressing the Id.
                        /// <summary>
                        /// Decompresses the passed byte array using RLE scheme.
                        /// </summary>
                        /// <param name=""input"">Bytes to decompress</param>
                        /// <param name=""maxLength"">Max allowed length for the byte array</param>
                        /// <returns>decompressed bytes</returns>
                        ///
                        public MemoryStream Decompress(byte[] input, int maxLength)
                        {
                            // It can't be assumed that the compressed data size must be less than maxLength.
                            // If the compressed data consists of a series of double characters
                            // followed by a 0 character count, compressed data will be larger than 
                            // decompressed. (i.e. xx0 decompresses to xx.)
                            //
                            int initialStreamSize = Math.Min(input.Length, maxLength);

                            MemoryStream stream = new MemoryStream(initialStreamSize);
                            BinaryWriter writer = new BinaryWriter(stream);

                            // Ignore the first byte, which the caller used to identify the compression
                            // scheme.
                            //
                            for (int i = 1; i < input.Length; ++i)
                            {
                                // If this byte differs from the following one (or it's at the end of the
                                //  array), then just output the byte.
                                if (i == input.Length - 1 ||
                                    input[i] != input[i + 1])
                                {
                                    writer.Write(input[i]);
                                }
                                else // input[i] == input[i+1]
                                {
                                    // Because repeat characters are always followed by a character count,
                                    // if i == input.Length - 2, the character count is missing & the id is 
                                    // invalid.
                                    //
                                    if (i == input.Length - 2)
                                    {
                                        throw new InvalidIdMalformedException();
                                    }

                                    // The bytes are the same. Read the third byte to see how many additional
                                    // times to write the byte (over and above the two that are already 
                                    // there).
                                    //
                                    byte runLength = input[i + 2];
                                    for (int j = 0; j < runLength + 2; ++j)
                                    {
                                        writer.Write(input[i]);
                                    }

                                    // Skip the duplicate byte and the run length.
                                    //
                                    i += 2;
                                }

                                if (stream.Length > maxLength)
                                {
                                    throw new InvalidIdMalformedException();
                                }
                            }

                            writer.Flush();
                            stream.Position = 0L;
                            return stream;
                        }
                    }");
  
                byte[] compressedBytes = this.ITEMIDAdapter.Compress(newBytesId, 1);
                Site.Assert.AreEqual<int>(
                    bytesId.Length,
                    compressedBytes.Length,
                    "The compressed Id length should be {0}, actually {1}",
                    bytesId.Length,
                    compressedBytes.Length);

                for (int i = 0; i < bytesId.Length; i++)
                {
                    Site.Assert.AreEqual<byte>(
                        bytesId[i],
                        compressedBytes[i],
                        "Compressed byte at the position {0} should be {1}, actually {2}",
                        i,
                        bytesId[i],
                        compressedBytes[i]);
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R21");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R21
                // R21 can be captured after the asserts above succeed
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    21,
                    @"[In Id Compression Algorithm] The following code describes the algorithm for compressing the Id.

                        /// <summary>
                        /// Simple RLE compressor for item IDs. Bytes that do not repeat are written directly.
                        /// Bytes that repeat more than once are written twice, followed by the number of 
                        /// additional times to write the byte (i.e., total run length minus two).
                        /// </summary>
                        internal class RleCompressor
                        {
                            /// <summary>
                            /// Compresses the passed byte array using a simple RLE compression scheme.
                            /// </summary>
                            /// <param name=""streamIn"">input stream to compress</param>
                            /// <param name=""compressorId"">id of the compressor</param>
                            /// <param name=""outBytesRequired"">The number of bytes in the returned, 
                            ///              compressed byte array.</param>
                            /// <returns>compressed bytes</returns>
                            ///
                            public byte[] Compress(byte[] streamIn, byte compressorId, out int outBytesRequired)
                            {
                                byte[] streamOut = new byte[streamIn.Length];
                                outBytesRequired = streamIn.Length;
                                int index = 0;
                                streamOut[index++] = compressorId;
                                if (index == streamIn.Length)
                                {
                                    return streamIn;
                                }

                                //  Ignore the first byte, because it is a placeholder for the compression tag.
                                //  Keep a placeholder so that, if the caller ends up not doing any compression
                                //  at all, they can simply put the compression tag for ""NoCompression"" in the 
                                //  first byte and everything works.
                                //
                                byte[] input = streamIn;

                                for (int runStart = 1; runStart < (int)streamIn.Length; /* runStart incremented below */)
                                {
                                    // Always write the start character.
                                    //
                                    streamOut[index++] = input[runStart];
                                    if (index == streamIn.Length)
                                    {
                                        return streamIn;
                                    }
                
                                    //  Now look for a run of more than one character. The maximum run to be 
                                    //  handled at once is the maximum value that can be written out in an 
                                    //  (unsigned) byte _or_ the maximum remaining input, whichever is smaller.
                                    //  One caveat is that only the run length _minus two_ is written, 
                                    //  because the two characters that indicate a run are not written. So 
                                    //  Byte.MaxValue + 2 can be handled.
                                    //
                                    int maxRun = Math.Min(Byte.MaxValue + 2, (int)streamIn.Length - runStart);
                                    int runLength = 1;
                                    for (runLength = 1;
                                        runLength < maxRun && input[runStart] == input[runStart + runLength];
                                        ++runLength)
                                    {
                                        // Nothing.
                                    }

                                    // Is this a run of more than one byte?
                                    //
                                    if (runLength > 1)
                                    {
                                        //  Yes, write the byte again, followed by the number of additional
                                        //  times to write the byte (which is the total run length minus 2,
                                        //  because the byte has already been written twice).
                                        //
                                        streamOut[index++] = input[runStart];
                                        if (index == streamIn.Length)
                                        {
                                            return streamIn;
                                        }

                                        ExAssert.Assert(runLength <= Byte.MaxValue + 2, ""total run length exceeds."");
                                        streamOut[index++] = (byte)(runLength - 2);
                                        if (index == streamIn.Length)
                                        {
                                            return streamIn;
                                        }
                                    }

                                    // Move to the first byte following the run.
                                    //
                                    runStart += runLength;
                                }

                                outBytesRequired = index;
                                return streamOut;
                            }");
            }
        }

        /// <summary>
        /// Verify the accuracy of MS-OXWSITEMID RLE compressing method
        /// </summary>
        /// <param name="itemId">An ItemIdType object returned in response</param>
        protected void VerifyRLECompress(ItemIdType itemId)
        {
            ItemIdId itemIdId = this.ITEMIDAdapter.ParseItemId(itemId);
            if (itemIdId.CompressionByte == 0)
            {
                byte[] bytesId = Convert.FromBase64String(itemId.Id);

                byte[] compressedBytes = this.ITEMIDAdapter.Compress(bytesId, 1);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R19");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R19
                Site.CaptureRequirementIfAreEqual<int>(
                    bytesId.Length,
                    compressedBytes.Length,
                    "MS-OXWSITEMID",
                    19,
                    "[In Compression Type (byte)] Otherwise [the compressed Id is greater than or equal to the uncompressed Id], the value of this byte [Compression Type] is 0, indicating that no compressions was used.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R17");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R17
                // R19 is verified means R17 can be captured directly
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    17,
                    "[In Compression Type (byte)] If RLE compression is used, then for each Id generated the full Id is compressed (minus the compression byte) and compared with the size of the uncompressed Id.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSITEMID_R21");

                // Verify MS-OXWSITEMID requirement: MS-OXWSITEMID_R21
                // R19 can be captured means the compressing algorithm is correct, 
                // so R21 can be captured
                Site.CaptureRequirement(
                    "MS-OXWSITEMID",
                    21,
                    @"[In Id Compression Algorithm] The following code describes the algorithm for compressing the Id.

                        /// <summary>
                        /// Simple RLE compressor for item IDs. Bytes that do not repeat are written directly.
                        /// Bytes that repeat more than once are written twice, followed by the number of 
                        /// additional times to write the byte (i.e., total run length minus two).
                        /// </summary>
                        internal class RleCompressor
                        {
                            /// <summary>
                            /// Compresses the passed byte array using a simple RLE compression scheme.
                            /// </summary>
                            /// <param name=""streamIn"">input stream to compress</param>
                            /// <param name=""compressorId"">id of the compressor</param>
                            /// <param name=""outBytesRequired"">The number of bytes in the returned, 
                            ///              compressed byte array.</param>
                            /// <returns>compressed bytes</returns>
                            ///
                            public byte[] Compress(byte[] streamIn, byte compressorId, out int outBytesRequired)
                            {
                                byte[] streamOut = new byte[streamIn.Length];
                                outBytesRequired = streamIn.Length;
                                int index = 0;
                                streamOut[index++] = compressorId;
                                if (index == streamIn.Length)
                                {
                                    return streamIn;
                                }

                                //  Ignore the first byte, because it is a placeholder for the compression tag.
                                //  Keep a placeholder so that, if the caller ends up not doing any compression
                                //  at all, they can simply put the compression tag for ""NoCompression"" in the 
                                //  first byte and everything works.
                                //
                                byte[] input = streamIn;

                                for (int runStart = 1; runStart < (int)streamIn.Length; /* runStart incremented below */)
                                {
                                    // Always write the start character.
                                    //
                                    streamOut[index++] = input[runStart];
                                    if (index == streamIn.Length)
                                    {
                                        return streamIn;
                                    }
                
                                    //  Now look for a run of more than one character. The maximum run to be 
                                    //  handled at once is the maximum value that can be written out in an 
                                    //  (unsigned) byte _or_ the maximum remaining input, whichever is smaller.
                                    //  One caveat is that only the run length _minus two_ is written, 
                                    //  because the two characters that indicate a run are not written. So 
                                    //  Byte.MaxValue + 2 can be handled.
                                    //
                                    int maxRun = Math.Min(Byte.MaxValue + 2, (int)streamIn.Length - runStart);
                                    int runLength = 1;
                                    for (runLength = 1;
                                        runLength < maxRun && input[runStart] == input[runStart + runLength];
                                        ++runLength)
                                    {
                                        // Nothing.
                                    }

                                    // Is this a run of more than one byte?
                                    //
                                    if (runLength > 1)
                                    {
                                        //  Yes, write the byte again, followed by the number of additional
                                        //  times to write the byte (which is the total run length minus 2,
                                        //  because the byte has already been written twice).
                                        //
                                        streamOut[index++] = input[runStart];
                                        if (index == streamIn.Length)
                                        {
                                            return streamIn;
                                        }

                                        ExAssert.Assert(runLength <= Byte.MaxValue + 2, ""total run length exceeds."");
                                        streamOut[index++] = (byte)(runLength - 2);
                                        if (index == streamIn.Length)
                                        {
                                            return streamIn;
                                        }
                                    }

                                    // Move to the first byte following the run.
                                    //
                                    runStart += runLength;
                                }

                                outBytesRequired = index;
                                return streamOut;
                            }");
            }
        }

        #region ItemIdType Structure
        /// <summary>
        /// Verify the ItemIdType structure
        /// </summary>
        /// <param name="itemIdResponse">An ItemIdType instance returned in response.</param>
        protected void VerifyItemIdType(ItemIdType itemIdResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1308");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1308
            Site.CaptureRequirementIfIsNotNull(
                itemIdResponse,
                1308,
                @"[In t:ItemType Complex Type] The type of ItemId is t:ItemIdType (section 2.2.4.19).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R2015");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R2015
            Site.CaptureRequirementIfIsNotNull(
                itemIdResponse,
                2015,
                @"[In t:ItemType Complex Type] This element [ItemId] can be returned by the server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1385");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1385
            this.Site.CaptureRequirementIfIsNotNull(
                itemIdResponse.Id,
                1385,
                @"[In t:ItemIdType Complex Type] The type of Id is xs:string [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1386");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1386
            this.Site.CaptureRequirementIfIsNotNull(
                itemIdResponse.ChangeKey,
                1386,
                @"[In t:ItemIdType Complex Type] The type of ChangeKey is xs:string.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R70");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R70
            // The Id and change key child elements of the ItemId are validated in above,
            // so this requirement can be validated.
            Site.CaptureRequirement(
                70,
                @"[In t:ItemType Complex Type] [The element ""ItemId""] Specifies an instance of the ItemIdType class that represents the unique identifier and change key of an item in the server data store.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R153");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R153
            // The ItemIdType complex type extends the BaseItemIdType complex type as defined in schema, if schema is validated,
            // this requirement can be validated.
            Site.CaptureRequirement(
                153,
                @"[In t:ItemIdType Complex Type] The ItemIdType complex type extends the BaseItemIdType complex type ([MS-OXWSCDATA] section 2.2.4.13).");
        }
        #endregion

        #region FolderIdType Structure
        /// <summary>
        /// Verify the FolderIdType structure
        /// </summary>
        /// <param name="parentFolderIdResponse">A FolderIdType instance returned in response.</param>
        protected void VerifyFolderIdType(FolderIdType parentFolderIdResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1309");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1309
            Site.CaptureRequirementIfIsNotNull(
                parentFolderIdResponse,
                1309,
                @"[In t:ItemType Complex Type] The type of ParentFolderId is t:FolderIdType ([MS-OXWSCDATA] section 2.2.4.35).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1628");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1628
            Site.CaptureRequirementIfIsNotNull(
                parentFolderIdResponse.Id,
                "MS-OXWSCDATA",
                1628,
                @"[In t:FolderIdType Complex Type] The attribute ""Id"" is ""xs:string"" type ([XMLSCHEMA2]).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1629");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1629
            Site.CaptureRequirementIfIsNotNull(
                parentFolderIdResponse.ChangeKey,
                "MS-OXWSCDATA",
                1629,
                @"[In t:FolderIdType Complex Type] The attribute ""ChangeKey"" is ""xs:string"" type.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R71");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R71
            // The parentFolderIdResponse element is the child element of the item, and its type is validated in schema,
            // so this requirement can be validated.
            Site.CaptureRequirement(
                71,
                @"[In t:ItemType Complex Type] [The element ""ParentFolderId""] Specifies an instance of the FolderIdType class that represents the identifier of the parent folder that contains an item or folder.");
        }
        #endregion

        #region ItemClassType Structure
        /// <summary>
        /// Verify the ItemClassType structure
        /// </summary>
        /// <param name="itemClassResponse">A string value of ItemClass returned in response.</param>
        /// <param name="itemClassRequest">A string value of ItemClass in request.</param>
        protected void VerifyItemClassType(string itemClassResponse, string itemClassRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1310");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1310
            Site.CaptureRequirementIfIsNotNull(
                itemClassResponse,
                1310,
                @"[In t:ItemType Complex Type] The type of ItemClass is t:ItemClassType (section 2.2.5.4).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R72");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R72
            Site.CaptureRequirementIfAreEqual<string>(
                itemClassRequest,
                itemClassResponse,
                72,
                @"[In t:ItemType Complex Type] [The element ""ItemClass""] Specifies a string value that indicates the message class of an item.");
        }
        #endregion

        #region Subject Structure
        /// <summary>
        /// Verify the Subject structure
        /// </summary>
        /// <param name="subjectResponse">A string value of subject returned in response.</param>
        /// <param name="subjectRequest">A string value of subject in request.</param>
        protected void VerifySubject(string subjectResponse, string subjectRequest)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1311");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1311
            // The schema is validated and the subjectResponse is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                subjectResponse,
                1311,
                @"[In t:ItemType Complex Type] The type of  Subject is xs:string [XMLSCHEMA2].");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R73");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R73
            this.Site.CaptureRequirementIfAreEqual<string>(
                subjectRequest,
                subjectResponse,
                73,
                @"[In t:ItemType Complex Type] [The element ""Subject""] Specifies a string value that represents the subject property of items.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R74, the actual subject is '{0}'", subjectResponse);

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R74
            this.Site.CaptureRequirementIfIsTrue(
                subjectResponse.Length <= 255,
                74,
                @"[In t:ItemType Complex Type] This value [Subject] is limited to 255 characters.");
        }
        #endregion

        #region SensitivityChoicesType Structure
        /// <summary>
        /// Verify the SensitivityChoicesType structure
        /// </summary>
        /// <param name="sensitivityResponse">A enumeration value of Sensitivity returned in response.</param>
        /// <param name="sensitivityRequest">A enumeration value of Sensitivity in request.</param>
        protected void VerifySensitivityChoicesType(SensitivityChoicesType sensitivityResponse, SensitivityChoicesType sensitivityRequest)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R75");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R75
            this.Site.CaptureRequirementIfAreEqual<SensitivityChoicesType>(
                sensitivityRequest,
                sensitivityResponse,
                75,
                @"[In t:ItemType Complex Type] [The element ""Sensitivity""] Specifies one of the valid SensitivityChoicesType simple type enumeration values that indicates the sensitivity level of an item.");
        }
        #endregion

        #region BodyType Structure
        /// <summary>
        /// Verify the BodyType structure
        /// </summary>
        /// <param name="bodyResponse">A BodyType instance returned in response.</param>
        /// <param name="bodyRequest">A BodyType instance in request.</param>
        protected void VerifyBodyType(BodyType bodyResponse, BodyType bodyRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1665");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1665
            // The schema is validated and the bodyResponse is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                bodyResponse,
                "MS-OXWSCDATA",
                1665,
                @"[In t:BodyType Complex Type] The attribute ""BodyType"" is  ""t:BodyTypeType"" type.");

            Site.Assert.AreEqual<BodyTypeType>(
                bodyRequest.BodyType1,
                bodyResponse.BodyType1,
                string.Format("The value of BodyType1 should be {0}, actual {1}.", bodyRequest.BodyType1, bodyResponse.BodyType1));

            Site.Assert.AreEqual<string>(
                bodyRequest.Value,
                bodyResponse.Value,
                string.Format("The value of Body should be {0}, actual {1}.", bodyRequest.Value, bodyResponse.Value));

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R76");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R76
            // The value of BodyType1 from response is equal to the value of BodyType1 from the request,
            // and the value of Body from response is equal to the value of Body from request,
            // so this requirement can be captured.
            Site.CaptureRequirement(
                76,
                @"[In t:ItemType Complex Type] [The element ""Body""] Specifies the body content of an item.");

            // Verify the BodyTypeType schema.
            this.VerifyBodyTypeType(bodyResponse);
        }
        #endregion

        #region BodyTypeType Structure
        /// <summary>
        /// Verify the BodyTypeType structure
        /// </summary>
        /// <param name="body">A BodyType instance.</param>
        protected void VerifyBodyTypeType(BodyType body)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1665");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R1665
            // The schema is validated and the body is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                body,
                "MS-OXWSCDATA",
                1665,
                @"[In t:BodyType Complex Type] The attribute ""BodyType"" is  ""t:BodyTypeType"" type.");
        }
        #endregion

        #region ArrayOfStringsType Structure
        /// <summary>
        /// Verify the ArrayOfStringsType structure
        /// </summary>
        /// <param name="categoriesResponse">A string array values of categories returned in response.</param>
        /// <param name="categoriesRequest">A string array values of categories in request.</param>
        protected void VerifyArrayOfStringsType(string[] categoriesResponse, string[] categoriesRequest)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1317");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1317
            // The schema is validated and the categoriesResponse is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                categoriesResponse,
                1317,
                @"[In t:ItemType Complex Type] The type of Categories is t:ArrayOfStringsType ([MS-OXWSCDATA] section 2.2.4.11).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R1541");

            // Verify MS-OXWSCDATA requirement: MS-OXWSCDATA_R1541
            // The schema is validated and the categoriesResponse is not null, so this requirement can be captured.
            Site.CaptureRequirementIfIsNotNull(
                categoriesResponse,
                "MS-OXWSCDATA",
                1541,
                @"[In t:ArrayOfStringsType Complex Type] The element ""String"" is  xs:string type ([XMLSCHEMA2]).");

            bool areCategriesEqual = true;

            // Check if all categories are match in request and response
            if (categoriesResponse.GetLength(0) == categoriesRequest.GetLength(0))
            {
                for (int i = 0; i < categoriesResponse.GetLength(0); i++)
                {
                    if (categoriesResponse[i] != categoriesRequest[i])
                    {
                        areCategriesEqual = false;
                        break;
                    }
                }
            }
            else
            {
                areCategriesEqual = false;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R80");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R80
            Site.CaptureRequirementIfIsTrue(
                areCategriesEqual,
                80,
                @"[In t:ItemType Complex Type] [The element ""Categories""] Specifies a string array that identifies the categories to which an item in a mailbox belongs.");
        }
        #endregion

        #region ImportanceChoicesType Structure
        /// <summary>
        /// Verify the ImportanceChoicesType structure
        /// </summary>
        /// <param name="isImportanceSpecifiedResponse">Indicate whether ImportanceChoicesType is specified in response.</param>
        /// <param name="importanceResponse">A enumeration value of Importance returned in response.</param>
        /// <param name="importanceRequest">A enumeration value of Importance in request.</param>
        protected void VerifyImportanceChoicesType(bool isImportanceSpecifiedResponse, ImportanceChoicesType importanceResponse, ImportanceChoicesType importanceRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1318");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1318
            Site.CaptureRequirementIfIsTrue(
                isImportanceSpecifiedResponse,
                1318,
                @"[In t:ItemType Complex Type] The type of Importance is t:ImportanceChoicesType (section 2.2.5.3).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R81");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R81
            this.Site.CaptureRequirementIfAreEqual<ImportanceChoicesType>(
                importanceRequest,
                importanceResponse,
                81,
                @"[In t:ItemType Complex Type] [The element ""Importance""] Specifies one of the valid ImportanceChoicesType values to indicate the importance of an item.");
        }
        #endregion

        #region InReplyTo Structure
        /// <summary>
        /// Verify the InReplyTo structure
        /// </summary>
        /// <param name="replyToResponse">A string value of InReplyTo returned in response.</param>
        /// <param name="replyToRequest">A string value of InReplyTo in request.</param>
        protected void VerifyInReplyTo(string replyToResponse, string replyToRequest)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1319");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1319
            // The schema is validated and the replyToResponse is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                replyToResponse,
                1319,
                @"[In t:ItemType Complex Type] The type of InReplyTo is xs:string.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R82");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R82
            this.Site.CaptureRequirementIfAreEqual<string>(
                replyToRequest,
                replyToResponse,
                82,
                @"[In t:ItemType Complex Type] [The element ""InReplyTo""] Specifies a string value that contains the identifier of the item to which this item is a reply.");
        }
        #endregion

        #region ReminderMinutesBeforeStartType Structure
        /// <summary>
        /// Verify the ReminderMinutesBeforeStartType structure
        /// </summary>
        /// <param name="reminderMinutesBeforeStartResponse">The string value of ReminderMinutesBeforeStart element returned in response.</param>
        /// <param name="reminderMinutesBeforeStartRequest">The string value of ReminderMinutesBeforeStart element in request.</param>
        protected void VerifyReminderMinutesBeforeStartType(string reminderMinutesBeforeStartResponse, string reminderMinutesBeforeStartRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1331");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1331
            // This requirement can be validated after the schema is validate and the element is not null.
            Site.CaptureRequirement(
                1331,
                @"[In t:ItemType Complex Type] The type of ReminderMinutesBeforeStart is t:ReminderMinutesBeforeStartType (section 2.2.5.5).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R95");
            Site.Log.Add(LogEntryKind.Debug, "The value of reminderMinutesBeforeStartResponse should not be null, actual {0}.", reminderMinutesBeforeStartResponse);
            Site.Log.Add(LogEntryKind.Debug, "The value of ReminderMinutesBeforeStart from response should be consistent with request, expected {0}, actual {1}.", reminderMinutesBeforeStartRequest, reminderMinutesBeforeStartResponse);
            
            int reminderMinutesBeforeStartInt;

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R95
            bool isVerifyR95 = reminderMinutesBeforeStartResponse != null
                && reminderMinutesBeforeStartResponse == reminderMinutesBeforeStartRequest
                && int.TryParse(reminderMinutesBeforeStartResponse, out reminderMinutesBeforeStartInt);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR95,
                95,
                @"[In t:ItemType Complex Type] [The element ""ReminderMinutesBeforeStart""] Specifies an int value that indicates the number of minutes before an event occurs when a reminder is displayed.");
        }
        #endregion

        #region NonEmptyArrayOfResponseObjectsType Structure
        /// <summary>
        /// Verify the NonEmptyArrayOfResponseObjectsType structure
        /// </summary>
        /// <param name="responseObjects">An array of ResponseObjectType instances.</param>
        protected void VerifyNonEmptyArrayOfResponseObjectsType(ResponseObjectType[] responseObjects)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1328");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1328
            // The schema is validated and the responseObjects is not null, so this requirement can be captured.
            this.Site.CaptureRequirementIfIsNotNull(
                responseObjects,
                1328,
                @"[In t:ItemType Complex Type] The type of ResponseObjects is t:NonEmptyArrayOfResponseObjectsType (section 2.2.4.13).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R91");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R91
            this.Site.CaptureRequirementIfIsTrue(
                responseObjects.Length > 0,
                91,
                @"[In t:ItemType Complex Type] [The element ""ResponseObjects""] Specifies an array of type ResponseObjectType that contains a collection of all the response objects that are associated with an item.");
        }
        #endregion

        #region ReminderDueBy Structure
        /// <summary>
        /// Verify the ReminderDueBy structure
        /// </summary>
        /// <param name="reminderDueBySpecifiedResponse">Indicate whether ReminderDueBy is specified in response.</param>
        /// <param name="reminderDueByResponse">A DateTime value of ReminderDueBy returned in response.</param>
        /// <param name="reminderDueByRequest">A DateTime value of ReminderDueBy in request.</param>
        protected void VerifyReminderDueBy(bool reminderDueBySpecifiedResponse, DateTime reminderDueByResponse, DateTime reminderDueByRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1329");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1329
            Site.CaptureRequirementIfIsTrue(
                reminderDueBySpecifiedResponse,
                1329,
                @"[In t:ItemType Complex Type] The type of ReminderDueBy is xs:dateTime.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R92");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R92
            this.Site.CaptureRequirementIfAreEqual<string>(
                reminderDueByRequest.ToUniversalTime().ToString(),
                reminderDueByResponse.ToUniversalTime().ToString(),
                92,
                @"[In t:ItemType Complex Type] [The element ""ReminderDueBy""] Specifies an instance of the DateTime structure that represents the date and time when an event is to occur.");
        }
        #endregion

        #region ReminderIsSet Structure
        /// <summary>
        /// Verify the ReminderIsSet structure
        /// </summary>
        /// <param name="reminderIsSetSpecifiedResponse">Indicate whether ReminderIsSet is specified in response.</param>
        protected void VerifyReminderIsSet(bool reminderIsSetSpecifiedResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1330");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1330
            Site.CaptureRequirementIfIsTrue(
                reminderIsSetSpecifiedResponse,
                1330,
                @"[In t:ItemType Complex Type] The type of ReminderIsSet is xs:boolean.");
        }
        #endregion

        #region DisplayTo Structure
        /// <summary>
        /// Verify the DisplayTo structure
        /// </summary>
        /// <param name="displayToResponse">A string value of DisplayTo returned in response.</param>
        protected void VerifyDisplayTo(string displayToResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1334");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1334
            Site.CaptureRequirementIfIsNotNull(
                displayToResponse,
                1334,
                @"[In t:ItemType Complex Type] The type of DisplayTo is xs:string.");
        }
        #endregion

        #region DisplayCc Structure
        /// <summary>
        /// Verify the DisplayCc structure
        /// </summary>
        /// <param name="displayCcResponse">A string value of DisplayCc returned in response.</param>
        protected void VerifyDisplayCc(string displayCcResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1333");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1333
            Site.CaptureRequirementIfIsNotNull(
                displayCcResponse,
                1333,
                @"[In t:ItemType Complex Type] The type of DisplayCc is  xs:string.");
        }
        #endregion

        #region Culture Structure
        /// <summary>
        /// Verify the Culture structure
        /// </summary>
        /// <param name="cultureResponse">A string value of Culture returned in response.</param>
        /// <param name="cultureRequest">A string value of Culture in request.</param>
        protected void VerifyCulture(string cultureResponse, string cultureRequest)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1337");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1337
            Site.CaptureRequirementIfIsNotNull(
                cultureResponse,
                1337,
                @"[In t:ItemType Complex Type] The type of Culture is xs:language [XMLSCHEMA2].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R102");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R102
            Site.CaptureRequirementIfAreEqual<string>(
                cultureRequest.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                cultureResponse.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                102,
                @"[In t:ItemType Complex Type] [The element ""Culture""] Specifies the culture for an item in a mailbox.");
        }
        #endregion

        #region LastModifiedName Structure
        /// <summary>
        /// Verify the LastModifiedName structure
        /// </summary>
        /// <param name="lastModifiedNameResponse">A string value of LastModifiedName returned in response.</param>
        /// <param name="actualUserName">A string value of actual user name used for editing the item</param>
        protected void VerifyLastModifiedName(string lastModifiedNameResponse, string actualUserName)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1339");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1339
            Site.CaptureRequirementIfIsNotNull(
                lastModifiedNameResponse,
                1339,
                @"[In t:ItemType Complex Type] The type of LastModifiedName is xs:string.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R104");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R104
            Site.CaptureRequirementIfAreEqual<string>(
                lastModifiedNameResponse.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                actualUserName.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                104,
                @"[In t:ItemType Complex Type] [The element ""LastModifiedName""] Specifies a string value that contains the name of the user who last modified an item.");
        }
        #endregion

        #region EffectiveRightsType Structure
        /// <summary>
        /// Verify the EffectiveRightsType structure
        /// </summary>
        /// <param name="effectiveRights">An EffectiveRightsType instance.</param>
        protected void VerifyEffectiveRightsType(EffectiveRightsType effectiveRights)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1338");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1338
            Site.CaptureRequirementIfIsNotNull(
                effectiveRights,
                1338,
                @"[In t:ItemType Complex Type] The type of EffectiveRights is t:EffectiveRightsType ([MS-OXWSCDATA] section 2.2.4.25).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21131");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21131
            // The item always has CreateAssociated property equal to false, since CreateAssociated is a folder permission.
            Site.CaptureRequirementIfIsFalse(
                effectiveRights.CreateAssociated,
                "MS-OXWSCDATA",
                21131,
                @"[In t:EffectiveRightsType Complex Type] otherwise [CreateAssociated is] false, specifies [a client can not create an associated contents table].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21133");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21133
            // The item always has CreateContents property equal to false, since CreateContents is a folder permission.
            Site.CaptureRequirementIfIsFalse(
                effectiveRights.CreateContents,
                "MS-OXWSCDATA",
                21133,
                @"[In t:EffectiveRightsType Complex Type] otherwise [CreateContents is] false, specifies [a client can not create a contents table].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21135");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21135
            // The item always has CreateHierarchy property equal to false, since CreateHierarchy is a folder permission.
            Site.CaptureRequirementIfIsFalse(
                effectiveRights.CreateHierarchy,
                "MS-OXWSCDATA",
                21135,
                @"[In t:EffectiveRightsType Complex Type] otherwise [CreateHierarchy is] false, specifies [a client can not create a hierarchy table].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21136");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21136
            // This test case is run by a user with permission of deleting items,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                effectiveRights.Delete,
                "MS-OXWSCDATA",
                21136,
                @"[In t:EffectiveRightsType Complex Type] [Delete is] True, specifies a client can delete a folder or item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21138");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21138
            // This test case is run by a user with permission of modifying items,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                effectiveRights.Modify,
                "MS-OXWSCDATA",
                21138,
                @"[In t:EffectiveRightsType Complex Type] [Modify is] True, specifies a client can modify a folder or item.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21140");

            // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21140
            // This test case is run by a user with permission of reading items,
            // this requirement can be validated.
            Site.CaptureRequirementIfIsTrue(
                effectiveRights.Read,
                "MS-OXWSCDATA",
                21140,
                @"[In t:EffectiveRightsType Complex Type] [Read is] True, specifies a client can read a folder or item.");

            if (Common.IsRequirementEnabled(21142, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCDATA_R21142");

                // Verify MS-OXWSCORE requirement: MS-OXWSCDATA_R21142
                // This test case is run by a user with permission of viewing private items,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsTrue(
                    effectiveRights.ViewPrivateItems,
                    "MS-OXWSCDATA",
                    21142,
                    @"[In Appendix C: Product Behavior]  Implementation does include the element ""ViewPrivateItems"" with type ""xs:boolean"" which is true specifying a client can read private items. (Exchange Server 2013 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R103");

            // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R103
            // The Delete, Modify, Read and ViewPrivateItems permissions represents the client's rights for an item.
            // This requirement can be validated.
            Site.CaptureRequirement(
                103,
                @"[In t:ItemType Complex Type] [The element ""EffectiveRights""] Specifies an EffectiveRightsType element that represents the client's rights based on the permission settings for an item.");
        }
        #endregion

        #region FlagType Structure
        /// <summary>
        /// Verify the FlagType structure
        /// </summary>
        /// <param name="flagType">An FlagType instance.</param>
        protected void VerifyFlagType(FlagType flagType)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1346, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1042");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1042
                // FlagStatus is a required child element of flagType,
                // if schema is validated and flagType is not null, this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    flagType,
                    1042,
                    @"[In t:FlagType Complex Type] FlagStatus: An element of type FlagStatusType, as defined in [MS-OXWSCONV] section 3.1.4.1.4.2, that represents the flag status.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1043");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1043
                Site.CaptureRequirementIfIsTrue(
                    flagType.StartDateSpecified,
                    1043,
                    @"[In t:FlagType Complex Type] StartDate: An element of type dateTime, as defined in [XMLSCHEMA2] section 3.2.7, that represents the start date.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1044");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1044
                Site.CaptureRequirementIfIsTrue(
                    flagType.DueDateSpecified,
                    1044,
                    @"[In t:FlagType Complex Type] DueDate: An element of type dateTime that represents the due date.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1346");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1346
                // The child elements of "Flag" element have been validated in above,
                // so this requirement can be validated.
                Site.CaptureRequirement(
                    1346,
                    @"[In Appendix C: Product Behavior] Implementation does support the element ""Flag"" with type ""t:FlagType"" which specifies a flag indicating status, start date, due date or completion date for an item . (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1502, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1502");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1502
                // The child elements of "Flag" element have been validated in above,
                // so this requirement can be validated.
                Site.CaptureRequirement(
                    1502,
                    @"[In Appendix C: Product Behavior] Implementation does support FlagType complex type which specifies a flag indicating status, start date, due date or completion date for an item. (Exchange 2013 and above follow this behavior.)");
            }
        }
        #endregion

        #region EntityExtractionResultType structure
        /// <summary>
        /// Verify the EntityExtractionResultType structure.
        /// </summary>
        /// <param name="entityExtractionResult">An EntityExtractionResultType instance.</param>
        protected void VerifyEntityExtractionResultType(EntityExtractionResultType entityExtractionResult)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1288, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1288");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1288
                // If the EntityExtractionResultType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult,
                    1288,
                    @"[In Appendix C: Product Behavior] Implementation does support the EntityExtractionResultType complex type specifies the result of an entity extraction. (Exchange 2013 and above follow this behavior.)
  <xs:complexType name=""EntityExtractionResultType""> 
  <xs:sequence>
    <xs:element name=""Addresses"" type=""t:ArrayOfAddressEntitiesType"" minOccurs=""0"" maxOccurs=""1"" />
    <xs:element name=""MeetingSuggestions"" type=""t:ArrayOfMeetingSuggestionsType"" minOccurs=""0"" maxOccurs=""1"" />
    <xs:element name=""TaskSuggestions"" type=""t:ArrayOfTaskSuggestionsType"" minOccurs=""0"" maxOccurs=""1"" />
    <xs:element name=""EmailAddresses"" type=""t:ArrayOfEmailAddressEntitiesType"" minOccurs=""0"" maxOccurs=""1"" />
    <xs:element name=""Contacts"" type=""t:ArrayOfContactsType"" minOccurs=""0"" maxOccurs=""1"" />     
    <xs:element name=""Urls"" type=""t:ArrayOfUrlEntitiesType"" minOccurs=""0"" maxOccurs=""1"" /> 
    <xs:element name=""PhoneNumbers"" type=""t:ArrayOfPhoneEntitiesType"" minOccurs=""0"" maxOccurs=""1"" />
   </xs:sequence>
 </xs:complexType>");
            }

            if (Common.IsRequirementEnabled(1350, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1350");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1350
                // If the EntityExtractionResultType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult,
                    1350,
                    @"[In Appendix C: Product Behavior] Implementation does support element  ""EntityExtractionResult"" with type ""t:EntityExtractionResultType (section 2.2.4.35)"" which specifies the result of an entity extraction . (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1708, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1708");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1708
                // If the EntityExtractionResultType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult,
                    1708,
                    @"[In Appendix C: Product Behavior] Implementation does support EntityExtractionResultType complex type which specifies the result of an entity extraction. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1135");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1135
                // If the MeetingSuggestions element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult.MeetingSuggestions,
                    1135,
                    @"[In t:EntityExtractionResultType Complex Type] MeetingSuggestions: An element of type ArrayOfMeetingSuggestionsType, as defined in section 2.2.4.24, that represents the meeting suggestions returned.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1136");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1136
                // If the TaskSuggestions element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult.TaskSuggestions,
                    1136,
                    @"[In t:EntityExtractionResultType Complex Type] TaskSuggestions: An element of type ArrayOfTaskSuggestionsType, as defined in section 2.2.4.26, that represents the task suggestions returned.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1138");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1138
                // If the Contacts element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    entityExtractionResult.Contacts,
                    1138,
                    @"[In t:EntityExtractionResultType Complex Type] Contacts: An element of type ArrayOfContactsType, as defined in section 2.2.4.27, that represents the contacts returned.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfAddressEntitiesType structure.
        /// </summary>
        /// <param name="arrayOfAddressEntities">An array of AddressEntityType instance which is transformed from ArrayOfAddressEntitiesType.</param>
        /// <param name="address">Address information extracted from the message.</param>
        protected void VerifyArrayOfAddressEntitiesType(AddressEntityType[] arrayOfAddressEntities, string address)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1712, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1712");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1712
                // If the array of AddressEntityType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfAddressEntities,
                    1712,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfAddressEntitiesType complex type which specifies an array of address entities. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1714, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1714");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1714
                // If the array element of AddressEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfAddressEntities[0],
                    1714,
                    @"[In Appendix C: Product Behavior] Implementation does support the AddressEntityType complex type which specifies an address entity. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1750");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1750
                // If the array element of AddressEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfAddressEntities[0],
                    1750,
                    @"[In t:ArrayOfAddressEntitiesType Complex Type] AddressEntity: An element of type AddressEntityType, as defined in section 2.2.4.40, that specifies an address.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1754");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1754
                this.Site.CaptureRequirementIfAreEqual<string>(
                    address,
                    arrayOfAddressEntities[0].Address,
                    1754,
                    @"[In t:AddressEntityType Complex Type] Address: An element of type string , as defined in [XMLSCHEMA2] section 3.2.1, that represents a street address.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1134");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1134
                this.Site.CaptureRequirementIfAreEqual<string>(
                    address,
                    arrayOfAddressEntities[0].Address,
                    1134,
                    @"[In t:EntityExtractionResultType Complex Type] Addresses: An element of type ArrayOfAddressEntitiesType, as defined in section 2.2.4.2, that represents the address entities returned.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfMeetingSuggestionsType structure.
        /// </summary>
        /// <param name="arrayOfMeetingSuggestions">An array of MeetingSuggestionType instance which is transformed from ArrayOfMeetingSuggestionsType.</param>
        /// <param name="meetingSuggestion">Meeting suggestion string extracted from the message.</param>
        /// <param name="startTime">Start time of the meeting.</param>
        /// <param name="endTime">End time of the meeting.</param>
        /// <param name="attendeeName">Attendee name of the meeting.</param>
        /// <param name="attendeeEmail">Attendee email of the meeting.</param>
        protected void VerifyArrayOfMeetingSuggestionsType(MeetingSuggestionType[] arrayOfMeetingSuggestions, string meetingSuggestion, DateTime startTime, DateTime endTime, string attendeeName, string attendeeEmail)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1504, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1504");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1504
                // If the array of MeetingSuggestionType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfMeetingSuggestions,
                    1504,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfMeetingSuggestionsType complex type which specifies an array of meeting suggestions. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1505, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1505");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1505
                // If the array element of MeetingSuggestionType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfMeetingSuggestions[0],
                    1505,
                    @"[In Appendix C: Product Behavior] Implementation does support MeetingSuggestionType complex type which specifies a meeting suggestion. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1083");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1083
                // If the array element of MeetingSuggestionType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfMeetingSuggestions[0],
                    1083,
                    @"[In t:ArrayOfMeetingSuggestionsType Complex Type] MeetingSuggestion: An element of type MeetingSuggestionType, as defined in section 2.2.4.25, that represents a single meeting suggestion.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1088");
                this.Site.Log.Add(
                    LogEntryKind.Debug,
                    string.Format(
                        "The extracted meeting subject should be a sentence extracted from the meeting string. The extracted meeting subject is: {0}. The meeting string is: {1}.",
                        arrayOfMeetingSuggestions[0].Subject,
                        meetingSuggestion));

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1088
                this.Site.CaptureRequirementIfIsTrue(
                    meetingSuggestion.Contains(arrayOfMeetingSuggestions[0].Subject),
                    1088,
                    @"[In t:MeetingSuggestionType Complex Type] Subject: An element of type string that represents the subject of the meeting.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1089");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1089
                this.Site.CaptureRequirementIfAreEqual<string>(
                    meetingSuggestion,
                    arrayOfMeetingSuggestions[0].MeetingString,
                    1089,
                    @"[In t:MeetingSuggestionType Complex Type] MeetingString: An element of type string that represents the name of the meeting.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1090");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1090
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    startTime,
                    arrayOfMeetingSuggestions[0].StartTime,
                    1090,
                    @"[In t:MeetingSuggestionType Complex Type] StartTime: An element of type dateTime, as defined in [XMLSCHEMA2] section 3.2.7, that represents the start time of the meeting.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1091");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1091
                this.Site.CaptureRequirementIfAreEqual<DateTime>(
                    endTime,
                    arrayOfMeetingSuggestions[0].EndTime,
                    1091,
                    @"[In t:MeetingSuggestionType Complex Type] EndTime: An element of type dateTime that represents the ending time of the meeting.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1086");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1086
                // If the Attendees element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfMeetingSuggestions[0].Attendees,
                    1086,
                    @"[In t:MeetingSuggestionType Complex Type] Attendees: An element of type ArrayOfEmailUsersType, as defined in section 2.2.4.32, that represents those invited to the meeting.");

                this.VerifyArrayOfEmailUsersType(arrayOfMeetingSuggestions[0].Attendees, attendeeName, attendeeEmail);
            }
        }

        /// <summary>
        /// Verify the ArrayOfContactsType structure.
        /// </summary>
        /// <param name="arrayOfContacts">An array of ContactType instance which is transformed from ArrayOfContactsType.</param>
        /// <param name="contactDisplayName">Display name of the contact extracted from the message.</param>
        /// <param name="businessName">Business name of the contact extracted from the message.</param>
        /// <param name="url">Url of the contact extracted from the message.</param>
        /// <param name="phoneNumber">Phone number of the contact extracted from the message.</param>
        /// <param name="phoneNumberType">Phone number type of the contact extracted from the message.</param>
        /// <param name="emailAddress">Email address of the contact extracted from the message.</param>
        /// <param name="address">Address of the contact extracted from the message.</param>
        protected void VerifyArrayOfContactsType(ContactType[] arrayOfContacts, string contactDisplayName, string businessName, Uri url, string phoneNumber, string phoneNumberType, string emailAddress, string address)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1507, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1507");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1507
                // If the array of ContactType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts,
                    1507,
                    @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfContactsType complex type which specifies an array of contacts. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1508, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1508");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1508
                // If the array of ContactType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts,
                    1508,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfContactsType complex type which specifies the type of a contact. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1097");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1097
                // If the array element of ContactType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts[0],
                    1097,
                    @"[In t:ArrayOfContactsType Complex Type] Contact: An element of type ContactType, as defined in section 2.2.4.28, that represents a single contact.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1100");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1100
                this.Site.CaptureRequirementIfAreEqual<string>(
                    contactDisplayName,
                    arrayOfContacts[0].PersonName,
                    1100,
                    @"[In t:ContactType Complex Type] PersonName: An element of type string, as defined in [XMLSCHEMA2] 3.2.1, that represents the name of a person.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1101");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1101
                this.Site.CaptureRequirementIfAreEqual<string>(
                    businessName,
                    arrayOfContacts[0].BusinessName,
                    1101,
                    @"[In t:ContactType Complex Type] BusinessName: An element of type string that represents the name of a business.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1102");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1102
                // If the PhoneNumbers element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts[0].PhoneNumbers,
                    1102,
                    @"[In t:ContactType Complex Type] PhoneNumbers: An element of type ArrayOfPhonesType, as defined in section 2.2.4.30, that represents phone number contacts.");

                this.VerifyArrayOfPhonesType(arrayOfContacts[0].PhoneNumbers, phoneNumber, phoneNumberType);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1103");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1103
                // If the Urls element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts[0].Urls,
                    1103,
                    @"[In t:ContactType Complex Type] Urls: An element of type ArrayOfUrlsType, as defined in section 2.2.4.29, that represents URL contacts.");

                this.VerifyArrayOfUrlsType(arrayOfContacts[0].Urls, url);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1104");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1104
                // If the EmailAddresses element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts[0].EmailAddresses,
                    1104,
                    @"[In t:ContactType Complex Type] EmailAddresses: An element of type ArrayOfExtractedEmailAddresses, as defined in section 2.2.4.35, that represents email contacts.");

                this.VerifyArrayOfExtractedEmailAddresses(arrayOfContacts[0].EmailAddresses, emailAddress);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1105");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1105
                // If the Addresses element is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfContacts[0].Addresses,
                    1105,
                    @"[In t:ContactType Complex Type] Addresses: An element of type ArrayOfAddressesType, as defined in section 2.2.4.23, that represents postal addresses of contacts.");

                this.VerifyArrayOfAddressesType(arrayOfContacts[0].Addresses, address);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1106");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1106
                this.Site.CaptureRequirementIfIsTrue(
                    arrayOfContacts[0].ContactString.Contains(contactDisplayName),
                    1106,
                    @"[In t:ContactType Complex Type] ContactString: An element of type string that represents the display name of a contact.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfPhoneEntitiesType structure.
        /// </summary>
        /// <param name="arrayOfPhoneEntities">An array of PhoneEntityType instance which is transformed from ArrayOfPhoneEntitiesType.</param>
        /// <param name="phoneNumber">Phone number extracted from the message.</param>
        /// <param name="phoneNumberType">Type of the phone number extracted from the message.</param>
        protected void VerifyArrayOfPhoneEntitiesType(PhoneEntityType[] arrayOfPhoneEntities, string phoneNumber, string phoneNumberType)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1720, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1720");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1720
                // If the array of PhoneEntityType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhoneEntities,
                    1720,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfPhoneEntitiesType complex type which specifies an array of phone entities. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1722, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1722");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1722
                // If the array element of PhoneEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhoneEntities[0],
                    1722,
                    @"[In Appendix C: Product Behavior] Implementation does support PhoneEntityType complex type which specifies a phone entity. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1764");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1764
                // If the array element of PhoneEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhoneEntities[0],
                    1764,
                    @"[In t:ArrayOfPhoneEntitiesType Complex Type] Phone: An element of type PhoneEntityType, as defined in section 2.2.4.44, that specifies a phone entity.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1768");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1768
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneNumber,
                    arrayOfPhoneEntities[0].OriginalPhoneString,
                    1768,
                    @"[In t:PhoneEntityType Complex Type] OriginalPhoneString: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents the original phone number.");

                // Get the phone string from the original phone string.
                string phoneString = string.Empty;
                for (int i = 0; i < phoneNumber.Length; i++)
                {
                    if (char.IsDigit(phoneNumber[i]))
                    {
                        phoneString += phoneNumber[i];
                    }
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1769");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1769
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneString,
                    arrayOfPhoneEntities[0].PhoneString,
                    1769,
                    @"[In t:PhoneEntityType Complex Type] PhoneString: An element of type string that represents the current phone number.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1770");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1770
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneNumberType,
                    arrayOfPhoneEntities[0].Type,
                    1770,
                    @"[In t:PhoneEntityType Complex Type] Type: An element of type string that represents the type of phone number, for example, ""Business"" or ""Home"".");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1140");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1140
                // This requirement can be captured directly after above requirements are captured.
                this.Site.CaptureRequirement(
                    1140,
                    @"[In t:EntityExtractionResultType Complex Type] PhoneNumbers: An element of type ArrayOfPhoneEntitiesType, as defined in section 2.2.4.30, that represents the phone numbers returned.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfUrlEntitiesType structure.
        /// </summary>
        /// <param name="arrayOfUrlEntities">An array of UrlEntityType instance which is transformed from ArrayOfUrlEntitiesType.</param>
        /// <param name="url">Url information extracted from the message.</param>
        protected void VerifyArrayOfUrlEntitiesType(UrlEntityType[] arrayOfUrlEntities, Uri url)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1724, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1724");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1724
                // If the array of UrlEntityType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfUrlEntities,
                    1724,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfUrlEntitiesType complex type which specifies an array of URL entities. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1726, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1726");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1726
                // If the array element of UrlEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfUrlEntities[0],
                    1726,
                    @"[In Appendix C: Product Behavior] Implementation does support UrlEntityType complex type which specifies a URL entity. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1773");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1773
                // If the array element of UrlEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfUrlEntities[0],
                    1773,
                    @"[In t:ArrayOfUrlEntitiesType Complex Type] UrlEntity: An element of type UrlEntityType, as defined in section 2.2.4.46, that specifies a URL entity.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1777");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1777
                this.Site.CaptureRequirementIfAreEqual<string>(
                    url.OriginalString,
                    arrayOfUrlEntities[0].Url,
                    1777,
                    @"[In t:UrlEntityType Complex Type] URL: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that specifies a URL.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1139");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1139
                this.Site.CaptureRequirementIfAreEqual<string>(
                    url.OriginalString,
                    arrayOfUrlEntities[0].Url,
                    1139,
                    @"[In t:EntityExtractionResultType Complex Type] Urls: An element of type ArrayOfUrlEntitiesType, as defined in section 2.2.4.29, that represents the URLs returned.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfEmailAddressEntitiesType structure.
        /// </summary>
        /// <param name="arrayOfEmailAddressEntities">An array of EmailAddressEntityType instance which is transformed from ArrayOfEmailAddressEntitiesType.</param>
        /// <param name="emailAddress">Email address extracted from the message.</param>
        protected void VerifyArrayOfEmailAddressEntitiesType(EmailAddressEntityType[] arrayOfEmailAddressEntities, string emailAddress)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1716, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1716");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1716
                // If the array of EmailAddressEntityType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailAddressEntities,
                    1716,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfAddressesType complex type which specifies an array of addresses. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1718, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1718");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1718
                // If the array element of EmailAddressEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailAddressEntities[0],
                    1718,
                    @"[In Appendix C: Product Behavior] Implementation does support EmailAddressEntityType complex type which specifies an email address entity. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1757");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1757
                // If the array element of EmailAddressEntityType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailAddressEntities[0],
                    1757,
                    @"[In t:ArrayOfEmailAddressEntitiesType Complex Type] EmailAddressEntity: An element of type EmailAddressEntityType, as defined in section 2.2.4.42, that represents an email address.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1761");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1761
                this.Site.CaptureRequirementIfAreEqual<string>(
                    emailAddress,
                    arrayOfEmailAddressEntities[0].EmailAddress,
                    1761,
                    @"[In t:EmailAddressEntityType Complex Type] EmailAddress: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that specifies an email address.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1137");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1137
                this.Site.CaptureRequirementIfAreEqual<string>(
                    emailAddress,
                    arrayOfEmailAddressEntities[0].EmailAddress,
                    1137,
                    @"[In t:EntityExtractionResultType Complex Type] EmailAddresses: An element of type ArrayOfEmailAddressEntitiesType, as defined in section 2.2.4.35, that represents the email addresses returned.");

                this.VerifyEntityType(arrayOfEmailAddressEntities[0]);
            }
        }

        /// <summary>
        /// Verify the ArrayOfTaskSuggestionsType structure.
        /// </summary>
        /// <param name="arrayOfTaskSuggestions">An array of TaskSuggestionType instance which is transformed from ArrayOfTaskSuggestionsType.</param>
        /// <param name="taskSuggestion">Task suggestion string extracted from the message.</param>
        /// <param name="assigneeName">Assignee name extracted from the message.</param>
        /// <param name="assigneeEmail">Assignee email extracted from the message.</param>
        protected void VerifyArrayOfTaskSuggestionsType(TaskSuggestionType[] arrayOfTaskSuggestions, string taskSuggestion, string assigneeName, string assigneeEmail)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1506, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1506");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1506
                // If the array of TaskSuggestionType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfTaskSuggestions,
                    1506,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfTaskSuggestionsType complex type which specifies an array of task suggestions. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1704, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1704");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1704
                // If the array element of TaskSuggestionType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfTaskSuggestions[0],
                    1704,
                    @"[In Appendix C: Product Behavior] Implementation does support TaskSuggestionType complex type which specifies a task suggestion. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1094");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1094
                // If the array element of TaskSuggestionType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfTaskSuggestions[0],
                    1094,
                    @"[In t:ArrayOfTaskSuggestionsType Complex Type] TaskSuggestion: An element of type TaskSuggestionType, as defined in section 2.2.4.34, that represents a single task suggestion.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1127");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1127
                this.Site.CaptureRequirementIfAreEqual<string>(
                    taskSuggestion,
                    arrayOfTaskSuggestions[0].TaskString,
                    1127,
                    @"[In t:TaskSuggestionType Complex Type] TaskString: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents the name of the task.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1128");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1128
                // If the Assignees element of TaskSuggestionType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfTaskSuggestions[0].Assignees,
                    1128,
                    @"[In t:TaskSuggestionType Complex Type] Assignees: An element of type ArrayOfEmailUsersType, as defined in section 2.2.4.31, that represents the persons that are to accomplish the task.");

                this.VerifyArrayOfEmailUsersType(arrayOfTaskSuggestions[0].Assignees, assigneeName, assigneeEmail);
            }
        }

        /// <summary>
        /// Verify the EntityType structure.
        /// </summary>
        /// <param name="entity">An EntityType instance.</param>
        protected void VerifyEntityType(EntityType entity)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1710, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1710");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1701
                // If the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirement(
                    1710,
                    @"[In Appendix C: Product Behavior] Implementation does support EntityType complex type which specifies a single entity. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1747");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1747
                // If the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirement(
                    1747,
                    @"[In t:EntityType Complex Type] Position: An element of type EmailPositionType, as defined in section 2.2.5.2, that specifies where the entity was found in the message.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1782");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1782
                this.Site.CaptureRequirementIfAreEqual<EmailPositionType>(
                    EmailPositionType.LatestReply,
                    entity.Position[0],
                    1782,
                    @"[In t:EmailPositionType Simple Type] [The value ""LatestReply""] Specifies the latest reply.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfExtractedEmailAddresses structure.
        /// </summary>
        /// <param name="arrayOfExtractedEmailAddresses">A string array which is transformed from ArrayOfExtractedEmailAddresses.</param>
        /// <param name="emailAddress">Email address information extracted from the message.</param>
        protected void VerifyArrayOfExtractedEmailAddresses(string[] arrayOfExtractedEmailAddresses, string emailAddress)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1706, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1706");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1706
                // If the arrayOfExtractedEmailAddresses is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfExtractedEmailAddresses,
                    1706,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfExtractedEmailAddresses complex type which specifies an array of email addresses. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1131");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1131
                this.Site.CaptureRequirementIfAreEqual<string>(
                    emailAddress,
                    arrayOfExtractedEmailAddresses[0],
                    1131,
                    @"[In t:ArrayOfExtractedEmailAddresses Complex Type] EmailAddress: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents a single email address.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfPhonesType structure.
        /// </summary>
        /// <param name="arrayOfPhones">An array of PhoneType instance which is transformed from ArrayOfPhonesType.</param>
        /// <param name="phoneNumber">Phone number extracted from the message.</param>
        /// <param name="phoneNumberType">Type of the phone number extracted from the message.</param>
        protected void VerifyArrayOfPhonesType(PhoneType[] arrayOfPhones, string phoneNumber, string phoneNumberType)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1510, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1510");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1510
                // If the array of PhoneType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhones,
                    1510,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfPhonesType complex type which specifies an array of phone numbers. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1511, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1511");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1511
                // If the array element of PhoneType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhones[0],
                    1511,
                    @"[In Appendix C: Product Behavior] Implementation does support PhoneType complex type which specifies a phone number and its type. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1112");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1112
                // If the array element of PhoneType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfPhones[0],
                    1112,
                    @"[In t:ArrayOfPhonesType Complex Type] Phone: An element of type PhoneType, as defined in section 2.2.4.31, that represents a single phone number.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1115");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1115
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneNumber,
                    arrayOfPhones[0].OriginalPhoneString,
                    1115,
                    @"[In t:PhoneType Complex Type] OriginalPhoneString: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents the original phone number.");

                // Get the phone string from the original phone string.
                string phoneString = string.Empty;
                for (int i = 0; i < phoneNumber.Length; i++)
                {
                    if (char.IsDigit(phoneNumber[i]))
                    {
                        phoneString += phoneNumber[i];
                    }
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1116");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1116
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneString,
                    arrayOfPhones[0].PhoneString,
                    1116,
                    @"[In t:PhoneType Complex Type] PhoneString: An element of type string that represents the current phone number.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1117");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1117
                this.Site.CaptureRequirementIfAreEqual<string>(
                    phoneNumberType,
                    arrayOfPhones[0].Type,
                    1117,
                    @"[In t:PhoneType Complex Type] Type: An element of type string that represents the type of phone, for example, ""Business"" or ""Home"".");
            }
        }

        /// <summary>
        /// Verify the ArrayOfEmailUsersType structure.
        /// </summary>
        /// <param name="arrayOfEmailUsers">An array of EmailUserType instance which is transformed from ArrayOfEmailUsersType.</param>
        /// <param name="userName">User name of the mailbox extracted from the message.</param>
        /// <param name="userEmail">Email address of the mailbox extracted from the message.</param>
        protected void VerifyArrayOfEmailUsersType(EmailUserType[] arrayOfEmailUsers, string userName, string userEmail)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1512, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1512");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1512
                // If the array of EmailUserType instance is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailUsers,
                    1512,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfEmailUsersType complex type which specifies an array of email users. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1702, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1702");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1702
                // If the array element of EmailUserType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailUsers[0],
                    1702,
                    @"[In Appendix C: Product Behavior] Implementation does support EmailUserType complex type which specifies an email user. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1120");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1120
                // If the array element of EmailUserType is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfEmailUsers[0],
                    1120,
                    @"[In t:ArrayOfEmailUsersType Complex Type] EmailUser: An element of type EmailUserType, as defined in section 2.2.4.33, that represents a single email user.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1123");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1123
                this.Site.CaptureRequirementIfAreEqual<string>(
                    userName.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                    arrayOfEmailUsers[0].Name.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                    1123,
                    @"[In t:EmailUserType Complex Type] Name: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents the name of the email user.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1124");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1124
                this.Site.CaptureRequirementIfAreEqual<string>(
                    userEmail.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                    arrayOfEmailUsers[0].UserId.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)),
                    1124,
                    @"[In t:EmailUserType Complex Type] UserId: An element of type string that represents the user identifier of the email user.");
            }
        }

        /// <summary>
        /// Verify the array of url structure.
        /// </summary>
        /// <param name="arrayOfUrls">A string array.</param>
        /// <param name="url">Url information extracted from the message.</param>
        protected void VerifyArrayOfUrlsType(string[] arrayOfUrls, Uri url)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1509, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1509");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1509
                // If the arrayOfUrls is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfUrls,
                    1509,
                    @"[In Appendix C: Product Behavior] Implementation does support ArrayOfUrlsType complex type which specifies an array of URLs. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1109");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1109
                this.Site.CaptureRequirementIfAreEqual<string>(
                    url.OriginalString,
                    arrayOfUrls[0],
                    1109,
                    @"[In t:ArrayOfUrlsType Complex Type] Url: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that specifies a single URL.");
            }
        }

        /// <summary>
        /// Verify the ArrayOfAddressesType structure.
        /// </summary>
        /// <param name="arrayOfAddresses">A string array which is transformed from ArrayOfAddressesType.</param>
        /// <param name="address">Address information extracted from the message.</param>
        protected void VerifyArrayOfAddressesType(string[] arrayOfAddresses, string address)
        {
            Site.Assert.IsTrue(this.IsSchemaValidated, "The schema should be validated.");

            if (Common.IsRequirementEnabled(1503, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1503");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1503
                // If the arrayOfAddresses is not null and the schema could be validated successfully, this requirement can be verified.
                this.Site.CaptureRequirementIfIsNotNull(
                    arrayOfAddresses,
                    1503,
                    @"[In Appendix C: Product Behavior] Implementation does support the ArrayOfAddressesType complex type which specifies an array of addresses. (Exchange 2013 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1080");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1080
                this.Site.CaptureRequirementIfAreEqual<string>(
                    address,
                    arrayOfAddresses[0],
                    1080,
                    @"[In t:ArrayOfAddressesType Complex Type] Address: An element of type string, as defined in [XMLSCHEMA2] section 3.2.1, that represents a single address.");
            }
        }

        #region Preview Structure
        /// <summary>
        /// Verify the Preview structure
        /// </summary>
        /// <param name="preview">An Preview string.</param>
        /// <param name="bodyContent">The content of body element.</param>
        protected void VerifyPreview(string preview, string bodyContent)
        {
            if (Common.IsRequirementEnabled(1354, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1354. Expected length: no longer than 256, actual Length: {0}", preview.Length);

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1354
                // If schema is validated and the length of preview is no longer than 256 characters,
                // and the preview element equals to the first 256 characters of the body,
                // this requirement can be validated.
                bool isVerifyR1354 = preview.Length <= 256 && string.Equals(preview, bodyContent.Length > 256 ? bodyContent.Substring(0, 256) : bodyContent);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1354,
                    1354,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""Preview"" with type ""xs:string"" which specifies the first 256 characters of the body of a message for preview without opening the message. (Exchange 2013 and above follow this behavior.)");
            }
        }
        #endregion

        #region TextBody Structure
        /// <summary>
        /// Verify the TextBody structure
        /// </summary>
        /// <param name="textBody">An BodyType instance of textBody.</param>
        protected void VerifyTextBody(BodyType textBody)
        {
            if (Common.IsRequirementEnabled(1731, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1731");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1731
                // if the element is not null, and the schema is validated,
                // this requirement can be validated.
                Site.CaptureRequirementIfIsNotNull(
                    textBody,
                    1731,
                    @"[In Appendix C: Product Behavior] Implementation does support element name ""TextBody"" with type ""t:BodyType"" which specifies the text body of the item. (Exchange 2013 and above follow this behavior.)");
            }
        }
        #endregion

        #region InstanceKey Structure
        /// <summary>
        /// Verify the InstanceKey structure
        /// </summary>
        /// <param name="instanceKey">A byte array of InstanceKey.</param>
        protected void VerifyInstanceKey(byte[] instanceKey)
        {
            if (Common.IsRequirementEnabled(1348, this.Site))
            {
                Site.Assert.IsNotNull(instanceKey, "The InstanceKey element should not be null.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXWSCORE_R1348");

                // Verify MS-OXWSCORE requirement: MS-OXWSCORE_R1348
                // if the element is not null, the schema is validated, and the instanceKey is base64 binary data
                // this requirement can be validated.
                Site.CaptureRequirementIfIsTrue(
                    TestSuiteHelper.IsBase64Binary(instanceKey),
                    1348,
                    @"[In Appendix C: Product Behavior] Implementation does support element ""InstanceKey"" with type ""xs:base64Binary [XMLSCHEMA2]"" which specifies the key for an instance. (Exchange 2013 and above follow this behavior.)");
            }
        }
        #endregion

        #endregion

        #endregion

        #region Private methods
        /// <summary>
        ///  Check an item id exists or not.
        /// </summary>
        /// <param name="itemId">The item id need to check</param>
        /// <returns>A boolean value indicates the id exist or not.</returns>
        private bool IsIdExisted(ItemIdType itemId)
        {
            bool isIdExisted = false;
            for (int count = 0; count < this.ExistItemIds.Count; count++)
            {
                if (itemId.Id.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)) == this.ExistItemIds[count].Id.ToUpper(new CultureInfo(TestSuiteHelper.Culture, false)))
                {
                    return isIdExisted = true;
                }
            }

            return isIdExisted;
        }
        #endregion
    }
   
}
