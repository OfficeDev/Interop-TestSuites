namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The bass class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Fields
        /// <summary>
        /// The subject of message.
        /// </summary>
        private string subject;

        /// <summary>
        /// The domain information
        /// </summary>
        private string domain;

        /// <summary>
        /// The mail sender.
        /// </summary>
        private string sender;

        /// <summary>
        /// The first mail recipient.
        /// </summary>
        private string recipient1;

        /// <summary>
        /// The second mail recipient.
        /// </summary>
        private string recipient2;

        /// <summary>
        /// The meeting room's address.
        /// </summary>
        private string meetingRoom;

        /// <summary>
        /// The reference of message.
        /// </summary>
        private string reference;

        /// <summary>
        /// The wait time for meeting request, meeting response or meeting cancellation message to be received.
        /// </summary>
        private int waitTime;

        /// <summary>
        /// The upper bound of a loop for waiting for a valid Server response.
        /// </summary>
        private int upperBound;

        /// <summary>
        /// The related Items of ItemInfoResponseMessageType returned from server.
        /// </summary>
        private ItemInfoResponseMessageType[] infoItems;

        /// <summary>
        /// The first Item of the first responseMessageItem in infoItems returned from server response.
        /// </summary>
        private ItemType firstItemOfFirstInfoItem;
        #endregion

        #region Properties
        /// <summary>
        /// Gets the MS-OXWSMSG protocol adapter.
        /// </summary>
        protected IMS_OXWSMSGAdapter MSGAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSSRCH protocol adapter.
        /// </summary>
        protected IMS_OXWSSRCHAdapter SRCHAdapter { get; private set; }

        /// <summary>
        /// Gets the SUT Control adapter which provides Cleanup method.
        /// </summary>
        protected IMS_OXWSMSGSUTControlAdapter MSGSUTControlAdapter { get; private set; }

        /// <summary>
        /// Gets the subject of message.
        /// </summary>
        protected string Subject
        {
            get { return this.subject; }
        }

        /// <summary>
        /// Gets the domain information
        /// </summary>
        protected string Domain
        {
            get { return this.domain; }
        }

        /// <summary>
        /// Gets the first contact email address.
        /// </summary>
        protected string Sender
        {
            get { return this.sender; }
        }

        /// <summary>
        /// Gets the second contact email address.
        /// </summary>
        protected string Recipient1
        {
            get { return this.recipient1; }
        }

        /// <summary>
        /// Gets the third contact email address.
        /// </summary>
        protected string Recipient2
        {
            get { return this.recipient2; }
        }

        /// <summary>
        /// Gets the fourth contact email address.
        /// </summary>
        protected string MeetingRoom
        {
            get { return this.meetingRoom; }
        }

        /// <summary>
        /// Gets the reference of message.
        /// </summary>
        protected string MsgReference
        {
            get { return this.reference; }
        }

        /// <summary>
        /// Gets the wait time for meeting request, meeting response or meeting cancellation message to be received.
        /// </summary>
        protected int WaitTime
        {
            get { return this.waitTime; }
        }

        /// <summary>
        /// Gets the upper bound of a loop for waiting for a valid Server response.
        /// </summary>
        protected int UpperBound
        {
            get { return this.upperBound; }
        }
        #endregion

        #region Test case initialize
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();

            this.MSGAdapter = Site.GetAdapter<IMS_OXWSMSGAdapter>();
            this.SRCHAdapter = Site.GetAdapter<IMS_OXWSSRCHAdapter>();
            this.MSGSUTControlAdapter = Site.GetAdapter<IMS_OXWSMSGSUTControlAdapter>();
            this.subject = Common.GenerateResourceName(this.Site, Common.GetConfigurationPropertyValue("Subject", this.Site));

            // The sender is used to connect with Exchange service and as the sender of message.
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.sender = Common.GetConfigurationPropertyValue("Sender", this.Site) + "@" + this.domain;
            this.recipient1 = Common.GetConfigurationPropertyValue("Recipient1", this.Site) + "@" + this.domain;
            this.recipient2 = Common.GetConfigurationPropertyValue("Recipient2", this.Site) + "@" + this.domain;
            this.meetingRoom = Common.GetConfigurationPropertyValue("RoomName", this.Site) + "@" + this.domain;
            this.reference = Common.GetConfigurationPropertyValue("Reference", this.Site);
            this.waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));

            if (!int.TryParse(Common.GetConfigurationPropertyValue("RetryCount", this.Site), out this.upperBound))
            {
                Site.Assert.Fail("The value of RequestUpperBound property in ptf file could not be converted to an integer value. The current value is {0}.", Common.GetConfigurationPropertyValue("RequestUpperBound", this.Site));
            }
        }
        #endregion

        #region Test case base methods
        #region Define general GetItem request message
        /// <summary>
        /// Define general GetItem request message
        /// </summary>
        /// <param name="itemId">The item identifier of the item.</param>
        /// <param name="baseShape">The basic configuration of properties to be returned in an item response.</param>
        /// <returns>A request to get an item from a mailbox</returns>
        protected GetItemType DefineGeneralGetItemRequestMessage(ItemIdType itemId, DefaultShapeNamesType baseShape)
        {
            GetItemType getItemRequest = new GetItemType
            {
                ItemIds = new ItemIdType[]
                {
                    itemId
                },

                ItemShape = new ItemResponseShapeType
                {
                    BaseShape = baseShape,
                }
            };

            return getItemRequest;
        }
        #endregion

        #region Define general CreateItem request message
        /// <summary>
        /// Define general CreateItem request message
        /// </summary>
        /// <param name="messageDisposition">A value specifies how a message item is handled after it is created or updated</param>
        /// <param name="savedTargetFolderName">The folder in which new items are saved</param>
        /// <returns>A request to create an item</returns>
        protected CreateItemType GetCreateItemType(MessageDispositionType messageDisposition, DistinguishedFolderIdNameType savedTargetFolderName)
        {
            CreateItemType createItemRequest = new CreateItemType
            {
                MessageDisposition = messageDisposition,
                MessageDispositionSpecified = true,

                SavedItemFolderId = new TargetFolderIdType
                {
                    Item = new DistinguishedFolderIdType
                    {
                        Id = savedTargetFolderName
                    }
                },

                Items = new NonEmptyArrayOfAllItemsType
                {
                    Items = new MessageType[]
                    {
                        new MessageType
                        {
                            Sender = new SingleRecipientType
                            {
                                Item = new EmailAddressType
                                {
                                    EmailAddress = this.sender
                                }                                
                            },

                            ToRecipients = new EmailAddressType[]
                            {
                                new EmailAddressType
                                {
                                        EmailAddress = this.recipient1                             
                                }
                            },

                            Subject = this.subject,  
                        }
                    }
                },
            };

            return createItemRequest;
        }
        #endregion

        #region Verify the validation of a response returned by server
        /// <summary>
        /// Verify the validation of a response returned by server.
        /// </summary>
        /// <param name="response">A response returned by server.</param>
        /// <returns>The validation result.</returns>
        protected bool VerifyResponse(BaseResponseMessageType response)
        {
            bool isValidResponse = false;
            if (response != null
                && response.ResponseMessages != null
                && response.ResponseMessages.Items != null
                && response.ResponseMessages.Items.Length > 0
                && response.ResponseMessages.Items[0] != null)
            {
                isValidResponse = string.Equals(response.ResponseMessages.Items[0].ResponseClass.ToString(), "Success", StringComparison.CurrentCultureIgnoreCase);
                if (!isValidResponse && response.ResponseMessages.Items[0].ResponseCodeSpecified)
                {
                    Site.Assert.Fail("The operation failed with response code '{0}'", response.ResponseMessages.Items[0].ResponseCode.ToString());
                }
            }

            return isValidResponse;
        }

        /// <summary>
        /// Verify the validation of a response of CreateItem operation returned by server.
        /// </summary>
        /// <param name="createItemResponse">A response of CreateItem operation returned by server.</param>
        /// <param name="messageDisposition">A value specifies how a message item is handled after it is created or updated.</param>
        /// <returns>The validation result.</returns>
        protected bool VerifyCreateItemResponse(CreateItemResponseType createItemResponse, MessageDispositionType messageDisposition)
        {
            bool isValidCreateResponse = false;
            if (this.VerifyResponse(createItemResponse))
            {
                this.infoItems = TestSuiteHelper.GetInfoItemsInResponse(createItemResponse);
                Site.Assert.IsNotNull(this.infoItems, @"The CreateItem response should contain one or more items of ItemInfoResponseMessageType.");
                switch (messageDisposition)
                {
                    case MessageDispositionType.SaveOnly:
                        this.firstItemOfFirstInfoItem = TestSuiteHelper.GetItemTypeItemFromInfoItemsByIndex(this.infoItems, 0, 0);
                        Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem, @"The first item of the array of ItemType type returned from server response should not be null.");
                        Site.Assert.IsNotNull(this.firstItemOfFirstInfoItem.ItemId, @"The ItemId property of the first item should not be null.");
                        isValidCreateResponse = true;
                        break;
                    case MessageDispositionType.SendAndSaveCopy:
                    case MessageDispositionType.SendOnly:
                        Site.Assert.IsTrue(this.infoItems.Length > 0, "infoItems instance should contain at least one item.");
                        Site.Assert.IsNotNull(this.infoItems[0], "The first item of infoItems instance should not be null.");
                        Site.Assert.IsNotNull(this.infoItems[0].Items, "The Items property of the first item of infoItems instance should not be null.");
                        break;
                }

                isValidCreateResponse = true;
            }

            return isValidCreateResponse;
        }
        #endregion

        #region Search the items with the specified subject in the specified folder.
        /// <summary>
        /// Search the items with the specified subject in the specified folder.
        /// </summary>
        /// <param name="role">The role used to communicate with server</param>
        /// <param name="folder">A string that specifies the folder to search.</param>
        /// <param name="subject">The subject of the item to search</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If one or more items are found, return true; otherwise, false.</returns>
        protected bool SearchItems(Role role, string folder, string subject, string field)
        {
            string userName = string.Empty;
            string password = string.Empty;
            string domain = string.Empty;

            switch (role)
            {
                case Role.Sender:
                    userName = Common.GetConfigurationPropertyValue("Sender", this.Site);
                    password = Common.GetConfigurationPropertyValue("SenderPassword", this.Site);
                    break;
                case Role.Recipient1:
                    userName = Common.GetConfigurationPropertyValue("Recipient1", this.Site);
                    password = Common.GetConfigurationPropertyValue("Recipient1Password", this.Site);
                    break;
                case Role.Recipient2:
                    userName = Common.GetConfigurationPropertyValue("Recipient2", this.Site);
                    password = Common.GetConfigurationPropertyValue("Recipient2Password", this.Site);
                    break;
            }

            domain = Common.GetConfigurationPropertyValue("Domain", this.Site);

            return this.FindItem(userName, password, domain, folder, subject, field);
        }
        #endregion

        #region Verify the response returned by server when the operating multiple items
        /// <summary>
        /// Verify the response returned by server when the operating multiple items.
        /// </summary>
        /// <param name="response">A response returned by server.</param>
        /// <returns>The validation result.</returns>
        protected bool VerifyMultipleResponse(BaseResponseMessageType response)
        {
            bool isValidResponse = false;
            if (response != null
                && response.ResponseMessages != null
                && response.ResponseMessages.Items != null
                && response.ResponseMessages.Items.Length > 0)
            {
                Site.Assert.AreEqual<int>(2, response.ResponseMessages.Items.Length, "There should be two Items elements in the operation response because this case only design for operate two items");
                bool isValidFirstItemsofResponse = string.Equals(response.ResponseMessages.Items[0].ResponseClass.ToString(), "Success", System.StringComparison.CurrentCultureIgnoreCase);
                if (!isValidFirstItemsofResponse)
                {
                    Site.Assert.Fail("The operation failed with response code '{0}'", response.ResponseMessages.Items[0].ResponseCode.ToString());
                }
                
                bool isValidSecondItemsofResponse = string.Equals(response.ResponseMessages.Items[1].ResponseClass.ToString(), "Success", System.StringComparison.CurrentCultureIgnoreCase);
                if (!isValidSecondItemsofResponse)
                {
                    Site.Assert.Fail("The operation failed with response code '{0}'", response.ResponseMessages.Items[1].ResponseCode.ToString());
                }

                isValidResponse = isValidFirstItemsofResponse && isValidSecondItemsofResponse;
            }

            return isValidResponse;
        }
        #endregion

         /// <summary>
        /// The operation searches the mailbox and returns the result whether one or more valid items are found.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the operation succeeds, return true; otherwise, return false.</returns>
        protected bool FindItem(string userName, string password, string domain, string folderName, string value, string field)
        {
            ItemType specificItem = this.FindSpecificItem(userName, password, domain, folderName, value, field);
            if (specificItem != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// The operation searches the specified user's mailbox and returns the result whether the valid items are found after the MoveItem or DeleteItem operation completed.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="password">Password of the user.</param>
        /// <param name="domain">Domain of the user.</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the item existed in the specific folder, return true; otherwise, return false.</returns>
        protected bool IsItemAvailableAfterMoveOrDelete(string userName, string password, string domain, string folderName, string value, string field)
        {
            this.SRCHAdapter.SwitchUser(userName, password, domain);

            // Construct a request for FindItem operation.
            FindItemType findRequest = this.ConstructFindItemRequest(folderName, value, field);

            #region Invoke FindItem operation
            int counter = 0;
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            FindItemResponseType findResponse = new FindItemResponseType();

            while (counter < upperBound)
            {
                Thread.Sleep(waitTime);

                findResponse = this.SRCHAdapter.FindItem(findRequest);

                if (findResponse != null
                    && findResponse.ResponseMessages != null
                    && findResponse.ResponseMessages.Items != null
                    && findResponse.ResponseMessages.Items.Length > 0)
                {
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        if (item.ResponseClass == ResponseClassType.Success)
                        {
                            FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                            if (findItem.RootFolder.Item != null)
                            {
                                ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                                if (realItems.Items != null && realItems.Items.Length > 0)
                                {
                                    counter++;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
            }

            Site.Log.Add(LogEntryKind.Debug, "Even after retrying {0} times, Message with specified subject is still available, this means previous call to MoveItem or DeleteItem has not worked yet.", counter);
            return true;
            #endregion
        }

        /// <summary>
        /// The operation searches the mailbox and returns the result whether one or more valid items are found.
        /// </summary>
        /// <param name="userName">The userName of the user used to communicate with server</param>
        /// <param name="password">The password of the user used to communicate with server</param>
        /// <param name="domain">The domain of the user used to communicate with server</param>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>If the operation succeeds, return the specific item; otherwise, return null.</returns>
        protected ItemType FindSpecificItem(string userName, string password, string domain, string folderName, string value, string field)
        {
            this.SRCHAdapter.SwitchUser(userName, password, domain);

            // Construct a request for FindItem operation.
            FindItemType findRequest = this.ConstructFindItemRequest(folderName, value, field);

            #region Invoke FindItem operation
            int counter = 0;
            int upperBound = int.Parse(Common.GetConfigurationPropertyValue("RetryCount", this.Site));
            int waitTime = int.Parse(Common.GetConfigurationPropertyValue("WaitTime", this.Site));
            FindItemResponseType findResponse = new FindItemResponseType();

            while (counter < upperBound)
            {
                Thread.Sleep(waitTime);

                findResponse = this.SRCHAdapter.FindItem(findRequest);

                if (findResponse != null
                    && findResponse.ResponseMessages != null
                    && findResponse.ResponseMessages.Items != null
                    && findResponse.ResponseMessages.Items.Length > 0)
                {
                    foreach (ResponseMessageType item in findResponse.ResponseMessages.Items)
                    {
                        if (item.ResponseClass == ResponseClassType.Success)
                        {
                            FindItemResponseMessageType findItem = item as FindItemResponseMessageType;
                            if (findItem.RootFolder.Item != null)
                            {
                                ArrayOfRealItemsType realItems = findItem.RootFolder.Item as ArrayOfRealItemsType;
                                if (realItems.Items != null && realItems.Items.Length > 0)
                                {
                                    return realItems.Items[0];
                                }
                            }
                        }
                    }
                }

                counter++;
            }

            Site.Log.Add(LogEntryKind.Debug, "Even after retrying {0} times, operation FindItem could not find any message.", counter);
            return null;
            #endregion
        }

        #endregion

        #region Private Methods
        /// <summary>
        /// The operation construct a request for FindItem operation.
        /// </summary>
        /// <param name="folderName">A string that specifies the folder to search.</param>
        /// <param name="value">A string that specifies the value for a search restriction.</param>
        /// <param name="field">A string that specifies the type of referenced field URI.</param>
        /// <returns>The request of FindItem operation which constructed with special folder name, search restriction and referenced field URI</returns>
        protected FindItemType ConstructFindItemRequest(string folderName, string value, string field)
        {
            FindItemType findRequest = new FindItemType();
            findRequest.ItemShape = new ItemResponseShapeType();
            findRequest.ItemShape.BaseShape = DefaultShapeNamesType.AllProperties;

            DistinguishedFolderIdType folderId = new DistinguishedFolderIdType();
            DistinguishedFolderIdNameType folderIdName;
            if (Enum.TryParse<DistinguishedFolderIdNameType>(folderName, out folderIdName))
            {
                folderId.Id = folderIdName;
            }
            else
            {
                Site.Assert.Fail("The value of the first argument (foldIdNameType) of FindItem operation is invalid.");
            }

            findRequest.ParentFolderIds = new BaseFolderIdType[1];
            findRequest.ParentFolderIds[0] = folderId;

            PathToUnindexedFieldType itemClass = new PathToUnindexedFieldType();
            UnindexedFieldURIType fieldURI;
            if (Enum.TryParse<UnindexedFieldURIType>(field, out fieldURI))
            {
                // set search field.
                itemClass.FieldURI = fieldURI;
            }
            else
            {
                Site.Assert.Fail("The value of the second argument (fieldURIType) of FindItem operation is invalid.");
            }

            ContainsExpressionType expressionType = new ContainsExpressionType();
            expressionType.Item = itemClass;
            expressionType.ContainmentMode = ContainmentModeType.Substring;
            expressionType.ContainmentModeSpecified = true;
            expressionType.ContainmentComparison = ContainmentComparisonType.IgnoreCaseAndNonSpacingCharacters;
            expressionType.ContainmentComparisonSpecified = true;
            expressionType.Constant = new ConstantValueType();
            expressionType.Constant.Value = value;

            RestrictionType restriction = new RestrictionType();
            restriction.Item = expressionType;

            if (!string.IsNullOrEmpty(value))
            {
                findRequest.Restriction = restriction;
            }

            return findRequest;
        }
        #endregion
    }
}