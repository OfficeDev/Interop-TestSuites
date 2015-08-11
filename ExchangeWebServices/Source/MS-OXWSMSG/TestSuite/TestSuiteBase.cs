namespace Microsoft.Protocols.TestSuites.MS_OXWSMSG
{
    using System;
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
        /// Gets the SUT Control adapter which provides FindItem method.
        /// </summary>
        protected IMS_OXWSSRCHSUTControlAdapter SRCHSUTControlAdapter { get; private set; }

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
            this.SRCHSUTControlAdapter = Site.GetAdapter<IMS_OXWSSRCHSUTControlAdapter>();
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

            return this.SRCHSUTControlAdapter.FindItem(userName, password, domain, folder, subject, field);
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
        #endregion
    }
}