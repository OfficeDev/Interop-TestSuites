namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.Security.Policy;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// The base test suite class defines common initialization method and cleanup method for all the five scenarios.
    /// </summary>
    [TestClass]
    public class TestSuiteBase : TestClassBase
    {
        #region Private variables

        /// <summary>
        /// Folder handle of out box.
        /// </summary>
        private static uint outBoxFolderHandle;

        /// <summary>
        /// Folder id of out box
        /// </summary>
        private static ulong outBoxFolderID;

        /// <summary>
        /// Folder handle of sent items folder.
        /// </summary>
        private static uint sentItemsFolderHandle;

        /// <summary>
        /// Folder id of sent items folder.
        /// </summary>
        private static ulong sentItemsFolderID;

        /// <summary>
        /// Inbox folder handle.
        /// </summary>
        private uint inboxFolderHandle;

        /// <summary>
        /// DAF folder handle.
        /// </summary>
        private uint dafFolderHandle;

        /// <summary>
        /// Logon handle.
        /// </summary>
        private uint logonHandle;

        /// <summary>
        /// Folder id of inbox folder.
        /// </summary>
        private ulong inboxFolderID;

        /// <summary>
        /// Folder id of daf folder.
        /// </summary>
        private ulong dafFolderID;

        /// <summary>
        /// The value identifies the mailbox on which the logon is performed
        /// </summary>
        private byte[] mailboxGUID;

        /// <summary>
        /// The instance of the server adapter.
        /// </summary>
        private IMS_OXORULEAdapter oxoruleAdapter;

        /// <summary>
        /// The instance of the PowerShell adapter.
        /// </summary>
        private IMS_OXORULESUTControlAdapter sutAdapter;

        /// <summary>
        /// The instance of the set OOF rule adapter.
        /// </summary>
        private IMS_OXWOOFSUTControlAdapter sutSetOOFAdapter;

        /// <summary>
        /// String server name.
        /// </summary>
        private string server;

        /// <summary>
        /// String user1 ESSDN.
        /// </summary>
        private string user1ESSDN;

        /// <summary>
        /// String user2 ESSDN.
        /// </summary>
        private string user2ESSDN;

        /// <summary>
        /// String user2 name.
        /// </summary>
        private string user2Name;

        /// <summary>
        /// String domain name.
        /// </summary>
        private string domain;

        /// <summary>
        /// String user1 name.
        /// </summary>
        private string user1Name;

        /// <summary>
        /// String user1 password.
        /// </summary>
        private string user1Password;

        /// <summary>
        /// String user2 password.
        /// </summary>
        private string user2Password;

        /// <summary>
        /// The time of waiting for the rule to take effect.
        /// </summary>
        private int waitForTheRuleToTakeEffect;

        /// <summary>
        /// The value specifies the maximum Size allowed for a property value returned.
        /// </summary>
        private ushort propertySizeLimitFlag;

        /// <summary>
        /// The time of waiting for setting OFF complete.
        /// </summary>
        private int waitForSetOOFComplete;

        /// <summary>
        /// The maximum number of retry times if the client cannot retrieve the expected message in the preconfigured time period for some unknown reasons.
        /// </summary>
        private ushort getMessageRepeatTime;

        /// <summary>
        /// Identify test user account.
        /// </summary>
        private TestUser testUser;
        #endregion

        #region Protected Properties

        /// <summary>
        /// Gets or sets a value for Inbox folder handle.
        /// </summary>
        protected uint InboxFolderHandle
        {
            get { return this.inboxFolderHandle; }
            set { this.inboxFolderHandle = value; }
        }

        /// <summary>
        /// Gets or sets a value for DAF folder handle.
        /// </summary>
        protected uint DAFFolderHandle
        {
            get { return this.dafFolderHandle; }
            set { this.dafFolderHandle = value; }
        }

        /// <summary>
        /// Gets or sets a value for Logon handle.
        /// </summary>
        protected uint LogonHandle
        {
            get { return this.logonHandle; }
            set { this.logonHandle = value; }
        }

        /// <summary>
        /// Gets or sets a value for folder id of inbox folder.
        /// </summary>
        protected ulong InboxFolderID
        {
            get { return this.inboxFolderID; }
            set { this.inboxFolderID = value; }
        }

        /// <summary>
        /// Gets or sets a value for folder id of daf folder.
        /// </summary>
        protected ulong DAFFolderID
        {
            get { return this.dafFolderID; }
            set { this.dafFolderID = value; }
        }

        /// <summary>
        /// Gets or sets a value which identifies the mailbox on which the logon is performed
        /// </summary>
        protected byte[] MailboxGUID
        {
            get { return this.mailboxGUID; }
            set { this.mailboxGUID = value; }
        }

        /// <summary>
        /// Gets or sets an instance of the server adapter.
        /// </summary>
        protected IMS_OXORULEAdapter OxoruleAdapter
        {
            get { return this.oxoruleAdapter; }
            set { this.oxoruleAdapter = value; }
        }

        /// <summary>
        /// Gets or sets an instance of the PowerShell adapter.
        /// </summary>
        protected IMS_OXORULESUTControlAdapter SUTAdapter
        {
            get { return this.sutAdapter; }
            set { this.sutAdapter = value; }
        }

        /// <summary>
        /// Gets or sets an instance of the set OOF rule adapter.
        /// </summary>
        protected IMS_OXWOOFSUTControlAdapter SUTSetOOFAdapter
        {
            get { return this.sutSetOOFAdapter; }
            set { this.sutSetOOFAdapter = value; }
        }

        /// <summary>
        /// Gets or sets a value for the server name.
        /// </summary>
        protected string Server
        {
            get { return this.server; }
            set { this.server = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user1 ESSDN.
        /// </summary>
        protected string User1ESSDN
        {
            get { return this.user1ESSDN; }
            set { this.user1ESSDN = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user2 ESSDN.
        /// </summary>
        protected string User2ESSDN
        {
            get { return this.user2ESSDN; }
            set { this.user2ESSDN = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user2 name.
        /// </summary>
        protected string User2Name
        {
            get { return this.user2Name; }
            set { this.user2Name = value; }
        }

        /// <summary>
        /// Gets or sets a value for the domain name.
        /// </summary>
        protected string Domain
        {
            get { return this.domain; }
            set { this.domain = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user1 name.
        /// </summary>
        protected string User1Name
        {
            get { return this.user1Name; }
            set { this.user1Name = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user1 password.
        /// </summary>
        protected string User1Password
        {
            get { return this.user1Password; }
            set { this.user1Password = value; }
        }

        /// <summary>
        /// Gets or sets a value for the user2 password.
        /// </summary>
        protected string User2Password
        {
            get { return this.user2Password; }
            set { this.user2Password = value; }
        }

        /// <summary>
        /// Gets or sets the time of waiting for the rule to take effect.
        /// </summary>
        protected int WaitForTheRuleToTakeEffect
        {
            get { return this.waitForTheRuleToTakeEffect; }
            set { this.waitForTheRuleToTakeEffect = value; }
        }

        /// <summary>
        /// Gets or sets a value which specifies the maximum Size allowed for a property value returned.
        /// </summary>
        protected ushort PropertySizeLimitFlag
        {
            get { return this.propertySizeLimitFlag; }
            set { this.propertySizeLimitFlag = value; }
        }

        /// <summary>
        /// Gets or sets the time of waiting for setting OFF complete.
        /// </summary>
        protected int WaitForSetOOFComplete
        {
            get { return this.waitForSetOOFComplete; }
            set { this.waitForSetOOFComplete = value; }
        }
        #endregion

        #region TestCase Helper
        /// <summary>
        /// Log on mailbox with specified user.
        /// </summary>
        /// <param name="testUser">Identify the user to log on to the server.</param>
        public void LogonMailbox(TestUser testUser)
        {
            this.testUser = testUser;
            RopLogonResponse logonResponse = new RopLogonResponse();
            if (testUser == TestUser.TestUser1)
            {
                this.OxoruleAdapter.Connect(ConnectionType.PrivateMailboxServer, this.User1Name, this.User1ESSDN, this.User1Password);

                // Logon to private mailbox.
                this.LogonHandle = this.OxoruleAdapter.RopLogon(LogonType.Mailbox, this.User1ESSDN, out logonResponse);
            }
            else
            {
                this.OxoruleAdapter.Connect(ConnectionType.PrivateMailboxServer, this.User2Name, this.User2ESSDN, this.User2Password);

                // Logon to private mailbox.
                this.LogonHandle = this.OxoruleAdapter.RopLogon(LogonType.Mailbox, this.User2ESSDN, out logonResponse);
            }

            // Get the mailbox's GUID
            this.MailboxGUID = logonResponse.MailboxGuid;

            // Get folder ids.
            this.InboxFolderID = logonResponse.FolderIds[4];
            this.DAFFolderID = logonResponse.FolderIds[1];
            outBoxFolderID = logonResponse.FolderIds[5];
            sentItemsFolderID = logonResponse.FolderIds[6];

            // Open inbox and get inbox folder handle.
            RopOpenFolderResponse openInboxResponse = new RopOpenFolderResponse();
            this.InboxFolderHandle = this.OxoruleAdapter.RopOpenFolder(this.LogonHandle, this.InboxFolderID, out openInboxResponse);

            // Open DAF folder and get daf folder handle.
            RopOpenFolderResponse openDAFResponse = new RopOpenFolderResponse();
            this.DAFFolderHandle = this.OxoruleAdapter.RopOpenFolder(this.LogonHandle, this.DAFFolderID, out openDAFResponse);

            // Open outbox folder and get outbox folder handle.
            RopOpenFolderResponse openOutBoxResponse = new RopOpenFolderResponse();
            outBoxFolderHandle = this.OxoruleAdapter.RopOpenFolder(this.LogonHandle, outBoxFolderID, out openOutBoxResponse);

            // Open sent items folder and get sent items folder handle.
            RopOpenFolderResponse openSentItemsFolderResponse = new RopOpenFolderResponse();
            sentItemsFolderHandle = this.OxoruleAdapter.RopOpenFolder(this.LogonHandle, sentItemsFolderID, out openSentItemsFolderResponse);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R508");

            // Verify MS-OXORULE requirement: MS-OXORULE_R508.
            // If the DAF folder handle is not zero. It means the DAF has been created.
            Site.CaptureRequirementIfAreNotEqual<ulong>(
                0,
                this.DAFFolderHandle,
                508,
                @"[In Initialization] Prior to any client connecting to a mailbox, the server MUST ensure that the DAF has been created for that mailbox as specified in [MS-OXOSFLD] section 3.1.4.1.");
        }

        /// <summary>
        /// Submit a message to the server
        /// </summary>
        /// <param name="mailTo">The user name of the recipient</param>
        /// <param name="mailToUserDN">The userDN of the recipient</param>
        /// <param name="subject">The subject of the mail</param>
        /// <param name="addedProperties">The added properties of the mail</param>
        /// <returns>Return code of the message delivering</returns>
        public uint DeliverMessageToTriggerRule(string mailTo, string mailToUserDN, string subject, TaggedPropertyValue[] addedProperties)
        {
            RopCreateMessageResponse createMsgRes;
            uint msgHandle = this.OxoruleAdapter.RopCreateMessage(outBoxFolderHandle, outBoxFolderID, 0, out createMsgRes);
            TaggedPropertyValue[] clientSpecificProperties = new TaggedPropertyValue[1];
            clientSpecificProperties[0] = new TaggedPropertyValue();
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagSubject,
                PropertyType = (ushort)PropertyType.PtypString
            };
            clientSpecificProperties[0].PropertyTag = propertyTag;
            clientSpecificProperties[0].Value = Encoding.Unicode.GetBytes(subject + "\0");
            this.OxoruleAdapter.RopSetProperties(msgHandle, clientSpecificProperties);
            if (addedProperties != null && addedProperties.Length > 0)
            {
                this.OxoruleAdapter.RopSetProperties(msgHandle, addedProperties);
            }

            #region recipientColumns,  PropertyTag[] sampleRecipientColumns
            // The following sample data is from MS-OXCMSG 4.7.1
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            PropertyTag tag = new PropertyTag
            {
                PropertyId = (ushort)PropertyId.PidTagObjectType,
                PropertyType = (ushort)PropertyType.PtypInteger32
            };

            // PidTagObjectType
            sampleRecipientColumns[0] = tag;

            // PidTagDisplayType
            tag.PropertyId = (ushort)PropertyId.PidTagDisplayType;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[1] = tag;

            // PidTagAddressBookDisplayNamePrintable
            tag.PropertyId = (ushort)PropertyId.PidTagAddressBookDisplayNamePrintable;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            sampleRecipientColumns[2] = tag;

            // PidTagSmtpAddress
            tag.PropertyId = (ushort)PropertyId.PidTagSmtpAddress;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            sampleRecipientColumns[3] = tag;

            // PidTagSendInternetEncoding
            tag.PropertyId = (ushort)PropertyId.PidTagSendInternetEncoding;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[4] = tag;

            // PidTagDisplayTypeEx
            tag.PropertyId = (ushort)PropertyId.PidTagDisplayTypeEx;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[5] = tag;

            // PidTagRecipientDisplayName
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientDisplayName;
            tag.PropertyType = (ushort)PropertyType.PtypString;
            sampleRecipientColumns[6] = tag;

            // PidTagRecipientFlags
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientFlags;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[7] = tag;

            // PidTagRecipientTrackStatus
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientTrackStatus;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[8] = tag;

            // PidTagRecipientResourceState
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientResourceState;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[9] = tag;

            // PidTagRecipientOrder
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientOrder;
            tag.PropertyType = (ushort)PropertyType.PtypInteger32;
            sampleRecipientColumns[10] = tag;

            // PidTagRecipientEntryId
            tag.PropertyId = (ushort)PropertyId.PidTagRecipientEntryId;
            tag.PropertyType = (ushort)PropertyType.PtypBinary;
            sampleRecipientColumns[11] = tag;
            #endregion

            #region Configure a StandardPropertyRow: propertyRow
            PropertyValue[] propertyValueArray = new PropertyValue[12];
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            // PidTagObjectType
            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006);

            // PidTagDisplayType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000);

            // PidTagAddressBookDisplayNamePrintable
            propertyValueArray[2].Value = Encoding.Unicode.GetBytes(mailTo + "\0");

            // PidTagSmtpAddress
            propertyValueArray[3].Value = Encoding.Unicode.GetBytes(mailTo + "@" + this.Domain + "\0");

            // PidTagSendInternetEncoding
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000);

            // PidTagDisplayTypeEx
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000);

            // PidTagRecipientDisplayName
            propertyValueArray[6].Value = Encoding.Unicode.GetBytes(mailTo + "\0");

            // PidTagRecipientFlags
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001);

            // PidTagRecipientTrackStatus
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientResourceState
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000);

            // PidTagRecipientOrder
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000);

            AddressBookEntryID addressBookEntryID = new AddressBookEntryID(mailToUserDN);
            propertyValueArray[11].Value = Common.AddInt16LengthBeforeBinaryArray(addressBookEntryID.Serialize());

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = 0x01,
                PropertyValues = propertyValues
            };
            #endregion

            RecipientRow recipientRow = new RecipientRow
            {
                RecipientFlags = 0x065B,
                DisplayName = Encoding.Unicode.GetBytes(mailTo + "\0"),
                EmailAddress = Encoding.Unicode.GetBytes(mailTo + "@" + this.Domain + "\0"),
                SimpleDisplayName = Encoding.Unicode.GetBytes(mailTo + "\0"),
                RecipientColumnCount = 0x000C,
                RecipientProperties = propertyRow
            };

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = 0x00000000,
                RecipientType = 0x01,
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };

            ModifyRecipientRow[] sampleModifyRecipientRows = new ModifyRecipientRow[1];
            sampleModifyRecipientRows[0] = modifyRecipientRow;
            this.OxoruleAdapter.RopModifyRecipients(msgHandle, sampleRecipientColumns, sampleModifyRecipientRows);
            this.OxoruleAdapter.RopSaveChangesMessage(msgHandle);
            RopSubmitMessageResponse submitMsgRes = this.OxoruleAdapter.RopSubmitMessage(msgHandle, 0);
            return submitMsgRes.ReturnValue;
        }

        /// <summary>
        /// Get store object EntryID bytes array.
        /// </summary>
        /// <param name="storeObjectType">Identify the store object is a mailbox or a public folder.</param>
        /// <param name="serverShortName">Server short name string.</param>
        /// <param name="mailboxDN">Mailbox DN string.</param>
        /// <returns>Store object EntryID bytes array.</returns>
        public byte[] GetStoreObjectEntryID(StoreObjectType storeObjectType, string serverShortName, string mailboxDN)
        {
            StoreObjectEntryID objectEntryId = new StoreObjectEntryID(storeObjectType)
            {
                ServerShortname = serverShortName,
                MailBoxDN = mailboxDN
            };
            return objectEntryId.Serialize();
        }

        #endregion

        /// <summary>
        /// Overrides TestClassBase's TestInitialize()
        /// </summary>
        protected override void TestInitialize()
        {
            // Initial protocol adapter.
            this.oxoruleAdapter = Site.GetAdapter<IMS_OXORULEAdapter>();

            // Check whether the test initialize process should be disabled.
            if (this.CheckIfTestCaseDisabled())
            {
                return;
            }

            // Initial SUT control adapters.
            this.sutAdapter = Site.GetAdapter<IMS_OXORULESUTControlAdapter>();
            this.sutSetOOFAdapter = Site.GetAdapter<IMS_OXWOOFSUTControlAdapter>();
            this.server = Common.GetConfigurationPropertyValue(Constants.Server, this.Site);
            this.user1ESSDN = Common.GetConfigurationPropertyValue(Constants.User1ESSDN, this.Site);
            this.domain = Common.GetConfigurationPropertyValue(Constants.Domain, this.Site);
            this.user1Name = Common.GetConfigurationPropertyValue(Constants.User1Name, this.Site);
            this.user1Password = Common.GetConfigurationPropertyValue(Constants.User1Password, this.Site);
            this.user2Password = Common.GetConfigurationPropertyValue(Constants.User2Password, this.Site);
            this.user2Name = Common.GetConfigurationPropertyValue(Constants.User2Name, this.Site);
            this.user2ESSDN = Common.GetConfigurationPropertyValue(Constants.User2ESSDN, this.Site);
            this.waitForTheRuleToTakeEffect = int.Parse(Common.GetConfigurationPropertyValue(Constants.WaitForTheRuleToTakeEffect, this.Site));
            this.getMessageRepeatTime = ushort.Parse(Common.GetConfigurationPropertyValue(Constants.GetMessageRepeatTime, this.Site));
            this.propertySizeLimitFlag = ushort.Parse(Common.GetConfigurationPropertyValue(Constants.PropertySizeLimit, this.Site));
            this.waitForSetOOFComplete = int.Parse(Common.GetConfigurationPropertyValue(Constants.WaitForSetOOFComplete, this.Site));

            // Logon testuser1 mailbox.
            this.LogonMailbox(TestUser.TestUser1);
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            // Check whether the test cleanup process should be disabled.
            if (this.CheckIfTestCaseDisabled())
            {
                return;
            }

            // Clean current user's mailbox.
            this.ClearAllRules();
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(outBoxFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(this.DAFFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(sentItemsFolderHandle, 0);
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);

            // Clean another user's mailbox.
            if (this.testUser == TestUser.TestUser1)
            {
                // Logon testuser2.
                this.LogonMailbox(TestUser.TestUser2);
            }
            else
            {
                // Logon testuser1.
                this.LogonMailbox(TestUser.TestUser1);
            }

            this.ClearAllRules();
            this.OxoruleAdapter.RopEmptyFolder(this.InboxFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(outBoxFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(this.DAFFolderHandle, 0);
            this.OxoruleAdapter.RopEmptyFolder(sentItemsFolderHandle, 0);

            this.OxoruleAdapter.CleanUp();
            this.OxoruleAdapter.Reset();
            Thread.Sleep(this.WaitForTheRuleToTakeEffect);
        }

        /// <summary>
        /// Get the expected message properties included in a specific contents table and the message count in the table after retry preconfigured times.
        /// </summary>
        /// <param name="folderHandle">Handle of a specific folder.</param>
        /// <param name="contentsTableHandle">Handle of a specific contents table.</param>
        /// <param name="propertyTagList">>Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <param name="rowCount">The row count in the specific contents table, in which each row specifies a message.</param>
        /// <param name="expectedMessageCount">Specify the expected number of messages in target mailbox. The default value is 1.</param>
        /// <param name="expectedTitle">The default value is null. If the value is not null, the message with specified title must exist in target mailbox. If the expected message exists, the rowCount represents the index of the message in the target mailbox.</param>
        /// <returns>Response of the RopQueryRow ROP contents the expected message properties.</returns>
        protected RopQueryRowsResponse GetExpectedMessage(uint folderHandle, ref uint contentsTableHandle, PropertyTag[] propertyTagList, ref uint rowCount, int expectedMessageCount = 1, string expectedTitle = null)
        {
            RopQueryRowsResponse getNormalMailMessageContent;
            uint repeatTime = 0;
            bool isExpectedPropertyInPropertyList = false;

            // If the number of messages is greater than or equal the expectedMessageCount, terminates the loop
            while (rowCount < expectedMessageCount && repeatTime < this.getMessageRepeatTime)
            {
                RopGetContentsTableResponse ropGetContentsTableResponse = this.OxoruleAdapter.RopGetContentsTable(folderHandle, ContentTableFlag.None, out contentsTableHandle);
                Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponse.ReturnValue, "Getting contents table should succeed.");
                rowCount = ropGetContentsTableResponse.RowCount;
                if (rowCount > 0 && expectedTitle != null)
                {
                    getNormalMailMessageContent = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandle, propertyTagList);
                    Site.Assert.AreEqual<uint>(0, getNormalMailMessageContent.ReturnValue, "Getting mail message operation should succeed.");

                    for (int i = 0; i < propertyTagList.Length; i++)
                    {
                        if (propertyTagList[i].PropertyId == (ushort)PropertyId.PidTagSubject)
                        {
                            isExpectedPropertyInPropertyList = true;
                            for (int j = 0; j < getNormalMailMessageContent.RowData.PropertyRows.Count; j++)
                            {
                                string mailSubject = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[j].PropertyValues[i].Value);
                                if (mailSubject.Contains(expectedTitle))
                                {
                                    rowCount = (uint)j + 1;
                                    if (rowCount == expectedMessageCount)
                                    {
                                        return getNormalMailMessageContent;
                                    }
                                }
                            }
                        }
                    }

                    Site.Assert.IsTrue(isExpectedPropertyInPropertyList, "The property PidTagSubject should be included in the property list.");
                }

                if (rowCount >= expectedMessageCount || repeatTime == this.getMessageRepeatTime - 1)
                {
                    break;
                }
                else
                {
                    repeatTime++;
                    Thread.Sleep(this.WaitForTheRuleToTakeEffect);
                }
            }

            if (expectedTitle != null)
            {
                Site.Assert.Fail("Can't find the message which title is [{0}] in the target mailbox.", expectedTitle);
            }

            getNormalMailMessageContent = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandle, propertyTagList);
            Site.Assert.AreEqual<uint>(0, getNormalMailMessageContent.ReturnValue, "Getting mail message operation should succeed.");
            return getNormalMailMessageContent;
        }

        /// <summary>
        /// Get the expected message properties included in a specific contents table after retry preconfigured times.
        /// </summary>
        /// <param name="folderHandle">Handle of a specific folder.</param>
        /// <param name="contentsTableHandle">Handle of a specific contents table.</param>
        /// <param name="propertyTagList">>Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <param name="expectedMessageIndex">The index of the specific message in the table.</param>
        /// <param name="expectedPropertyValue">The value of a specific property of the message to be found in the target mailbox.</param>
        /// <param name="expectedPropertyName">The property name of a specific property of the message to be found in the target mailbox, which type should be string. The default property name is PidTagSubject.</param>
        /// <returns>Response of the RopQueryRow ROP contents the expected message properties.</returns>
        protected RopQueryRowsResponse GetExpectedMessage(uint folderHandle, ref uint contentsTableHandle, PropertyTag[] propertyTagList, ref int expectedMessageIndex, string expectedPropertyValue, PropertyId expectedPropertyName = PropertyId.PidTagSubject)
        {
            RopQueryRowsResponse getNormalMailMessageContent = new RopQueryRowsResponse();
            uint repeatTime = 0;
            uint rowCount = 0;
            bool isExpectedPropertyInPropertyList = false;

            // If retry time more than expected, terminates the loop
            while (repeatTime < this.getMessageRepeatTime)
            {
                RopGetContentsTableResponse ropGetContentsTableResponse = this.OxoruleAdapter.RopGetContentsTable(folderHandle, ContentTableFlag.None, out contentsTableHandle);
                Site.Assert.AreEqual<uint>(0, ropGetContentsTableResponse.ReturnValue, "Getting contents table should succeed.");
                rowCount = ropGetContentsTableResponse.RowCount;
                repeatTime++;

                if (rowCount > 0)
                {
                    getNormalMailMessageContent = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandle, propertyTagList);
                    Site.Assert.AreEqual<uint>(0, getNormalMailMessageContent.ReturnValue, "Getting mail message operation should succeed.");

                    for (int i = 0; i < propertyTagList.Length; i++)
                    {
                        if (propertyTagList[i].PropertyId == (ushort)expectedPropertyName)
                        {
                            isExpectedPropertyInPropertyList = true;
                            for (int j = 0; j < getNormalMailMessageContent.RowData.PropertyRows.Count; j++)
                            {
                                string propertyValue = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[j].PropertyValues[i].Value);
                                if (propertyValue.Contains(expectedPropertyValue))
                                {
                                    expectedMessageIndex = j;
                                    return getNormalMailMessageContent;
                                }
                            }
                        }
                    }

                    Site.Assert.IsTrue(isExpectedPropertyInPropertyList, "The property {0} to be checked should be included in the property list.", expectedPropertyName.ToString());
                }

                if (repeatTime == this.getMessageRepeatTime)
                {
                    break;
                }

                Thread.Sleep(this.WaitForTheRuleToTakeEffect);
            }

            Site.Assert.Fail("Can't find the message which has a property {0} ant its value is {1} in the target mailbox.", expectedPropertyName.ToString(), expectedPropertyValue);
            return getNormalMailMessageContent;
        }

        /// <summary>
        /// Check if an unexpected message with a specific property value exists in the target mailbox.
        /// </summary>
        /// <param name="folderHandle">Handle of a specific folder.</param>
        /// <param name="contentsTableHandle">Handle of a specific contents table.</param>
        /// <param name="propertyTagList">>Array of PropertyTag structures. This field specifies the property values that are visible in table rows.</param>
        /// <param name="unexpectedPropertyValue">The value of a specific property of the message to be checked in the target mailbox.</param>
        /// <param name="propertyName">The property name of a specific property of the message to be checked in the target mailbox, which type should be string. The default property name is PidTagSubject.</param>
        /// <returns>A bool value indicates whether a message with specific property value exists in the target mailbox. True means exist, otherwise not.</returns>
        protected bool CheckUnexpectedMessageExist(uint folderHandle, ref uint contentsTableHandle, PropertyTag[] propertyTagList, string unexpectedPropertyValue, PropertyId propertyName = PropertyId.PidTagSubject)
        {
            bool doesUnexpectedMessageExist = false;
            bool isExpectedPropertyInPropertyList = false;
            RopGetContentsTableResponse ropGetContentTableResponse = this.OxoruleAdapter.RopGetContentsTable(folderHandle, ContentTableFlag.None, out contentsTableHandle);
            Site.Assert.AreEqual<uint>(0, ropGetContentTableResponse.ReturnValue, "Getting contents table should succeed.");

            RopQueryRowsResponse getNormalMailMessageContent = this.OxoruleAdapter.QueryPropertiesInTable(contentsTableHandle, propertyTagList);
            Site.Assert.AreEqual<uint>(0, getNormalMailMessageContent.ReturnValue, "Getting mail message operation should succeed.");
            if (getNormalMailMessageContent.RowData.Count > 0)
            {
                for (int i = 0; i < propertyTagList.Length; i++)
                {
                    if (propertyTagList[i].PropertyId == (ushort)propertyName)
                    {
                        isExpectedPropertyInPropertyList = true;
                        for (int j = 0; j < getNormalMailMessageContent.RowData.PropertyRows.Count; j++)
                        {
                            string propertyValue = AdapterHelper.PropertyValueConvertToString(getNormalMailMessageContent.RowData.PropertyRows[j].PropertyValues[i].Value);
                            Site.Log.Add(LogEntryKind.Debug, "The value of the {0} property of the No.{1} message is : {2}", propertyName.ToString(), j + 1, propertyValue);
                            if (propertyValue.Contains(unexpectedPropertyValue))
                            {
                                doesUnexpectedMessageExist = true;
                                return doesUnexpectedMessageExist;
                            }
                        }
                    }
                }

                Site.Assert.IsTrue(isExpectedPropertyInPropertyList, "The property {0} to be checked should be included in the property list.", propertyName.ToString());
            }

            return doesUnexpectedMessageExist;
        }

        /// <summary>
        /// Get the Notify response after retry preconfigured times.
        /// </summary>
        /// <returns>Response of the RopNotify ROP.</returns>
        protected RopNotifyResponse GetNotifyResponse()
        {
            // Get notification detail from server.
            RopNotifyResponse ropNotifyResponse;
            uint repeatTime = 0;
            do
            {
                ropNotifyResponse = this.OxoruleAdapter.NotificationProcess();
                repeatTime++;
                if (ropNotifyResponse.NotificationData == null)
                {
                    Thread.Sleep(this.WaitForTheRuleToTakeEffect);
                }
                else
                {
                    break;
                }
            }
            while (ropNotifyResponse.NotificationData == null && repeatTime < this.getMessageRepeatTime);
            Site.Assert.IsNotNull(ropNotifyResponse.NotificationData, "The notification data should not be null!");
            return ropNotifyResponse;
        }

        /// <summary>
        /// Verify rule table.
        /// </summary>
        protected void VerifyRuleTable()
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXORULE_R121");

            // Verify MS-OXORULE requirement: MS-OXORULE_R121
            // Since the rows are retrieved by the QueryPropertiesInTable ROP using the table handle returned by the RopGetRulesTable ROP.
            // If the rule table is got successfully and the rule count is correct, the following requirement can be captured directly.
            this.Site.CaptureRequirement(
                121,
                @"[In RopGetRulesTable ROP] The RopGetRulesTable ROP ([MS-OXCROPS] section 2.2.11.2) creates a Table object through which the client can access the standard rules in a folder using table operations as specified in [MS-OXCTABL].");
        }

        /// <summary>
        /// Disable the test case if MAPIHTTP transport is selected but not supported by current test environment.
        /// </summary>
        protected void CheckMAPIHTTPTransportSupported()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture) == "mapi_http" && !Common.IsRequirementEnabled(933, this.Site))
            {
                Site.Assume.Inconclusive("This case can not run with MAPIHTTP transport with current SUT version. MAPIHTTP is supported from Exchange Server 2013 SP1.");
            }
        }

        /// <summary>
        /// Clear all rules of the inbox folder.
        /// </summary>
        protected void ClearAllRules()
        {
            // Call RopGetRulesTable with valid TableFlags.
            RopGetRulesTableResponse ropGetRulesTableResponse;
            uint ruleTableHandle = this.OxoruleAdapter.RopGetRulesTable(this.InboxFolderHandle, TableFlags.Normal, out ropGetRulesTableResponse);

            if (ropGetRulesTableResponse.ReturnValue == 0)
            {
                // Get rule ID if rules exist.
                PropertyTag ruleIDTag = new PropertyTag
                {
                    PropertyId = (ushort)PropertyId.PidTagRuleId,
                    PropertyType = (ushort)PropertyType.PtypInteger64
                };
                RopQueryRowsResponse queryRowsResponse = this.OxoruleAdapter.QueryPropertiesInTable(ruleTableHandle, new PropertyTag[1] { ruleIDTag });
                Site.Assert.AreEqual<uint>(0, queryRowsResponse.ReturnValue, "Retrieving rows from the rule table should succeed.");

                // Delete all rules if exist.
                if (queryRowsResponse.RowCount > 0)
                {
                    RuleData[] ruleDatas = new RuleData[queryRowsResponse.RowCount];
                    for (int i = 0; i < queryRowsResponse.RowCount; i++)
                    {
                        ulong ruleId = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[i].PropertyValues[0].Value, 0);
                        RuleProperties ruleProperties = new RuleProperties();
                        ruleDatas[i] = AdapterHelper.GenerateValidRuleData(ActionType.OP_MARK_AS_READ, TestRuleDataType.ForRemove, 0, RuleState.ST_ENABLED, null, ruleProperties, ruleId);
                    }

                    RopModifyRulesResponse modifyRulesResponse = this.OxoruleAdapter.RopModifyRules(this.InboxFolderHandle, ModifyRuleFlag.Modify_OnExisting, ruleDatas);
                    Site.Assert.AreEqual<uint>(0, modifyRulesResponse.ReturnValue, "Deleting rule should be success");

                    // Wait the rule to be deleted
                    Thread.Sleep(this.WaitForTheRuleToTakeEffect);
                }
            }
        }

        /// <summary>
        /// Check if a specific test case is disabled.
        /// </summary>
        /// <returns>A bool value indicates whether a test case is disabled. True means disabled, otherwise not.</returns>
        private bool CheckIfTestCaseDisabled()
        {
            bool testCaseDisabled = false;
            bool transportIsMAPI = Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower(System.Globalization.CultureInfo.CurrentCulture) == "mapi_http";
            if (transportIsMAPI && !Common.IsRequirementEnabled(933, this.Site))
            {
                testCaseDisabled = true;
            }
            else
            {
                switch (TestContext.TestName)
                {
                    case "MSOXORULE_S01_TC09_AddExtendedRuleForThreeTimes":
                        testCaseDisabled = !Common.IsRequirementEnabled(646, this.Site);
                        break;
                    case "MSOXORULE_S02_TC10_ServerExecuteRule_Action_OP_MOVE":
                        testCaseDisabled = !Common.IsRequirementEnabled(929, this.Site);
                        break;
                    case "MSOXORULE_S02_TC18_ServerExecuteRule_Action_MultipleOP_MOVE":
                        testCaseDisabled = !Common.IsRequirementEnabled(929, this.Site) || !Common.IsRequirementEnabled(904, this.Site);
                        break;
                    case "MSOXORULE_S02_TC11_ServerExecuteRule_Action_OP_COPY":
                        testCaseDisabled = !Common.IsRequirementEnabled(298, this.Site);
                        break;
                    case "MSOXORULE_S03_TC04_OOFBehaviorsForST_KEEP_OOF_HIST":
                        testCaseDisabled = !Common.IsRequirementEnabled(583, this.Site);
                        break;
                    case "MSOXORULE_S03_TC06_OOFBehaviorsForFlagSameSemanticAsST_ONLY_WHEN_OOF":
                        testCaseDisabled = !Common.IsRequirementEnabled(625, this.Site);
                        break;
                    case "MSOXORULE_S03_TC07_OOFBehaviorsForFlagDisableSpecificOOFRule":
                        testCaseDisabled = !Common.IsRequirementEnabled(624, this.Site);
                        break;
                    case "MSOXORULE_S05_TC01_ServerGenerateOneDAM_ForOP_DEFER_ACTION_BelongToSameRuleProvider":
                        testCaseDisabled = !Common.IsRequirementEnabled(908, this.Site);
                        break;
                    default:
                        break;
                }
            }

            return testCaseDisabled;
        }
    }
}