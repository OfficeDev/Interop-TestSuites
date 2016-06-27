namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that Verifies the requirements related to RopSetMessageReadFlags and RopSetReadFlags operations.
    /// </summary>
    [TestClass]
    public class S03_SetMessageFlags : TestSuiteBase
    {
        /// <summary>
        /// A Boolean indicates whether test case create a public folder.
        /// </summary>
        private bool isCreatePulbicFolder = false;

        /// <summary>
        /// The pulbic FolderID created by test case.
        /// </summary>
        private ulong publicFolderID;

        #region Test Case Initialization
        /// <summary>
        ///  Initializes the test class before running the test cases in the class.
        /// </summary>
        /// <param name="testContext">Test context which used to store information that is provided to unit tests.</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.Initialize(testContext);
        }

        /// <summary>
        /// Clear the class.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestSuiteBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case verifies the requirements related to RopSetReadFlags ROP operation.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC01_RopSetReadFlags()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagChangeKey],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreatorEntryId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRead],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRecipientDisplayName]
            };

            List<PropertyObj> ps = new List<PropertyObj>();

            #region Call RopLogon to logon the specific private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a new Message object in inbox folder.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[5], this.insideObjHandle);
            #endregion

            #region Call RopModifyRecipients to add recipient to message created by step2
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(Common.GetConfigurationPropertyValue("AdminUserName", this.Site), 0)
            };
            this.AddRecipients(modifyRecipientRow, targetMessageHandle, propertyTag);
            #endregion

            #region Call RopSetProperties to set the properties of message created by step 2.
            PropertyTag[] tagArray = propertyTags.ToArray();

            // Get properties for Created Message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            
            List<PropertyObj> propertyList = new List<PropertyObj>();
            string title = Common.GenerateResourceName(Site, "Mail");
            propertyList.Add(new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(title)));

            // Set PidTagReadReceiptRequested property.
            propertyList.Add(new PropertyObj(0x0029, (ushort)PropertyType.PtypBoolean, new byte[] { 0x01 }));

            // Set PidTagNonReceiptNotificationRequested property.
            propertyList.Add(new PropertyObj(0x0C06, (ushort)PropertyType.PtypBoolean, new byte[] { 0x01 }));

            // Set PidTagReadReceiptAddressType property.
            propertyList.Add(new PropertyObj(0x4029, (ushort)PropertyType.PtypString, Common.GetBytesFromUnicodeString("EX")));

            // Set PidTagReadReceiptEmailAddress property.
            string userName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            string mailAddress = string.Format("{0}@{1}", userName, domain);

            propertyList.Add(new PropertyObj(0x402A, (ushort)PropertyType.PtypString, Common.GetBytesFromUnicodeString(mailAddress)));

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save the Message object created by step 2.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSubmitMessage to submit message
            RopSubmitMessageRequest submitMessageRequest = new RopSubmitMessageRequest()
            {
                RopId = (byte)RopId.RopSubmitMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                SubmitFlags = (byte)SubmitFlags.None
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(submitMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSubmitMessageResponse submitMessageResponse = (RopSubmitMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, submitMessageResponse.ReturnValue, "Call RopSubmitMessage should success.");
            #endregion

            #region Receive the message sent by step 2.
            ulong messageId = 0;
            bool isMessageReceived = this.WaitEmailBeDelivered(title, logonResponse.FolderIds[4], this.insideObjHandle, out messageId);

            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion

            #region Verify requirements
            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageId;

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetBefore = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R529");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R529
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000100,
                Convert.ToInt32(pidTagMessageFlagsSetBefore.Value) & (int)MessageFlags.MfNotifyRead,
                529,
                @"[In PidTagMessageFlags Property] [mfNotifyRead (0x00000100)] The user who sent the message has requested notification when a recipient (1) first reads it.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R531");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R531
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000200,
                Convert.ToInt32(pidTagMessageFlagsSetBefore.Value) & (int)MessageFlags.MfNotifyUnread,
                531,
                @"[In PidTagMessageFlags Property] [mfNotifyUnread (0x00000200)] The user who sent the message has requested notification when a recipient (1) deletes it before reading or the Message object expires.");
            #endregion

            #region Call RopOpenFolder to open inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfDefault on Message object within inbox Folder.
            RopSetReadFlagsRequest setReadFlagsRequet = new RopSetReadFlagsRequest()
            {
                RopId = (byte)RopId.RopSetReadFlags,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                WantAsynchronous = 0x00, // Does not asynchronous 
                MessageIds = messageIds,
                MessageIdCount = Convert.ToUInt16(messageIds.Length),
                ReadFlags = (byte)ReadFlags.Default
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetReadFlagsResponse setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetDefault = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            PropertyObj pidTagChangeKeyBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);
            PropertyObj pidTagRead = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRead);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R815");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R815
            this.Site.CaptureRequirementIfAreEqual<int>(
                (int)MessageFlags.MfRead,
                Convert.ToInt32(pidTagMessageFlagsSetDefault.Value) & (int)MessageFlags.MfRead,
                815,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfDefault (0x00)] The server sets the read flag and sends the receipt.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2045");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2045
            this.Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(pidTagRead.Value),
                2045,
                @"[In PidTagRead Property] The PidTagRead property ([MS-OXPROPS] section 2.867) indicates whether a message has been read.");    
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfClearReadFlag on Message object within inbox Folder.
            setReadFlagsRequet.ReadFlags = (byte)ReadFlags.ClearReadFlag;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetClearReadFlag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            PropertyObj pidTagChangeKeyAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R822");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R822
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)MessageFlags.MfRead,
                Convert.ToInt32(pidTagMessageFlagsSetClearReadFlag.Value) & (int)MessageFlags.MfRead,
                822,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfClearReadFlag (0x04)] Server clears the mfRead bit.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R795");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R795
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                Convert.ToInt32(pidTagMessageFlagsSetDefault.Value),
                Convert.ToInt32(pidTagMessageFlagsSetClearReadFlag.Value),
                795,
                @"[In RopSetReadFlags ROP] The RopSetReadFlags ROP ([MS-OXCROPS] section 2.2.6.10) changes the state of the PidTagMessageFlags property (section 2.2.1.6) on one or more Message objects within a Folder object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1145");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1145
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                Convert.ToInt32(pidTagMessageFlagsSetDefault.Value),
                Convert.ToInt32(pidTagMessageFlagsSetClearReadFlag.Value),
                1145,
                @"[In PidTagMessageFlags Property] The pidTagMessageFlages property is modified using the RopSetReadFlags ROP ([MS-OXCROPS] section 2.2.6.10), as described in section 2.2.3.10.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1703, the PidTagMessageFlags value is {0}.", pidTagMessageFlagsSetClearReadFlag.Value);

            bool isSameChangekey = Common.CompareByteArray((byte[])pidTagChangeKeyBeforeSet.Value, (byte[])pidTagChangeKeyAfterSet.Value);
            Site.Assert.IsTrue(isSameChangekey, "The PidTagChangeKey property should not be changed.");

            Site.Assert.AreEqual<DateTime>(
                Convert.ToDateTime(pidTagLastModificationTimeBeforeSet.Value),
                Convert.ToDateTime(pidTagLastModificationTimeAfterSet.Value),
                "The PidTagLastModificationTime property should not be changed.");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1703
            // Because above step has verified only changes the PidTagMessageFlags property, not the PidTagChangeKey property and PidTagLastModificationTime property.
            // R1703 will be direct verified.
            this.Site.CaptureRequirement(
               1703,
               @"[In Receiving a RopSetReadFlags ROP Request] The server immediately commits the changes to the Message objects as if the Message objects had been opened and the RopSaveMessageChanges ROP ([MS-OXCROPS] section 2.2.6.3) had been called, except that it [server] only changes the PidTagMessageFlags property (section 2.2.1.6), not the PidTagChangeKey property ([MS-OXCFXICS] section 2.2.1.2.7), the PidTagLastModificationTime property (section 2.2.2.2), or any other property that is modified during a RopSaveChangesMessage ROP request ([MS-OXCROPS] section 2.2.6.3).");

            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfSuppressReceipt on Message object within inbox Folder.
            setReadFlagsRequet.ReadFlags = (byte)ReadFlags.SuppressReceipt;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetSuppressReceipt = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R818");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R818
            this.Site.CaptureRequirementIfAreEqual<int>(
                (int)MessageFlags.MfRead,
                Convert.ToInt32(pidTagMessageFlagsSetSuppressReceipt.Value) & (int)MessageFlags.MfRead,
                818,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfSuppressReceipt (0x01)] The server sets the mfRead bit.");
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfGenerateReceiptOnly on Message object within inbox Folder.
            setReadFlagsRequet.ReadFlags = (byte)ReadFlags.GenerateReceiptOnly;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R825");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R825
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                setReadFlagesResponse.ReturnValue,
                825,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfGenerateReceiptOnly (0x10)] The server sends a read receipt if one is pending, but does not change the mfRead bit.");
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfClearNotifyRead on Message object within inbox Folder.
            setReadFlagsRequet.ReadFlags = (byte)ReadFlags.ClearNotifyRead; // rfClearNotifyRead
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetClearNotifyRead = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R827");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R827
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)MessageFlags.MfNotifyRead,
                Convert.ToInt32(pidTagMessageFlagsSetClearNotifyRead.Value) & (int)MessageFlags.MfNotifyRead,
                827,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfClearNotifyRead (0x20)] The server clears the mfNotifyRead bit but does not send a read receipt.");
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property to rfClearNotifyUnread on Message object within inbox Folder.
            setReadFlagsRequet.ReadFlags = (byte)ReadFlags.ClearNotifyUnread; // rfClearNotifyUnread
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSetClearNotifyUnread = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R829");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R829
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)MessageFlags.MfNotifyUnread,
                Convert.ToInt32(pidTagMessageFlagsSetClearNotifyUnread.Value) & (int)MessageFlags.MfNotifyUnread,
                829,
                @"[In RopSetReadFlags ROP Request Buffer] [ReadFlags] [rfClearNotifyUnread (0x40)] The server clears the mfNotifyUnread bit but does not send a nonread receipt.");
            #endregion

            #region Call RopRelease to release all resources.
            this.ReleaseRop(targetMessageHandle);
            this.ReleaseRop(folderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the RopSetReadFlags ROP to be processed asynchronously.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC02_RopSetReadFlagsAsyncEnable()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon specific private mailbox and SUPPORT_PROGRESS flag is set in the OpenFlags field.
            string userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site) + "\0";
            RopLogonRequest logonRequest = new RopLogonRequest()
            {
                RopId = (byte)RopId.RopLogon,
                LogonId = CommonLogonId,
                OutputHandleIndex = 0x00, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                StoreState = 0,
                LogonFlags = (byte)LogonFlags.Private,
                OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping | 0x20000000,
                EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN),
                Essdn = Encoding.ASCII.GetBytes(userDN)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(logonRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopLogonResponse logonResponse = (RopLogonResponse)this.response;
            uint objHandle = this.ResponseSOHs[0][logonResponse.OutputHandleIndex];
            #endregion

            #region Call RopCreateMessage to create a new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], objHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message created by step 2.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ulong[] messageIds = new ulong[1];
            messageIds[0] = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopOpenFolder to open Inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], objHandle);
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property and the WantAsynchronous flag is set in RopSetReadFlags.
            RopSetReadFlagsRequest setReadFlagsRequet = new RopSetReadFlagsRequest
            {
                RopId = (byte)RopId.RopSetReadFlags,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                WantAsynchronous = 0x01,
                ReadFlags = (byte)ReadFlags.ClearReadFlag,
                MessageIds = messageIds,
                MessageIdCount = Convert.ToUInt16(messageIds.Length)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            #region Verify MS-OXCMSG_R1705, MS-OXCMSG_R1706
            if (Common.IsRequirementEnabled(1705, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1705, server returns response type is {0}.", this.response.GetType());

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1705
                bool isVerifiedR1705 = this.response is RopSetReadFlagsResponse;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR1705,
                    1705,
                    @"[In Appendix A: Product Behavior] Implementation does return a RopSetReadFlags ROP response if the WantAsynchronous flag is nonzero. (Exchange 2007 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1706, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1706,server returns response type is {0}.", this.response.GetType());

                this.Site.CaptureRequirementIfIsNotInstanceOfType(
                    this.response,
                    typeof(RopProgressResponse),
                    1706,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return a RopProgress ROP response instead if the WantAsynchronous flag is nonzero. (Exchange 2007 and above follow this behavior.)");
            }
            #endregion
            #endregion

            #region Call RopRelease to release all resources.
            this.ReleaseRop(targetMessageHandle);
            this.ReleaseRop(folderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests disabling the asynchronous processing of the RopSetReadFlags ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC03_RopSetReadFlagsAsyncDisable()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };
            List<PropertyObj> ps = new List<PropertyObj>();

            #region Call RopLogon to logon specific private mailbox and SUPPORT_PROGRESS flag is not set in the OpenFlags field.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the Message object created step 2.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ulong[] messageIds = new ulong[1];
            messageIds[0] = saveChangesMessageResponse.MessageId;

            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            #endregion

            #region Call RopOpenFolder to open inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property and the WantAsynchronous flag is set in RopSetReadFlags.
            // RopSetReadFlags
            RopSetReadFlagsRequest setReadFlagsRequet = new RopSetReadFlagsRequest()
            {
                RopId = (byte)RopId.RopSetReadFlags,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                WantAsynchronous = 0x01,
                ReadFlags = (byte)ReadFlags.ClearReadFlag,
                MessageIds = messageIds,
                MessageIdCount = Convert.ToUInt16(messageIds.Length)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetReadFlagsResponse setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setReadFlagesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #region Verify MS-OXCMSG_R1510, MS-OXCMSG_R1512, MS-OXCMSG_R1917, MS-OXCMSG_R1918
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1510");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1510
            bool isVerifiedR1510 = this.response is RopSetReadFlagsResponse;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1510,
                1510,
                @"[In Receiving a RopSetReadFlags ROP Request] If the server has not received a SUPPORT_PROGRESS flag in the request buffer of the RopLogon ROP ([MS-OXCROPS] section 2.2.3.1), the server MUST disable asynchronous processing for the RopSetReadFlags ROP ([MS-OXCROPS] section 2.2.6.10), overriding any value of the WantAsynchronous flag.");

            if (Common.IsRequirementEnabled(1917, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1917");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1917
                this.Site.CaptureRequirementIfIsNotInstanceOfType(
                    this.response,
                    typeof(RopProgressResponse),
                    1917,
                    @"[In Appendix A: Product Behavior] Implementation does support this behavior [disable asynchronous processing of the RopSetReadFlags ROP and doesn't return the RopProgress ROP whether or not the WantAsynchronous flag is set if the SUPPORT_PROGRESS flag is not set by the client in the OpenFlags field in the RopLogon ROP]. (Exchange 2010 SP2 and above follow this behavior.)");
            }
            
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1512");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1512
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)pidTagMessageFlagsBeforeSet.Value,
                (int)pidTagMessageFlagsAfterSet.Value,
                1512,
                @"[In Receiving a RopSetReadFlags ROP Request] If the client does not pass the SUPPORT_PROGRESS flag, the server will process the entire RopSetReadFlags ROP request before returning a response to the client.");
            #endregion
            #endregion

            #region Call RopRelease to release all resources.
            RopReleaseRequest releaseRequest = new RopReleaseRequest()
            {
                RopId = (byte)RopId.RopRelease, // RopId 0x01 indicates RopRelease
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(releaseRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            releaseRequest = new RopReleaseRequest()
            {
                RopId = (byte)RopId.RopRelease, // RopId 0x01 indicates RopRelease
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(releaseRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            #endregion
        }

        /// <summary>
        /// This test case tests RopSetMessageReadFlag ROP in public folder mode.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC04_RopSetMessageReadFlagsInPublicFolderMode()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PublicFolderServer);

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };
            List<PropertyObj> ps = new List<PropertyObj>();

            #region Call RopLogon to logon the public folder.
            RopLogonResponse logonResponse = this.Logon(LogonType.PublicFolder, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open the second folder.
            uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[1], this.insideObjHandle);
            #endregion

            #region Call RopCreateFolder to create a temporary public folder.
            ulong folderId = this.CreateSubFolder(openedFolderHandle);
            this.isCreatePulbicFolder = true;
            this.publicFolderID = folderId;
            #endregion

            #region Call RopOpenFolder to open the temporary public folder.
            this.OpenSpecificFolder(folderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message.
            uint targetMessageHandle = this.CreatedMessage(folderId, this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopLongTermIdFromId to get the Long Term Id.
            RopLongTermIdFromIdRequest longTermIdFromIdRequest = new RopLongTermIdFromIdRequest()
            {
                RopId = (byte)RopId.RopLongTermIdFromId, // RopId 0x43 indicates RopLongTermIdFromId
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                ObjectId = saveChangesMessageResponse.MessageId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(longTermIdFromIdRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopLongTermIdFromIdResponse longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, longTermIdFromIdResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the created message.
            uint openMessageHandle = this.OpenSpecificMessage(folderId, saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfDefault for the created message.
            RopSetMessageReadFlagRequest setMessageReadFlagRequest = new RopSetMessageReadFlagRequest()
            {
                RopId = (byte)RopId.RopSetMessageReadFlag,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                ClientData = longTermIdFromIdResponse.LongTermId.Serialize(),
                ReadFlags = (byte)ReadFlags.Default
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, openMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetMessageReadFlagResponse setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfSuppressReceipt for the created message.
            setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.SuppressReceipt;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, openMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfClearReadFlag for the created message.
            ps = this.GetSpecificPropertiesOfMessage(folderId, saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.ClearReadFlag;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, openMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(folderId, saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R921, the ReadStatusChanged field value is {0}.", setMessageReadFlagResponse.ReadStatusChanged);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R921
            bool isVerifiedR921 = setMessageReadFlagResponse.ReadStatusChanged > 0 && pidTagMessageFlagsBeforeSet.Value != pidTagMessageFlagsAfterSet.Value;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR921,
                921,
                @"[In RopSetMessageReadFlag ROP Response Buffer] [ReadStatusChanged] [The value non-zero indicates that] The read status on the Message object changed and the logon is in public folder mode.");
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfGenerateReceiptOnly for the created message.
            ps = this.GetSpecificPropertiesOfMessage(folderId, saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlagsBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.GenerateReceiptOnly;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, openMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            ps = this.GetSpecificPropertiesOfMessage(folderId, saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlagsAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R920, the ReadStatusChanged field value is {0}.", setMessageReadFlagResponse.ReadStatusChanged);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R920
            bool isVerifiedR920 = setMessageReadFlagResponse.ReadStatusChanged == 0 && Convert.ToInt32(pidTagMessageFlagsBeforeSet.Value) == Convert.ToInt32(pidTagMessageFlagsAfterSet.Value);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR920,
                920,
                @"[In RopSetMessageReadFlag ROP Response Buffer] [ReadStatusChanged] [The value 0x00 indicates that] The read status on the Message object was unchanged.");
            #endregion

            #region Call RopRelease to release the created folder and message
            this.ReleaseRop(openMessageHandle);
            this.ReleaseRop(openedFolderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the RopSetMessageReadFlag not in public folder mode.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC05_RopSetMessageReadFlagsNotInPublicFolderMode()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagChangeKey],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime]
            };
            List<PropertyObj> ps = new List<PropertyObj>();

            #region Call RopLogon to logon the private mailbox.
            // Create a message
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object in inbox folder.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenFolder request to open inbox folder.
            this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagMessageFlags property of created message.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            PropertyObj pidTagChangeKeyBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to 0x04 for the created message.
            RopSetMessageReadFlagRequest setMessageReadFlagRequest = new RopSetMessageReadFlagRequest()
            {
                RopId = (byte)RopId.RopSetMessageReadFlag,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                ReadFlags = (byte)ReadFlags.ClearReadFlag,
                ClientData = new byte[] { }
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetMessageReadFlagResponse setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1841");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1841
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                setMessageReadFlagResponse.ReadStatusChanged,
                1841,
                @"[In RopSetMessageReadFlag ROP Response Buffer] [ReadStatusChanged] [The value 0x00 indicates that] the logon is not in public folder mode.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagMessageFlags property of created message.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            PropertyObj pidTagChangeKeyAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R835");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R835
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)pidTagMessageFlagsBeforeSet.Value,
                (int)pidTagMessageFlagsAfterSet.Value,
                835,
                @"[In RopSetMessageReadFlag ROP] The RopSetMessageReadFlag ROP ([MS-OXCROPS] section 2.2.6.11) changes the state of the PidTagMessageFlags property (section 2.2.1.6) for the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R537");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R537
            this.Site.CaptureRequirementIfAreNotEqual<int>(
                (int)pidTagMessageFlagsBeforeSet.Value,
                (int)pidTagMessageFlagsAfterSet.Value,
                537,
                @"[In PidTagMessageFlags Property] The PidTagMessageFlags property is also modified using the RopSetMessageReadFlag ROP ([MS-OXCROPS] section 2.2.6.11), as described in section 2.2.3.11.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1710");

            bool isSameChangekey = Common.CompareByteArray((byte[])pidTagChangeKeyBeforeSet.Value, (byte[])pidTagChangeKeyAfterSet.Value);
            Site.Assert.IsTrue(isSameChangekey, "The PidTagChangeKey property should not be changed.");

            Site.Assert.AreEqual<DateTime>(
                Convert.ToDateTime(pidTagLastModificationTimeBeforeSet.Value),
                Convert.ToDateTime(pidTagLastModificationTimeAfterSet.Value),
                "The PidTagLastModificationTime property should not be changed.");

            // Because above step has verified only changes the PidTagMessageFlags property, not the PidTagChangeKey property and PidTagLastModificationTime property.
            // R1710 will be direct verified.
            this.Site.CaptureRequirement(
                1710,
                @"[In Receiving a RopSetMessageReadFlag ROP Request] The server immediately commits the changes to the Message object as if the Message object had been opened and the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) had been called, except that it [server] only changes the PidTagMessageFlags property (section 2.2.1.6), not the PidTagChangeKey property ([MS-OXCFXICS] section 2.2.1.2.7), the PidTagLastModificationTime property (section 2.2.2.2), or any other property that is modified during a RopSaveChangesMessage ROP request ([MS-OXCROPS] section 2.2.6.3).");
            #endregion
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfClearNotifyRead for the created message.
            setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.ClearNotifyRead;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetMessageReadFlag to set the ReadFlags to rfClearNotifyUnread for the created message.
            setMessageReadFlagRequest.ReadFlags = (byte)ReadFlags.ClearNotifyUnread;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageReadFlagResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRelease to release the created message
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the error code ecNullObject of the RopSetReadFlags ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC06_RopSetReadFlagsFailure()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the specific private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a new Message object in inbox folder.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopModifyRecipients to add recipient to message created by step2
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(Common.GetConfigurationPropertyValue("AdminUserName", this.Site), 0)
            };
            this.AddRecipients(modifyRecipientRow, targetMessageHandle, propertyTag);
            #endregion

            #region Call RopSaveChangesMessage to save the Message object created by step 2.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ulong[] messageIds = new ulong[1];
            messageIds[0] = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopGetContentsTable to get the contents table of inbox folder before submit message.
            RopGetContentsTableResponse getContentsTableResponse = this.GetContentTableSuccess(folderHandle);
            uint contentTableHandle = this.ResponseSOHs[0][getContentsTableResponse.OutputHandleIndex];
            uint rowCountBeforeSubmit = getContentsTableResponse.RowCount;
            this.ReleaseRop(contentTableHandle);
            #endregion

            #region Call RopSubmitMessage to submit message
            RopSubmitMessageRequest submitMessageRequest = new RopSubmitMessageRequest()
            {
                RopId = (byte)RopId.RopSubmitMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                SubmitFlags = (byte)SubmitFlags.None
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(submitMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSubmitMessageResponse submitMessageResponse = (RopSubmitMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, submitMessageResponse.ReturnValue, "Call RopSubmitMessage should success.");

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };
            List<PropertyObj> ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageIds[0], this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlagsSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R520");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R520
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000004,
                Convert.ToInt32(pidTagMessageFlagsSet.Value) & (int)MessageFlags.MfSubmitted,
                520,
                @"[In PidTagMessageFlags Property] [mfSubmitted (0x00000004)] The message is marked for sending as a result of a call to the RopSubmitMessage ROP.");
            #endregion

            #region Call RopSetReadFlags which contains an incorrect InputHandleIndex.
            // RopSetReadFlags
            RopSetReadFlagsRequest setReadFlagsRequet = new RopSetReadFlagsRequest()
            {
                RopId = (byte)RopId.RopSetReadFlags,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                WantAsynchronous = 0x00, // Does not asynchronous 
                ReadFlags = 0x00, // rfClearNotifyUnread
                MessageIds = messageIds,
                MessageIdCount = Convert.ToUInt16(messageIds.Length)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetReadFlagsResponse setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1516");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1516
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                setReadFlagesResponse.ReturnValue,
                1516,
                @"[In Receiving a RopSetReadFlags ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopSetReadFlags] was called does not refer to a Folder object.");
              #endregion

            #region Receive the message sent by step 7.
            bool isMessageReceived = WaitEmailBeDelivered(folderHandle, rowCountBeforeSubmit);
            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion
        }

        /// <summary>
        /// This test case tests the error code ecNullObject of the RopSetMessageReadFlags ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC07_RopSetMessageReadFlagsFailure()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            // Create a message
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object in inbox folder.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSetMessageReadFlag with invalid parameters to get the failure response.
            RopSetMessageReadFlagRequest setMessageReadFlagRequest = new RopSetMessageReadFlagRequest()
            {
                RopId = (byte)RopId.RopSetMessageReadFlag,
                LogonId = CommonLogonId,
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                InputHandleIndex = InvalidInputHandleIndex,
                ReadFlags = (byte)ReadFlags.ClearReadFlag,
                ClientData = new byte[] { }
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageReadFlagRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetMessageReadFlagResponse setMessageReadFlagResponse = (RopSetMessageReadFlagResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1521");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1521
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                setMessageReadFlagResponse.ReturnValue,
                1521,
                @"[In Receiving a RopSetMessageReadFlag ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopSetMessageReadFlag] was called does not refer to a Folder object.");
            #endregion

            #region Call RopRelease to release the created message
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the error code ecNullObject of the RopSetMessageReadFlags ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S03_TC08_RopSetReadFlagsWithPartialCompletion()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            // Create a message
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object in inbox folder.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong[] messageIds = new ulong[2];
            messageIds[0] = saveChangesMessageResponse.MessageId;
            messageIds[1] = saveChangesMessageResponse.MessageId + 1;

            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenFolder to open inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetReadFlags to change the state of the PidTagMessageFlags property and the MessageIds contains a MessageId that the message does not exist.
            // RopSetReadFlags
            RopSetReadFlagsRequest setReadFlagsRequet = new RopSetReadFlagsRequest()
            {
                RopId = (byte)RopId.RopSetReadFlags,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                WantAsynchronous = 0x01,
                ReadFlags = (byte)ReadFlags.ClearReadFlag,
                MessageIds = messageIds,
                MessageIdCount = Convert.ToUInt16(messageIds.Length)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setReadFlagsRequet, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetReadFlagsResponse setReadFlagesResponse = (RopSetReadFlagsResponse)this.response;
            if(Common.IsRequirementEnabled(3021,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3021");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3021
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    setReadFlagesResponse.PartialCompletion,
                    3021,
                    @"[In Appendix A: Product Behavior] [PartialCompletion] A nonzero value indicates the server was unable to modify one or more of the Message objects represented in the MessageIds field. (Exchange 2007 and exchange 2016 follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1514");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1514
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    setReadFlagesResponse.PartialCompletion,
                    1514,
                    @"[In Receiving a RopSetReadFlags ROP Request] If the server is unable to modify one or more of the Message objects that are specified in the MessageIds field, as specified in section 2.2.3.10.1, of the request buffer, then the server returns the PartialCompletion flag, as specified in section 2.2.3.10.2, in the response buffer.");
            }
            #endregion

            #region Call RopRelease to release all resources.
            RopReleaseRequest releaseRequest = new RopReleaseRequest()
            {
                RopId = (byte)RopId.RopRelease, // RopId 0x01 indicates RopRelease
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(releaseRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            releaseRequest = new RopReleaseRequest()
            {
                RopId = (byte)RopId.RopRelease, // RopId 0x01 indicates RopRelease
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(releaseRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            #endregion
        }
        /// <summary>
        /// Test cleanup method
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.isCreatePulbicFolder == true)
            {
                this.isNotNeedCleanupPrivateMailbox = true;

                #region Call RopLogon to logon the public folder.
                RopLogonResponse logonResponse = this.Logon(LogonType.PublicFolder, out this.insideObjHandle);
                #endregion

                #region Call RopOpenFolder to open the second folder.
                uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[1], this.insideObjHandle);
                #endregion

                #region Call RopDeleteFolder to delete the public folder created by test case
                RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest()
                {
                    RopId = (byte)RopId.RopDeleteFolder,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    DeleteFolderFlags = 0x01, // The folder and all of the Message objects in the folder are deleted.
                    FolderId = this.publicFolderID // Folder to be deleted
                };
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteFolderRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.response;
                Site.Assert.AreEqual<uint>(0x0000000, deleteFolderResponse.ReturnValue, "Delete Pulbic Folder should be success.");
                this.isCreatePulbicFolder = false;
                #endregion
            }

            base.TestCleanup();
        }
    }
}