namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that verify the requirements related to RopCreateMessage and RopSaveChangesMessage.
    /// </summary>
    [TestClass]
    public class S01_CreateAndSaveMessage : TestSuiteBase
    {
        #region Const definitions for test
        /// <summary>
        /// The value of SaveFlags is not supported.
        /// </summary>
        private const byte NotSupportedSaveFlags = 0x0F;

        /// <summary>
        /// Constant string for the test data of DateTime.
        /// </summary>
        private const string TestDataOfDateTime = "2010-8-8 19:05:59";
        #endregion

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
        /// This test case validates the operation of creating a message and saving the created message.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC01_RopCreateMessageAndRopSaveChangesMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open inbox folder
            uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopGetContentsTable to get the contents table of inbox folder before create message.
            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest()
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.None
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getContentsTableRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetContentsTableResponse getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getContentsTableResponse.ReturnValue, "Call RopGetContentsTable should success.");
            uint rowCount = getContentsTableResponse.RowCount;
            #endregion

            #region Call RopCreateMessage to create new not FAI Message object.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest()
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Create a message in INBOX which root is mailbox 
                AssociatedFlag = 0x00 // NOT an FAI message
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createMessageResponse.ReturnValue, "Call RopCreateMessage should success.");
            uint targetMessageHandle = this.ResponseSOHs[0][createMessageResponse.OutputHandleIndex];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R991, The HasMessageId field is {0}", createMessageResponse.HasMessageId);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R991
            bool isVerifiedR991 = createMessageResponse.HasMessageId == 0x00 && createMessageResponse.MessageId == null;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR991,
                991,
                @"[In RopCreateMessage ROP Response Buffer] [HasMessageId] The value 0x00 means this is the last byte in the buffer.");
            #endregion

            #region Call RopGetContentsTable to get the contents table of inbox folder before save message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getContentsTableRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createMessageResponse.ReturnValue, "Call RopGetContents should success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R339");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R339
            this.Site.CaptureRequirementIfAreEqual<uint>(
                rowCount,
                getContentsTableResponse.RowCount,
                339,
                @"[In Receiving a RopCreateMessage ROP Request] When processing the RopCreateMessage ROP ([MS-OXCROPS] section 2.2.6.2), the server MUST NOT commit the new Message object until it [server] receives a RopSaveChangesMessage ROP request ([MS-OXCROPS] section 2.2.6.3).");
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R728");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R728
            this.Site.CaptureRequirementIfAreEqual<int>(
                8,
                BitConverter.GetBytes(saveChangesMessageResponse.MessageId).Length,
                728,
                @"[In RopSaveChangesMessage ROP Response Buffer] Message Id: 8 bytes containing the MID ([MS-OXCDATA] section 2.2.1.2) for the saved Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R372, the MID in RopSaveChangesMessage Response is {0}.", saveChangesMessageResponse.MessageId);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R372
            // Because the R728 verify the RopSaveChangesMessage ROP response contains a MID and it's length is 8 bytes.
            // So R372 will be verified directly.
            this.Site.CaptureRequirement(
                372,
                @"[In Receiving a RopSaveChangesMessage ROP Request] The response contains the MID ([MS-OXCDATA] section 2.2.1.2) of the committed message.");

            #endregion

            #region Call RopGetContentsTable to get the contents table of inbox folder after save message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getContentsTableRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getContentsTableResponse = (RopGetContentsTableResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createMessageResponse.ReturnValue, "Call RopGetContents should success.");
            uint contentTableHandle = this.ResponseSOHs[0][getContentsTableResponse.OutputHandleIndex];

            #region Verify MS-OXCMSG_R700, MS-OXCMSG_R687
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R687");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R687
            // If the e-mail in the Inbox one more than before Call RopSaveChangeMessage, then R687 will be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                rowCount + 1,
                getContentsTableResponse.RowCount,
                687,
                @"[In RopCreateMessage ROP] The RopCreateMessage ROP ([MS-OXCROPS] section 2.2.6.2) is used to create a new Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R700");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R700
            this.Site.CaptureRequirementIfAreEqual<uint>(
                rowCount + 1,
                getContentsTableResponse.RowCount,
                700,
                @"[In RopSaveChangesMessage ROP] The RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) commits the changes made to the Message object.");
            #endregion
            #endregion

            #region Call RopSetColumns to sets the properties visible on contents table.
            PropertyTag[] propertyTags = new PropertyTag[1];

            // The PropertyTag of PidTagMid.
            propertyTags[0] = new PropertyTag(0x674A, (ushort)PropertyType.PtypInteger64);
            RopSetColumnsRequest setColumnsRequest = new RopSetColumnsRequest()
            {
                RopId = (byte)RopId.RopSetColumns,
                LogonId = CommonLogonId,

                // Set InputHandleIndex to 0x00, which specifies the location in the Server object handle table where the handle
                // for the input Server object is stored, as specified in [MS-OXCROPS].
                InputHandleIndex = CommonInputHandleIndex,
                SetColumnsFlags = (byte)AsynchronousFlags.None,
                PropertyTagCount = (ushort)propertyTags.Length,
                PropertyTags = propertyTags
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setColumnsRequest, contentTableHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetColumnsResponse setColumnsResponse = (RopSetColumnsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setColumnsResponse.ReturnValue, "Call RopSetColumns should success.");
            #endregion

            #region Call RopQueryRows to retrieve rows from contents table.
            RopQueryRowsRequest queryRowsRequest = new RopQueryRowsRequest()
            {
                RopId = (byte)RopId.RopQueryRows,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                QueryRowsFlags = (byte)QueryRowsFlags.Advance,
                ForwardRead = 0x01,
                RowCount = 0x1000
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(queryRowsRequest, contentTableHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopQueryRowsResponse queryRowsResponse = (RopQueryRowsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setColumnsResponse.ReturnValue, "Call RopQueryRowsRequest should success.");

            ulong messageID = 0;
            foreach (PropertyRow row in queryRowsResponse.RowData.PropertyRows)
            {
                ulong actualMID = BitConverter.ToUInt64(row.PropertyValues[0].Value, 0);

                if (actualMID == saveChangesMessageResponse.MessageId)
                {
                    messageID = actualMID;
                    break;
                }
            }

            Site.Assert.AreNotEqual<ulong>(messageID, 0, "The Message ID should in the contents table of the specified Folder object");
            #endregion

            #region Call RopOpenMessage to open the specific Message object.
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4],
                OpenModeFlags = 0x00,
                MessageId = messageID
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setColumnsResponse.ReturnValue, "Call RopOpenMessage should success.");

            #region Verify MS-OXCMSG_R647, MS-OXCMSG_R372
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R647");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R647
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                openMessageResponse.ReturnValue,
                647,
                @"[In RopOpenMessage ROP] The RopOpenMessage ROP ([MS-OXCROPS] section 2.2.6.1) provides access to an existing Message object, which is identified by the message ID (MID), whose structure is specified in [MS-OXCDATA] section 2.2.1.2.<7>");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R251");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R251
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                openMessageResponse.ReturnValue,
                251,
                @"[In Sending a RopOpenMessage ROP Request] The MID is accessible from the contents table of the Folder object that contains the Message object by including the PidTagMid property ([MS-OXCFXICS] section 2.2.1.2.1) in a RopSetColumns ROP request ([MS-OXCROPS] section 2.2.5.1), as specified in [MS-OXCTABL] section 2.2.2.2.");
            #endregion
            #endregion

            #region Call RopRelease to release all resources
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case validates the operation of creating a FAI message and saving the created message.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC02_CreateFAIMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new FAI Message object.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest()
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Create a message in INBOX which root is mailbox 
                AssociatedFlag = 0x01 // An FAI message.
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createMessageResponse.ReturnValue, "Call RopCreateMessage should success.");
            uint targetMessageHandle = this.ResponseSOHs[0][createMessageResponse.OutputHandleIndex];
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopGetPropertiesSpecific to get the value of PidTagMessageFlags property.
            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };
            List<PropertyObj> properties = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageFlags = PropertyHelper.GetPropertyByName(properties, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R988");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R988
            this.Site.CaptureRequirementIfAreEqual<int>(
                (int)MessageFlags.MfFAI,
                Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfFAI,
                988,
                @"[In RopCreateMessage ROP Request Buffer] [AssociatedFlag] Nonzero means the message to be created is an FAI message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R527");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R527
            this.Site.CaptureRequirementIfAreEqual<int>(
                (int)MessageFlags.MfFAI,
                Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfFAI,
                527,
                @"[In PidTagMessageFlags Property] [mfFAI (0x00000040)] The message is an FAI message.");
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created with SaveFlags is 0x05.
            RopSaveChangesMessageResponse saveChangesMessageResponseFirst = this.SaveMessage(targetMessageHandle, 0x05);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created with SaveFlags is 0x06.
            RopSaveChangesMessageResponse saveChangesMessageResponseSecond = this.SaveMessage(targetMessageHandle, 0x06);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1907");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1907
            this.Site.CaptureRequirementIfAreEqual<RopSaveChangesMessageResponse>(
                saveChangesMessageResponseFirst,
                saveChangesMessageResponseSecond,
                1907,
                @"[In RopSaveChangesMessage ROP Request Buffer] The server's responses are same with two different values [not 0x01, 0x02 or 0x04] of SaveFlags.");
            #endregion

            #region Call RopRelease to release all resources
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests properties to be initialized before committing the new Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC03_PropertiesInRopCreateMessageInitialization()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopGetPropertiesSpecific to get properties for created message before save message.
            // Prepare property Tag 
            PropertyTag[] tagArray = this.GetPropertyTagsForInitializeMessage();

            // Get properties for Created Message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse property response get Property Value to verify test  case requirement
            PropertyObj pidTagImportance = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagImportance);
            PropertyObj pidTagMessageClass = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageClass);
            PropertyObj pidTagSensitivity = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagSensitivity);
            PropertyObj pidTagDisplayBcc = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagDisplayBcc);
            PropertyObj pidTagDisplayCc = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagDisplayCc);
            PropertyObj pidTagDisplayTo = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagDisplayTo);
            PropertyObj pidTagHasAttachments = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagHasAttachments);
            PropertyObj pidTagTrustSender = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagTrustSender);
            PropertyObj pidTagAccess = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAccess);
            PropertyObj pidTagAccessLevel = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAccessLevel);
            PropertyObj pidTagCreationTime = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagCreationTime);
            PropertyObj pidTagLastModificationTime = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);
            PropertyObj pidTagSearchKey = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagSearchKey);
            PropertyObj pidTagCreatorName = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagCreatorName);
            PropertyObj pidTagLastModifierName = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModifierName);
            PropertyObj pidTagHasNamedProperties = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagHasNamedProperties);
            PropertyObj pidTagLocalCommitTime = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLocalCommitTime);
            PropertyObj pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            PropertyObj pidTagCreatorEntryId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagCreatorEntryId);
            PropertyObj pidTagLastModifierEntryId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModifierEntryId);
            PropertyObj pidTagMessageLocaleId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageLocaleId);
            PropertyObj pidTagLocaleId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLocaleId);


            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R987");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R987
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0x00000040,
                Convert.ToUInt32(pidTagMessageFlags.Value) & 0x00000040,
                987,
                @"[In RopCreateMessage ROP Request Buffer] [AssociatedFlag] Value 0x00 means the message to be created is not an FAI message.");

            int pidTagImportanceInitialValue = Convert.ToInt32(pidTagImportance.Value);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R341,the actual initial Data of PidTagImportance is {0}",
                pidTagImportanceInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R341
            Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                pidTagImportanceInitialValue,
                341,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial data of PidTagImportance is] 0x00000001.");

            string pidTagMessageClassInitialValue = Convert.ToString(pidTagMessageClass.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R342,the actual initial Data of PidTagMessageClass is {0}", pidTagMessageClassInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R342
            Site.CaptureRequirementIfAreEqual<string>(
                "IPM.Note",
                pidTagMessageClassInitialValue,
                342,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagMessageClass is] IPM.Note.");

            int pidTagSensitivityInitialValue = Convert.ToInt32(pidTagSensitivity.Value);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R343,the actual initial Data of PidTagSensitivity is {0}", pidTagSensitivityInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R343
            Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                pidTagSensitivityInitialValue,
                343,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagSensitivity is] 0x00000000.");

            string pidTagDisplayBccInitialValue = Convert.ToString(pidTagDisplayBcc.Value);

            if (Common.IsRequirementEnabled(344, this.Site))
            {
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R344,the actual initial Data of PidTagDisplayBcc is {0}",
                    pidTagDisplayBccInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R344
                Site.CaptureRequirementIfAreEqual<string>(
                    string.Empty,
                    pidTagDisplayBccInitialValue,
                    344,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagDisplayBcc is] """".");
            }

            if (Common.IsRequirementEnabled(345, this.Site))
            {
                string pidTagDisplayCcInitialValue = Convert.ToString(pidTagDisplayCc.Value);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R345,the actual initial Data of PidTagDisplayCc is {0}",
                    pidTagDisplayCcInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R345
                Site.CaptureRequirementIfAreEqual<string>(
                    string.Empty,
                    pidTagDisplayCcInitialValue,
                    345,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagDisplayCc is] """".");
            }

            if (Common.IsRequirementEnabled(346, this.Site))
            {
                string pidTagDisplayToInitialValue = Convert.ToString(pidTagDisplayTo.Value);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R346,the actual initial Data of PidTagDisplayTo is {0}",
                    pidTagDisplayToInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R346
                Site.CaptureRequirementIfAreEqual<string>(
                    string.Empty,
                    pidTagDisplayToInitialValue,
                    346,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagDisplayTo is] """".");
            }

            int pidTagHasAttachmentsInitialValue = Convert.ToInt32(pidTagHasAttachments.Value);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R349,the actual initial Data of PidTagHasAttachments is {0}",
                pidTagHasAttachmentsInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R349
            Site.CaptureRequirementIfAreEqual<int>(
                0x00,
                pidTagHasAttachmentsInitialValue,
                349,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagHasAttachments is] 0x00.");

            if (Common.IsRequirementEnabled(1713, this.Site))
            {
                int pidTagTrustSenderInitialValue = Convert.ToInt32(pidTagTrustSender.Value);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R352,the actual initial Data of PidTagTrustSender is {0}",
                    pidTagTrustSenderInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R352
                Site.CaptureRequirementIfAreEqual<int>(
                    0x00000001,
                    pidTagTrustSenderInitialValue,
                    352,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagTrustSender is] 0x00000001.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1713");

                // It indicates that Exchange 2007 supports the PidTagTrustSender property if code can run here.
                this.Site.CaptureRequirement(
                    1713,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagTrustSender property. (Exchange 2007 follows this behavior.)");
            }

            int pidTagAccessInitialValue = Convert.ToInt32(pidTagAccess.Value);
            if (Common.IsRequirementEnabled(1915, this.Site))
            {
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R1915,the actual initial Data of PidTagAccess is {0}",
                    pidTagAccessInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1915
                Site.CaptureRequirementIfAreEqual<int>(
                    0x00000003,
                    pidTagAccessInitialValue,
                    1915,
                    @"[In Appendix A: Product Behavior] Implementation does initialize the PidTagAccess property to 0x00000003. (Exchange 2007 and 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1914, this.Site))
            {
                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCMSG_R1914,the actual initial Data of PidTagAccess is {0}",
                    pidTagAccessInitialValue);

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1914
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000007,
                    pidTagAccessInitialValue,
                    1914,
                    @"[In Appendix A: Product Behavior] Implementation does initialize the PidTagAccess property to 0x00000007. (<23> Section 3.2.5.2: Exchange 2013 follows this behavior.)");
            }

            int pidTagAccessLevelInitialValue = Convert.ToInt32(pidTagAccessLevel.Value);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R354,the actual initial Data of PidTagAccessLevel is {0}",
                pidTagAccessLevelInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R354
            Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                pidTagAccessLevelInitialValue,
                354,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagAccessLevel is] 0x00000001.");

            DateTime pidTagCreationTimeInitialValue = Convert.ToDateTime(pidTagCreationTime.Value);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2034");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2034
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagCreationTimeInitialValue,
                2034,
                @"[In Receiving a RopCreateMessage ROP Request] PidTagCreationTime (section 2.2.2.3) [will be initialized when calling RopCreateMessage ROP].");

            if (Common.IsRequirementEnabled(1912, this.Site))
            {
                string pidTagCreatorNameInitialValue = Convert.ToString(pidTagCreatorName.Value).ToLower();
                string creatorName = Common.GetConfigurationPropertyValue("AdminUserName", this.Site).ToLower();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R360");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R360
                Site.CaptureRequirementIfAreEqual<string>(
                    creatorName,
                    pidTagCreatorNameInitialValue,
                    360,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagCreatorName is] Name of the creator.");

                bool isVerifyR357 = Convert.ToDateTime(pidTagLastModificationTime.Value) == pidTagCreationTimeInitialValue;

                // Above If condition has verified the initial data of PidTagLastModificationTime.
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR357,
                    357,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagLastModification is Same] as PidTagCreationTime property.");

                string pidTagLastModifierNameInitialValue = Convert.ToString(pidTagLastModifierName.Value).ToLower();

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R362");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R362
                Site.CaptureRequirementIfAreEqual<string>(
                    pidTagCreatorNameInitialValue,
                    pidTagLastModifierNameInitialValue,
                    362,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagLastModifierName is] Same as PidTagCreatorName property.");

                AddressBookEntryID creatorEntryId = new AddressBookEntryID();
                creatorEntryId.Deserialize((byte[])pidTagCreatorEntryId.Value, 0);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R361");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R361
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    creatorEntryId,
                    typeof(AddressBookEntryID),
                    361,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagCreatorEntryId is] Address Book EntryID of the creator.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1182");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1182
                this.Site.CaptureRequirementIfIsInstanceOfType(
                     creatorEntryId,
                     typeof(AddressBookEntryID),
                    1182,
                    @"[In PidTagCreatorEntryId Property] The PidTagCreatorEntryId property ([MS-OXPROPS] section 2.646) specifies the original author of the message according to their address book EntryID.");

                AddressBookEntryID modifierEntryId = new AddressBookEntryID();
                modifierEntryId.Deserialize((byte[])pidTagLastModifierEntryId.Value, 0);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R363");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R363
                this.Site.CaptureRequirementIfAreEqual<AddressBookEntryID>(
                    creatorEntryId,
                    modifierEntryId,
                    363,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagLastModifierEntryId is] Same as PidTagCreatorEntryId property.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1185");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1185
                this.Site.CaptureRequirementIfIsInstanceOfType(
                    modifierEntryId,
                    typeof(AddressBookEntryID),
                    1185,
                    @"[In PidTagLastModifierEntryId Property] The PidTagLastModifierEntryId property ([MS-OXPROPS] section 2.754) specifies the last user to modify the contents of the message according to their address book EntryID.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R359");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R359
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagMessageLocaleId.Value,
                    359,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagMessageLocaleId is] The Logon object LocaleID.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R365");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R365
                this.Site.CaptureRequirementIfAreEqual<object>(
                    pidTagMessageLocaleId.Value,
                    pidTagLocaleId.Value,
                    365,
                    @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagLocaleId is] Same as PidTagMessageLocaleId property.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1912");

                // MS-OXCMSG_R1912 can be verified if code can run here.
                this.Site.CaptureRequirement(
                    1912,
                    @"[In Appendix A: Product Behavior] Implementation does initialize the properties: PidTagCreatorName, PidTagCreatorEntryId, PidTagLastModifierName, PidTagLastModifierEntryId, PidTagLastModificationTime, PidTagMessageLocaleId  and PidTagLocaleId. (Exchange 2007 and Exchange 2010 follow this behavior.)");
            }

            // MS-OXCMSG_R358 can be verified if the value of PidTagSearchKey is not null since whether server generated SearchKey is unknown to client.
            Site.CaptureRequirementIfIsNotNull(
                pidTagSearchKey,
                358,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagSearchKey is] Server generated Search Key.");

            int pidTagHasNamedPropertiesInitialValue = Convert.ToInt32(pidTagHasNamedProperties.Value);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R364,the actual initial Data of PidTagHasNamedProperties is {0}",
                pidTagHasNamedPropertiesInitialValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R364
            Site.CaptureRequirementIfAreEqual<int>(
                0x00,
                pidTagHasNamedPropertiesInitialValue,
                364,
                @"[In Receiving a RopCreateMessage ROP Request] [The Initial Data of PidTagHasNamedProperties is] 0x00.");

            if(Common.IsRequirementEnabled(3015,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3015");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3015
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000009,
                    Convert.ToInt32(pidTagMessageFlags.Value),
                    3015,
                    @"[In Appendix A: Product Behavior] [The Initial data of PidTagMessageFlags] will be 0x00000009 (the mfEverRead flag combined by using the bitwise OR operation with the value 0x00000009) if the client does not explicitly set the read state. (<22> Section 3.2.5.2: Exchange 2007 follows this behavior.)");
            }

            if(Common.IsRequirementEnabled(3016,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3016");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3016
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000409,
                    Convert.ToInt32(pidTagMessageFlags.Value),
                    3016,
                    @"[In Appendix A: Product Behavior] [The Initial data of PidTagMessageFlags] will be 0x00000409 (the mfEverRead flag combined by using the bitwise OR operation with the value 0x00000009) if the client does not explicitly set the read state. (Exchange 2010 and above follow this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1140");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1140
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000400,
                    Convert.ToInt32(pidTagMessageFlags.Value)& 0x00000400,
                    1140,
                    @"[In PidTagMessageFlags Property] [mfEverRead (0x00000400)] The message has been read at least once.");
            }

            if(Common.IsRequirementEnabled(3006,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3006");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3006
                this.Site.CaptureRequirementIfIsTrue(
                    (Convert.ToInt32(pidTagMessageFlags.Value) & 0x00000400)== 0x00000400 && (Convert.ToInt32(pidTagMessageFlags.Value) & 0x00000001) == 0x00000001,
                    3006,
                    @"[In Appendix A: Product Behavior] [mfEverRead (0x00000400)] This flag is set by the implementation whenever the mfRead flag is set. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R510");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R510
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                Convert.ToUInt32(pidTagMessageFlags.Value) & 0x00000001,
                510,
                @"[In PidTagMessageFlags Property] [mfRead (0x00000001)] The message is marked as having been read.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R512");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R512
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000008,
                Convert.ToUInt32(pidTagMessageFlags.Value) & 0x00000008,
                512,
                @"[In PidTagMessageFlags Property] [mfUnsent (0x00000008)] The message is still being composed and is treated as a Draft Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R73");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R73
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                Convert.ToUInt32(pidTagImportance.Value),
                73,
                @"[In PidTagImportance Property] [The value 0x00000001 indicates the level of importance assigned by the end user to the Message object is] Normal importance.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R513");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R513
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000008,
                Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfUnsent,
                513,
                @"[In PidTagMessageFlags Property] [mfUnsent (0x00000008)] This bit is cleared by the server when responding to the RopSubmitMessage ROP ([MS-OXCROPS] section 2.2.7.1) with a success code.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R84");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R84
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                Convert.ToInt32(pidTagSensitivity.Value),
                84,
                @"[In PidTagSensitivity Property] [The value 0x00000000 indicates the sender's assessment of the sensitivity of the Message object is] Normal.");
            #endregion
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");
            #endregion

            #region Call RopGetPropertiesSpecific to get properties for created message after save message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            PropertyObj pidTagLocalCommitTimeAfterSave = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLocalCommitTime);
            #region Verify MS-OXCMSG_R1890
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1890, The value of PidTagLocalCommitTime is {0}", pidTagLocalCommitTimeAfterSave.Value);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1890
            bool isVerifiedR1890 = pidTagLocalCommitTime.Value != pidTagLocalCommitTimeAfterSave.Value;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1890,
                1890,
                @"[In Receiving a RopSaveChangesMessage ROP Request] The server sets the PidTagLocalCommitTime property (section 2.2.1.49) when the RopSaveChangesMessage ROP is processed.");
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests error code of RopCreateMessage and RopSaveChangesMessage.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC04_RopSaveChangesMessageWithInvalidSaveFlags()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage which contains an incorrect value of SaveFlags field.
            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
            {
                RopId = (byte)RopId.RopSaveChangesMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                SaveFlags = NotSupportedSaveFlags
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1052");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1052
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                saveChangesMessageResponse.ReturnValue,
                1052,
                @"[In Receiving a RopSaveChangesMessage ROP Request] [ecNotSupported (0x80040102)] The values of the SaveFlags are not a supported combination.");
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage to save read-only properties of message.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC05_SaveReadOnlyPropertiesOfMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyObj> propertyList;
            List<PropertyObj> ps;
            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageStatus]
            };
            PropertyObj pidTagMessageFlags;

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                #region Set PidTagCreationTime property.

                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagCreationTime, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1849");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1849
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1849,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only property PidTagCreationTime. (Exchange 2007 follows this behavior.)");
                #endregion

                #region Set PidTagLastModificationTime property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagLastModificationTime, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1850");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1850
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1850,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only property PidTagLastModificationTime. (Exchange 2007 follows this behavior.)");

                #endregion
            }

            #region Set PidTagHasAttachments property.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagHasAttachments, BitConverter.GetBytes(true))
            };
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1851");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1851
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1851,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only property PidTagHasAttachments. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1869");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1869
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1869,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only property PidTagHasAttachments. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R15");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R15
            // If According to above capture code, R15 will be verified.
            this.Site.CaptureRequirement(
                15,
                @"[In PidTagHasAttachments Property] This property [PidTagHasAttachments] is read-only for the client.");
            #endregion

            #region Set PidTagMessageSize property.
            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagMessageSize, BitConverter.GetBytes(0x00000002))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1865");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1865
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1865,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only property PidTagMessageSize. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1883");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1883
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1883,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only property PidTagMessageSize. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R47");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R47
            // If According to above capture code, R47 will be verified.
            this.Site.CaptureRequirement(
                47,
                @"[In PidTagMessageSize Property] This property [PidTagMessageSize] is read-only for the client.");
            #endregion

            #region Set MfRead flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            int messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfRead)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1852");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1852
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1852,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfRead. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1870");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1870
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1870,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfRead. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfUnsent flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfUnsent)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1853");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1853
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1853,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfUnsent. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1871");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1871
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1871,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfUnsent. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfResend flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfResend)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1854");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1854
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1854,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfResend. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1872");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1872
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1872,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfResend. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R31");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R31
            // If According to above capture code, R31 will be verified.
            this.Site.CaptureRequirement(
                31,
                @"[In PidTagMessageFlags Property] After the first successful call to the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3), as described in section 2.2.3.3, these flags [mfRead, mfUnsent, mfResend] are read-only for the client.");
            #endregion

            #region Set mfUnmodified flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfUnmodified)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1855");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1855
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1855,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfUnmodified. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1873");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1873
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1873,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfUnmodified. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfSubmitted flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfSubmitted)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1856");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1856
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1856,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfSubmitted. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1874");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1874
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1874,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfSubmitted. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfHasAttach flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfHasAttach)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1857");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1857
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1857,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfHasAttach. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1875");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1875
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1875,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfHasAttach. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfFromMe flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfFromMe)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1858");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1858
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1858,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfFromMe. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1876");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1876
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1876,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfFromMe. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfFAI flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfFAI)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1859");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1859
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1859,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfFAI. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1877");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1877
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1877,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfFAI. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfNotifyRead flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfNotifyRead)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1860");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1860
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1860,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfNotifyRead. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1878");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1878
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1878,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfNotifyRead. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfNotifyUnread flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfNotifyUnread)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1861");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1861
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1861,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfNotifyUnread. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1879");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1879
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1879,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfNotifyUnread. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfEverRead flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & 0x00000400)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1862");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1862
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1862,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfEverRead. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1880");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1880
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1880,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfEverRead. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfInternet flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfInternet)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1863");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1863
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1863,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfInternet. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1881");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1881
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1881,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfInternet. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion

            #region Set mfUntrusted flag of PidTagMessageFlags.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            propertyList = new List<PropertyObj>();
            messageFlags = (~(Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfUntrusted)) & Convert.ToInt32(pidTagMessageFlags.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1864");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1864
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1864,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag mfUntrusted. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1882");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1882
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1882,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag mfUntrusted. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R516");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R516
            // If According to above capture code, R516 will be verified.
            this.Site.CaptureRequirement(
                516,
                @"[In PidTagMessageFlags Property] These flags are always read-only for the client [mfUnmodified, mfSubmitted, mfHasAttach, mfFromMe, mfFAI, mfNotifyRead, mfNotifyUnread, mfEverRead, mfInternet, mfUntrusted].");
            #endregion

            #region Set msInConflict flag of PidTagMessageStatus.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageStatus = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageStatus);
            propertyList = new List<PropertyObj>();
            int messageStatus = (~(Convert.ToInt32(pidTagMessageStatus.Value) & (int)MessageStatusFlags.MsInConflict)) & Convert.ToInt32(pidTagMessageStatus.Value);
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageStatus, BitConverter.GetBytes(messageStatus)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);

            if (Common.IsRequirementEnabled(1112, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1866");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1866
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1866,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return an error if pending changes include changes to read-only flag msInConflict. (Exchange 2007 follows this behavior.)");
            }

            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                RopSaveChangesMessageResponse saveMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1884");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1884
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    saveMessageResponse.ReturnValue,
                    1884,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only flag msInConflict. (Exchange 2010 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R545");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R545
            // If According to above capture code, R545 will be verified.
            this.Site.CaptureRequirement(
                545,
                @"[In PidTagMessageStatus Property] [msInConflict (0x00000800)] This is a read-only value for the client.");
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage with read-only SaveFlag value.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC06_RopSaveChangeMessageWithReadOnly()
        {
            this.CheckMapiHttpIsSupported();
            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel]
            };
            List<PropertyObj> propertyValues;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageID = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopGetPropertiesSpecific to get PidTagAccessLevel property for created message before save message.
            // Prepare property Tag 
            PropertyTag[] tagArray = new PropertyTag[1];
            tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];

            // Get properties for Created Message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);
            #endregion

            #region Call RopSaveChangesMessage to save a Message object and server return an error.
            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                // Call RopSetProperties to set PidTagHasAttachments that is a read-only property.
                // The server will return a GeneralFailure error when call RopSaveChangesMessage if property R1644Enabled is true in SHOULD/MAY ptfconfig file.
                List<PropertyObj> propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagHasAttachments, BitConverter.GetBytes(true))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

                PropertyObj accesssLevelAfterReadOnlyFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterReadOnlyFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R1670 will be verified.
                this.Site.CaptureRequirement(
                    1670,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [KeepOpenReadOnly (0x01)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");
            }

            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created and keep the message open with read-only.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadOnly);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelAfterReadOnlySuccess = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

            if (Common.IsRequirementEnabled(2192, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2192");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R2192
                // Because the server has returned a success code when call RopSaveChangesMessage.
                // If the PidTagAccessLevel property does not change then R2192 will be verified.
                this.Site.CaptureRequirementIfAreEqual<int>(
                    Convert.ToInt32(accesssLevelBeforeSave.Value),
                    Convert.ToInt32(accesssLevelAfterReadOnlySuccess.Value),
                    2192,
                    @"[In Appendix B: Product Behavior] Implementation returns a success code and keeps the Message object open with read-only access.(<13> Section 2.2.3.3.1: Exchange 2010, Exchange 2013, and Exchange 2016 ignore the KeepOpenReadOnly flag.)");
            }

            if (Common.IsRequirementEnabled(2193, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2193");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R2193
                // Because the server has returned a success code when call RopSaveChangesMessage.
                // If the PidTagAccessLevel property is 0x00000000 (Read-only) then R2193 will be verified.
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000000,
                    Convert.ToInt32(accesssLevelAfterReadOnlySuccess.Value),
                    2193,
                    @"[In Appendix B: Product Behavior] Implementation returns a success code and keeps the Message object open with read-only access. (Exchange 2007 follow this behavior.)");
            }
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage with read-write SaveFlag value.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC07_RopSaveChangeMessageWithReadWrite()
        {
            this.CheckMapiHttpIsSupported();
            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel]
            };
            List<PropertyObj> propertyValues;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageID = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopGetPropertiesSpecific to get properties for created message before save message.
            // Prepare property Tag 
            PropertyTag[] tagArray = new PropertyTag[1];
            tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];

            // Get properties for Created Message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);
            #endregion

            #region Call RopSaveChangesMessage to save a Message object and server return an error.
            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                // Call RopSetProperties to set PidTagHasAttachments that is a read-only property.
                // The server will return a GeneralFailure error when call RopSaveChangesMessage if property R1644Enabled is true in SHOULD/MAY ptfconfig file.
                List<PropertyObj> propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagHasAttachments, BitConverter.GetBytes(true))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadWrite
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

                PropertyObj accesssLevelAfterReadWriteFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterReadWriteFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R718 will be verified.
                this.Site.CaptureRequirement(
                    718,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [KeepOpenReadWrite (0x02)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");

                this.ReleaseRop(targetMessageHandle);
            }
            #endregion
            
            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created and keep the message open with read-write.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelAfterReadWriteSuccess = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R719");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R719
            // Because the server has returned a success code when call RopSaveChangesMessage.
            // If the PidTagAccessLevel property is 0x00000001 (Modify) then R719 will be verified.
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                Convert.ToInt32(accesssLevelAfterReadWriteSuccess.Value),
                719,
                @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [KeepOpenReadWrite (0x02)] [If the RopSaveChangesMessage ROP succeeded] The server returns a success code and keeps the Message object open with read/write access.");
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage with ForceSave SaveFlag value.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC08_RopSaveChangeMessageWithForceSave()
        {
            this.CheckMapiHttpIsSupported();
            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel]
            };
            List<PropertyObj> propertyValues;
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageID = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopGetPropertiesSpecific to get properties for created message before save message.
            // Prepare property Tag 
            PropertyTag[] tagArray = new PropertyTag[1];
            tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];

            // Get properties for Created Message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);
            #endregion

            #region Call RopSaveChangesMessage to save a Message object and server return an error.
            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                // Call RopSetProperties to set PidTagHasAttachments that is a read-only property.
                // The server will return a GeneralFailure error when call RopSaveChangesMessage if property R1644Enabled is true in SHOULD/MAY ptfconfig file.
                List<PropertyObj> propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagHasAttachments, BitConverter.GetBytes(true))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = InvalidInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.ForceSave
                };
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
                PropertyObj accesssLevelAfterForceSaveFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterForceSaveFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R722 will be verified.
                this.Site.CaptureRequirement(
                    722,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [ForceSave (0x04)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");

                this.ReleaseRop(targetMessageHandle);
            }
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created and keep the message open with ForceSave.
            if (Common.IsRequirementEnabled(3011, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, 0x0C);
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3011");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3011
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    3011,
                    @"[In Appendix A: Product Behavior] <14> Section 2.2.3.3.1:  The value of ForceSave is 0x0C in Microsoft Exchange Server 2007 Service Pack 3 (SP3).");
            }

            if (Common.IsRequirementEnabled(3022,this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, 0x04);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3022");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3022
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    3022,
                    @"[In Appendix A: Product Behavior] The value of ForceSave is 0x04. (Exchange 2010 and above follow this behavior.)");
            }

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj accesssLevelAfterForceSaveSuccess = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R723");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R723
            // Because the server has returned a success code when call RopSaveChangesMessage.
            // If the PidTagAccessLevel property is 0x00000001 (Modify) then R719 will be verified.
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                Convert.ToInt32(accesssLevelAfterForceSaveSuccess.Value),
                723,
                @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [ForceSave (0x04)] [If the RopSaveChangesMessage ROP succeeded] The server returns a success code and keeps the Message object open with read/write access.");
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage failure when save Message object in different Transaction.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC09_RopSaveChangeMessageFailure()
        {
            if (Common.IsRequirementEnabled(1916, this.Site))
            {
                this.CheckMapiHttpIsSupported();
                this.ConnectToServer(ConnectionType.PrivateMailboxServer);

                List<PropertyTag> propertyTags = new List<PropertyTag>
                {
                    PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel]
                };
                List<PropertyObj> propertyValues;
                RopSaveChangesMessageResponse saveChangesMessageResponse;

                #region Call RopLogon to log on a private mailbox.
                RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
                #endregion

                #region Call RopCreateMessage to create new Message object.
                uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
                ulong messageID = saveChangesMessageResponse.MessageId;
                this.ReleaseRop(targetMessageHandle);
                #endregion

                #region Call RopOpenMessage to open the message object created by step above.
                targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
                #endregion

                #region Call RopGetPropertiesSpecific to get PidTagAccessLevel property for created message before save message.
                // Prepare property Tag 
                PropertyTag[] tagArray = new PropertyTag[1];
                tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];

                // Get properties for Created Message
                RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
                {
                    RopId = (byte)RopId.RopGetPropertiesSpecific,
                    LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                    InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                    PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                    PropertyTagCount = (ushort)tagArray.Length,
                    PropertyTags = tagArray
                };
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
                PropertyObj accesssLevelBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);
                #endregion

                #region Call RopSaveChangesMessage to save the created Message object on one transaction.
                // Call RopSaveChangesMessage to save the created message on one transaction.
                // The server will return an error when call RopSaveChangesMessage if property R1916Enabled is true in SHOULD/MAY ptfconfig file.
                uint messageHandleSecond = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
                saveChangesMessageResponse = this.SaveMessage(messageHandleSecond, (byte)SaveFlags.ForceSave);
                this.ReleaseRop(messageHandleSecond);
                #endregion

                #region Call RopSaveChangesMessage and Saveflag field is KeepOpenReadOnly.
                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

                PropertyObj accesssLevelAfterReadOnlyFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterReadOnlyFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R1670 will be verified.
                this.Site.CaptureRequirement(
                    1670,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [KeepOpenReadOnly (0x01)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");
                #endregion

                #region Call RopSaveChangesMessage and SaveFlag field is KeepOpenReadWrite.
                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadWrite
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

                PropertyObj accesssLevelAfterReadWriteFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterReadWriteFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R718 will be verified.
                this.Site.CaptureRequirement(
                    718,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [KeepOpenReadWrite (0x02)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");
                #endregion

                #region Call RopSaveChangesMessage and SaveFlag field is ForceSave.
                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = InvalidInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.ForceSave
                };
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
                PropertyObj accesssLevelAfterForceSaveFail = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAccessLevel);

                Site.Assert.AreEqual<int>(Convert.ToInt32(accesssLevelBeforeSave.Value), Convert.ToInt32(accesssLevelAfterForceSaveFail.Value), "The Message object access level should be unchanged.");

                // Because server return error code when call RopSaveChangesMessage and 
                // the access level has not been changed between call RopSaveChangesMessage before and after.
                // So R722 will be verified.
                this.Site.CaptureRequirement(
                    722,
                    @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [ForceSave (0x04)] [If the RopSaveChangesMessage ROP failed] The server returns an error and leaves the Message object open with unchanged access level.");
                #endregion

                #region Call RopRelease to release created message.
                this.ReleaseRop(targetMessageHandle);
                #endregion
            }
            else
            {
                this.isNotNeedCleanupPrivateMailbox = true;
                Site.Assume.Inconclusive("This case runs only if the implementation does not return Success for RopSaveChangesMessage ROP requests when a previous request has already been committed against the Message object, even though the changes to the object are not actually committed to the server store.");
            }
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage failure when the message has been opened or previously saved as read only;.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC10_RopSaveChangeMessageWithReadOnlyProperty()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            RopSaveChangesMessageResponse saveChangesMessageResponse;

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageID = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above as read-only.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadOnly);
            #endregion

            #region Call RopSaveChangesMessage to commit the Message object created and keep the message open with read-only.
           
            if(Common.IsRequirementEnabled(3019,this.Site))
            {
                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId, // The logonId 0x00 is associated with this operation.
                    InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response. 
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };
                uint returnValue = 0;
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None,out returnValue);

                bool isR3019Verifed = false;

                if (returnValue == 0)
                {
                    if (((RopSaveChangesMessageResponse)this.response).ReturnValue == 0x80004005)
                    {
                        isR3019Verifed = true;
                    }
                }
                else
                {
                    if (returnValue == 0x80004005)
                    {
                        isR3019Verifed = true;
                    }
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3019");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3019
                this.Site.CaptureRequirementIfIsTrue(
                    isR3019Verifed,
                    3019,
                    @"[In Appendix A: Product Behavior] [ecError (0x80004005)] The message has been opened or previously saved as read only; changes cannot be saved. (<27> Section 3.2.5.3: Exchange 2010, Exchange 2013, and Exchange 2016 follow this behavior.)");

            }
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage failure when the message has been opened or previously saved as read only;.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC11_CreateMessageWithoutPermissions()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            string commonUser = Common.GetConfigurationPropertyValue("CommonUser", Site);
            string commonUserPassword = Common.GetConfigurationPropertyValue("CommonUserPassword", Site);
            string commonUserEssdn = Common.GetConfigurationPropertyValue("CommonUserEssdn", Site);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open inbox folder
            uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Add Read permission to "CommonUser" on inbox folder.
            // Add folder visible permission for the inbox.
            uint pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible | (uint)PidTagMemberRights.ReadAny;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, openedFolderHandle);
            #endregion

            #region Call RopLogon to logon the private mailbox with "CommonUser"
            this.rawData = null;
            this.insideObjHandle = 0;
            this.response = null;
            this.ResponseSOHs = null;
            this.MSOXCMSGAdapter.RpcDisconnect();
            this.MSOXCMSGAdapter.Reset();
            this.MSOXCMSGAdapter.RpcConnect(ConnectionType.PrivateMailboxServer, commonUser, commonUserPassword, commonUserEssdn);

            string userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site) + "\0";
            RopLogonRequest logonRequest = new RopLogonRequest()
            {
                RopId = (byte)RopId.RopLogon,
                LogonId = CommonLogonId,
                OutputHandleIndex = 0x00, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                StoreState = 0,
                LogonFlags = 0x01, // Logon to a private mailbox
                OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping, // Requesting admin access to the mail box
                EssdnSize = (ushort)Encoding.ASCII.GetByteCount(userDN),
                Essdn = Encoding.ASCII.GetBytes(userDN)
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(logonRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, logonResponse.ReturnValue, "Call RopLogon should success.");

            uint objHandle = this.ResponseSOHs[0][logonResponse.OutputHandleIndex];
            #endregion

            #region Call RopCreateMessage to create new Message object in folder without permission.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest()
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Create a message in INBOX which root is mailbox 
                AssociatedFlag = 0x00 // NOT an FAI message
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.response;
            
            if(Common.IsRequirementEnabled(3017,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3017");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3017
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004FF,
                    createMessageResponse.ReturnValue,
                    3017,
                    @"[In Appendix A: Product Behavior] [ecNoCreateRight (0x000004FF)] The user does not have permissions to create this message [RopCreateMessage]. (<24> Section 3.2.5.2:  Exchange 2007 and Exchange 2010 follow this behavior.)");
            }

            if(Common.IsRequirementEnabled(3018,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3018");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3018
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    createMessageResponse.ReturnValue,
                    3018,
                    @"[In Appendix A: Product Behavior] [ecAccessDenied (0x80070005)] The user does not have permissions to create this message [RopCreateMessage]. (Exchange 2013 and above follow this behavior.)");
            }
            #endregion
        }

        /// <summary>
        /// This test case tests calling RopSaveChangesMessage failure when the Message object include read-only properties.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S01_TC12_RopSaveChangeMessageWithReadOnlyProperty()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            RopSaveChangesMessageResponse saveChangesMessageResponse;

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create new Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageID = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenMessage to open the message object created by step above.
            targetMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageID, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSaveChangesMessage to save a Message object and server return an error.
            if (Common.IsRequirementEnabled(1644, this.Site))
            {
                #region Call RopSetProperties to set PidTagHasAttachments that is a read-only property.
                List<PropertyObj> propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagHasAttachments, BitConverter.GetBytes(true))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagMessageSize that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagMessageSize, BitConverter.GetBytes(0))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagAccess that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagAccess, BitConverter.GetBytes(0x00000001))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagAccessLevel that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagAccessLevel, BitConverter.GetBytes(0x00000001))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagObjectType that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagObjectType, BitConverter.GetBytes(0x00000005))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagRecordKey that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagRecordKey, new byte[] { 0x10, 0x00, 0xa1, 0x3e, 0x25, 0x6f, 0xbb, 0x2c, 0x26, 0x44, 0x8c, 0x8a, 0x2c, 0x52, 0x0c, 0x85, 0xfb, 0x08})
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagMessageStatus that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagMessageStatus, BitConverter.GetBytes(0x00001000))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagChangeKey that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagChangeKey,new byte[] { 0x10, 0x00, 0xa1, 0x3e, 0x25, 0x6f, 0xbb, 0x2c, 0x26, 0x44, 0x8c, 0x8a, 0x2c, 0x52, 0x0c, 0x85, 0xfb, 0x08})
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagSearchKey that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagSearchKey,new byte[] { 0x10, 0x00, 0xa1, 0x3e, 0x25, 0x6f, 0xbb, 0x2c, 0x26, 0x44, 0x8c, 0x8a, 0x2c, 0x52, 0x0c, 0x85, 0xfb, 0x08})
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagCreationTime that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagCreationTime, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagLastModificationTime that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagLastModificationTime, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                #region Call RopSetProperties to set PidTagLastModifierName that is a read-only property.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagLastModifierName,Common.GetBytesFromUnicodeString("Last modifier name"))
                };
                this.SetPropertiesForMessage(targetMessageHandle, propertyList);

                saveChangesMessageRequest = new RopSaveChangesMessageRequest()
                {
                    RopId = (byte)RopId.RopSaveChangesMessage,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                    SaveFlags = (byte)SaveFlags.KeepOpenReadOnly
                };

                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesMessageRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
                saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.response;

                Site.Assert.AreEqual<uint>(0x80004005, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should failed.");
                #endregion

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1724");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1724
                this.Site.CaptureRequirement(
                    1724,
                    @"[In General Properties] These properties are read-only for the client: PidTagAccess, PidTagChangeKey, PidTagCreationTime, PidTagLastModificationTime, PidTagLastModifierName, PidTagObjectType, PidTagRecordKey and PidTagSearchKey.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1078");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1078
                this.Site.CaptureRequirement(
                    1078,
                    @"[In General Properties] These properties [PidTagAccessLevel, PidTagObjectType and PidTagRecordKey] are read-only for the client.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1644");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1644
                this.Site.CaptureRequirement(
                    1644,
                    @"[In Appendix A: Product Behavior] Implementation does return a GeneralFailure error if pending changes include changes to read-only properties PidTagMessageSize, PidTagAccess, PidTagAccessLevel, PidTagObjectType, PidTagRecordKey, PidTagMessageStatus, and PidTagHasAttachments [about RopSaveChangeMessage]. (Exchange 2010 and above follow this behavior.)");
            }

            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        #region Private methods
        /// <summary>
        /// Create property tags for Created Message 
        /// </summary>
        /// <returns>PropertyTag array</returns>
        private PropertyTag[] GetPropertyTagsForInitializeMessage()
        {
            PropertyTag[] tags = new PropertyTag[23];
            tags[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagImportance];
            tags[1] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageClass];
            tags[2] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagSensitivity];
            tags[3] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagDisplayBcc];
            tags[4] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagDisplayCc];
            tags[5] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagDisplayTo];
            tags[6] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags];
            tags[7] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageSize];
            tags[8] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasAttachments];
            tags[9] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagTrustSender];
            tags[10] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccess];
            tags[11] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagAccessLevel];
            tags[12] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreationTime];
            tags[13] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime];
            tags[14] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagSearchKey];
            tags[15] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageLocaleId];
            tags[16] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreatorName];
            tags[17] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreatorEntryId];
            tags[18] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModifierName];
            tags[19] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModifierEntryId];
            tags[20] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasNamedProperties];
            tags[21] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLocaleId];
            tags[22] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagLocalCommitTime];

            return tags;
        }
        #endregion
    }
}