namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that Verifies the requirements related to message properties.
    /// </summary>
    [TestClass]
    public class S04_MessageObjectProperties : TestSuiteBase
    {
        #region Const definitions for test
        /// <summary>
        /// Constant string for the test data of PidTagSubjectPrefix.
        /// </summary>
        private const string TestDataOfPidTagSubjectPrefix = "PidTagSubjectPrefix for test";

        /// <summary>
        /// Constant string for the test data of PidTagAutoForwardComment.
        /// </summary>
        private const string TestDataOfPidTagAutoForwardComment = "PidTagAutoForwardComment for test";

        /// <summary>
        /// Constant string for the test data of PidTagBody.
        /// </summary>
        private const string TestDataOfPidTagBody = "the body for test";

        /// <summary>
        /// Constant string for the test data of PidTagBodyHtml.
        /// </summary>
        private const string TestDataOfPidTagBodyHtml = "<Html><head>the head</head><body> the body for test </body></Html>";

        /// <summary>
        /// Constant string for the test data of PidTagAddressBookDisplayNamePrintable.
        /// </summary>
        private const string TestDataOfPidTagAddressBookDisplayNamePrintable = "AddressBookDisplayNamePrintable";

        /// <summary>
        /// Constant string for the test data of DateTime.
        /// </summary>
        private const string TestDataOfDateTime = "2010-8-8 19:05:59";

        /// <summary>
        /// Constant string for the test data of PidLidInfoPathFormName.
        /// </summary>
        private const string TestDataOfInfoPath = "InfoPathForm.ac0c89fc41be28b5$e72d1d7069579cdb.f";

        /// <summary>
        /// Constant string for the test data of PidTagBodyContentId.
        /// </summary>
        private const string TestDataOfPidTagBodyContentId = "ebdc14bd-deb4-4816-b00b-6e2a46097d17";

        /// <summary>
        /// Constant string for the test data of PidTagBodyContentLocation.
        /// </summary>
        private const string TestDataOfPidTagBodyContentLocation = "http://i3.asp.net/asp.net/images/people/";

        /// <summary>
        /// Constant string for the test data of PidNameContentBase.
        /// </summary>
        private const string TestDataOfPidNameContentBase = "http://i3.asp.net/asp.net/images/people/";

        /// <summary>
        /// Constant string for the test data of PidNameAcceptLanguage.
        /// </summary>
        private const string TestDataOfPidNameAcceptLanguage = "en-US";

        /// <summary>
        /// Constant string for the test data of PidNameKeywords.
        /// </summary>
        private const string TestDataOfPidNameKeywords = "Simple Mail";

        /// <summary>
        /// Constant string for the test data of PidNameContentClass.
        /// </summary>
        private const string TestDataOfPidNameContentClass = "voice";

        /// <summary>
        /// Constant string for the test data of PidNameContentType.
        /// </summary>
        private const string TestDataOfPidNameContentType = "application/ms-tnef";

        /// <summary>
        /// Constant string for the test data of PidTagInternetReferences.
        /// </summary>
        private const string TestDataOfPidTagInternetReferences = "a375a61600000001, a375a61600000002";

        /// <summary>
        /// The priority of message is un-urgent
        /// </summary>
        private static byte[] resultUnurgent = new byte[] { 0xFF, 0xFF, 0xFF, 0xFF };

        /// <summary>
        /// The body is the rtf text. Uncompressed is "{\ref1 WXYZWXYZWXYZWXYZWXYZ}"
        /// </summary>
        private byte[] rtfCompressedText = new byte[] { 0x1a, 0x00, 0x00, 0x00, 0x1c, 0x00, 0x00, 0x00, 0x4c, 0x5a, 0x46, 0x75, 0xe2, 0xd4, 0x4b, 0x51, 0x41, 0x00, 0x04, 0x20, 0x57, 0x58, 0x59, 0x5a, 0x0d, 0x6e, 0x7d, 0x01, 0x0e, 0xb0 };
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
        /// This test case tests general properties on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC01_GeneralPropertiesOnMessageObject()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopModifyRecipients to add recipient to message created by step2
            PropertyTag[] recipientColumns = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(Common.GetConfigurationPropertyValue("AdminUserName", this.Site), 0)
            };
            this.AddRecipients(modifyRecipientRow, targetMessageHandle, recipientColumns);
            #endregion

            #region Call RopGetPropertiesSpecific to get PidTagMessageFlags property for created message before save message.
            // Prepare property Tag 
            PropertyTag[] tagArray = new PropertyTag[1];
            tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags];

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

            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            PropertyObj pidTagMessageFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageFlags);
            #endregion

            #region Call RopSetProperties to set PidTagMessageFlags property of created message.
            List<PropertyObj> propertyList = this.SetGeneralPropertiesOfMessage();
            int messageFlags = Convert.ToInt32(pidTagMessageFlags.Value) | (int)MessageFlags.MfResend;
            propertyList.Add(new PropertyObj(PropertyNames.PidTagMessageFlags, BitConverter.GetBytes(messageFlags)));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);

            // The message ID of specific message created by above step.
            ulong messageId = saveChangesMessageResponse.MessageId;
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
            #endregion

            #region Call RopOpenMessage to open created message.
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF,
                FolderId = logonResponse.FolderIds[4],
                OpenModeFlags = (byte)MessageOpenModeFlags.ReadOnly,
                MessageId = messageId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, "Call RopOpenMessage should success.");
            targetMessageHandle = this.ResponseSOHs[0][openMessageResponse.OutputHandleIndex];

            #region Verify MS-OXCMSG_R676, MS-OXCMSG_R678 and MS-OXCMSG_R1298
            string subjectPrefixInOpenResponse = System.Text.ASCIIEncoding.ASCII.GetString(openMessageResponse.SubjectPrefix.String);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R676");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R676
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagSubjectPrefix,
                subjectPrefixInOpenResponse.Substring(0, subjectPrefixInOpenResponse.Length - 1),
                676,
                @"[In RopOpenMessage ROP Response Buffer] [SubjectPrefix] The SubjectPrefix field contains the value of the PidTagSubjectPrefix property (section 2.2.1.9).");

            string normalizedSubjectInOpenResponse = System.Text.ASCIIEncoding.ASCII.GetString(openMessageResponse.NormalizedSubject.String);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R678");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R678
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteBase.TestDataOfPidTagNormalizedSubject,
                normalizedSubjectInOpenResponse.Substring(0, normalizedSubjectInOpenResponse.Length - 1),
                678,
                @"[In RopOpenMessage ROP Response Buffer] [NormalizedSubject] The NormalizedSubject field contains the value of the PidTagNormalizedSubject property (section 2.2.1.10).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1298");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1298
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x10,
                openMessageResponse.RecipientRows[0].RecipientType,
                1298,
                @"[In RopOpenMessage ROP Response Buffer] [RecipientRows] The value 0x10 means when resending a previous failure, this flag indicates that this recipient (1) did not successfully receive the message on the previous attempt.");

            #endregion
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set PropertySizeLimit,which specifies the maximum size allowed for a property value returned,
                // as specified in [MS-OXCROPS] section 2.2.8.4.1.
                PropertySizeLimit = 0xFFFF,
                WantUnicode = 1,
            };

            // In process, call capture code to verify adapter requirement
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            RopGetPropertiesAllResponse getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetPropertiesAll should success.");
            List<PropertyObj> propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);

            PropertyObj pidTagImportance = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagImportance);
            PropertyObj pidTagPriority = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagPriority);
            PropertyObj pidTagSensitivity = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);
            PropertyObj pidTagTrustSender = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagTrustSender);
            PropertyObj pidTagSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubject);
            PropertyObj pidTagSubjectPrefix = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubjectPrefix);
            PropertyObj pidTagNormalizedSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNormalizedSubject);

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1668 ,the property count is {0}.", getPropertiesAllResponse.PropertyValueCount);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1668
            bool isVerifiedR1668 = getPropertiesAllResponse.PropertyValueCount > 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1668,
                1668,
                @"[In RopOpenMessage ROP Response Buffer] [HasNamedProperties] Nonzero: Named properties are defined for this Message object and can be obtained through a RopGetPropertiesAll ROP request ([MS-OXCROPS] section 2.2.8.4).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R72");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R72
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                Convert.ToInt32(pidTagImportance.Value),
                72,
                @"[In PidTagImportance Property] [The value 0x00000000 indicates the level of importance assigned by the end user to the Message object is] Low importance.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R80");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R80
            this.Site.CaptureRequirementIfAreEqual<int>(
                -1,
                Convert.ToInt32(pidTagPriority.Value),
                80,
                @"[In PidTagPriority Property] [The value 0xFFFFFFFF indicates  the client's request for the priority is] Not urgent.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R86");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R86
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                Convert.ToInt32(pidTagSensitivity.Value),
                86,
                @"[In PidTagSensitivity Property] [The value 0x00000001 indicates the sender's assessment of the sensitivity of the Message object is] Personal.");

            if (Common.IsRequirementEnabled(1713, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1236");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1236
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000001,
                    (int)pidTagTrustSender.Value,
                    1236,
                    @"[In PidTagTrustSender] The value 0x00000001 indicates that the message was delivered through a trusted transport channel.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1713");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1713
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagTrustSender.Value,
                    1713,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagTrustSender property. (Exchange 2007 follows this behavior.)");
            }

            PropertyObj pidTagPurportedSenderDomain = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagPurportedSenderDomain);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1214");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1214
            this.Site.CaptureRequirementIfAreEqual<string>(
                Common.GetConfigurationPropertyValue("Domain", this.Site),
                pidTagPurportedSenderDomain.Value.ToString(),
                1214,
                @"[In PidTagPurportedSenderDomain Property] The PidTagPurportedSenderDomain property ([MS-OXPROPS] section 2.865) contains the domain name of the last sender responsible for transmitting the current message.");

            PropertyObj pidTagAlternateRecipientAllowed = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAlternateRecipientAllowed);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1797");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1797
            this.Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(pidTagAlternateRecipientAllowed.Value),
                1797,
                @"[In PidTagAlternateRecipientAllowed Property] This property [PidTagAlternateRecipientAllowed] is set to ""TRUE"" if autoforwarding is allowed.");

            PropertyObj pidTagResponsibility = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagResponsibility);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1198");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1198
            this.Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(pidTagResponsibility.Value),
                1198,
                @"[In PidTagResponsibility Property] This property [PidTagResponsibility] is set to ""TRUE"" if another agent has accepted responsibility.");

            pidTagMessageFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R515");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R515
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000080,
                Convert.ToInt32(pidTagMessageFlags.Value) & 0x00000080,
                515,
                @"[In PidTagMessageFlags Property] [mfResend (0x00000080)] The message includes a request for a resend operation with a non-delivery report.");
            #endregion
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            List<PropertyTag> propertiesOfMessage = this.GetPropertyTagListOfMessageObject();
            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertiesOfMessage);

            // Parse property response get Property Value to verify test  case requirement
            pidTagImportance = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagImportance);
            pidTagSubjectPrefix = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubjectPrefix);
            pidTagNormalizedSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNormalizedSubject);
            pidTagSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubject);
            PropertyObj pidTagRecipientDisplayName = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRecipientDisplayName);

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R58");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R58
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidTagSubjectPrefix,
                pidTagSubjectPrefix.Value.ToString(),
                58,
                @"[In PidTagSubjectPrefix Property] The PidTagSubjectPrefix property ([MS-OXPROPS] section 2.1096) contains the prefix for the subject of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R64");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R64
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestSuiteBase.TestDataOfPidTagNormalizedSubject,
                pidTagNormalizedSubject.Value.ToString(),
                64,
                @"[In PidTagNormalizedSubject Property] The PidTagNormalizedSubject property ([MS-OXPROPS] section 2.877) contains the normalized subject of the message, as specified in [MS-OXCMAIL] section 2.2.3.2.6.1.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1238");

            string mailSubject = TestDataOfPidTagSubjectPrefix + TestDataOfPidTagNormalizedSubject;

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1238
            this.Site.CaptureRequirementIfAreEqual<string>(
                mailSubject,
                pidTagSubject.Value.ToString(),
                1238,
                @"[In PidTagSubject Property] The PidTagSubject property ([MS-OXPROPS] section 2.1021) contains the full subject of an e-mail message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1239");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1239
            this.Site.CaptureRequirementIfAreEqual<string>(
                mailSubject,
                pidTagSubject.Value.ToString(),
                1239,
                @"[In PidTagSubject Property] The full subject is a concatenation of the subject prefix, as identified by the PidTagSubjectPrefix property (section 2.2.1.9), and the normalized subject, as identified by the PidTagNormalizedSubject property (section 2.2.1.10).");

            PropertyObj pidTagAutoForwarded = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagAutoForwarded);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1108");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1108
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                Convert.ToByte(pidTagAutoForwarded.Value),
                1108,
                @"[In PidTagAutoForwarded Property] If this property [PidTagAutoForwarded] is unset, a default value of 0x00 is assumed.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2047");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2047
            this.Site.CaptureRequirementIfAreEqual<string>(
                Common.GetConfigurationPropertyValue("AdminUserName", this.Site),
                pidTagRecipientDisplayName.Value.ToString(),
                2047,
                @"[In PidTagRecipientDisplayName Property] The PidTagRecipientDisplayName property ([MS-OXPROPS] section 2.888) specifies the display name of a recipient (2).");
            #endregion
            #endregion

            #region Receive the message sent by step 9.
            bool isMessageReceived = WaitEmailBeDelivered(folderHandle, rowCountBeforeSubmit);
            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion

            #region Call RopRelease to release all resources.
            this.ReleaseRop(targetMessageHandle);
            this.ReleaseRop(folderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests body properties on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC02_BodyPropertiesOnMessageObject()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            bool isMessageReceived = false;
            List<PropertyTag> propertiesOfMessage = new List<PropertyTag>();

            List<PropertyObj> propertyValues;

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call ROPs to send a message that contains plain text body to inbox.
            // Initialize message information, and set the handle for this message.
            // Create a message in Outbox folder.
            uint messageHandle = this.CreatedMessage(logonResponse.FolderIds[5], this.insideObjHandle);

            // Add a Recipient that message send to.
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(Common.GetConfigurationPropertyValue("AdminUserName", this.Site), 0)
            };
            this.AddRecipients(modifyRecipientRow, messageHandle, propertyTag);

            // Set the message properties.
            List<PropertyObj> propertyListForSet = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagBody, Common.GetBytesFromUnicodeString(TestDataOfPidTagBody))
            };
            string title = Common.GenerateResourceName(Site, "Mail");
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(title)));
            this.SetPropertiesForMessage(messageHandle, propertyListForSet);

            RopSaveChangesMessageResponse saveMessageResp = this.SaveMessage(messageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveMessageResp.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            RopSubmitMessageRequest submitMessageRequest = new RopSubmitMessageRequest()
            {
                RopId = (byte)RopId.RopSubmitMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                SubmitFlags = (byte)SubmitFlags.None,
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(submitMessageRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSubmitMessageResponse submitMessageResponse = (RopSubmitMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, submitMessageResponse.ReturnValue, "Call RopSubmitMessage should success.");
            #endregion

            #region Receive the message sent by step 2.
            ulong messageId = 0;
            isMessageReceived = this.WaitEmailBeDelivered(title, logonResponse.FolderIds[4], this.insideObjHandle, out messageId);

            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            propertiesOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNativeBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfCompressed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfInSync],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetCodepage],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };

            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertiesOfMessage);

            #region Verify requirements
            PropertyObj pidTagNativeBody = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNativeBody);
            if (Common.IsRequirementEnabled(1714, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R129");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R129
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000001,
                    Convert.ToInt32(pidTagNativeBody.Value),
                    2059,
                    @"[In PidTagNativeBody Property] [The value 0x00000001 indicates the best available format for storing the message body is] Plain text body.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1714");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1714
                // Because PidTagNativeBody is not null then the implementation does support the PidTagNativeBody property.
                this.Site.CaptureRequirementIfIsNotNull(
                    pidTagNativeBody,
                    1714,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagNativeBody property. (Exchange 2010 and above follow this behavior.)");
            }

            PropertyObj pidTagMessageFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagMessageFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R519");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R519
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(pidTagMessageFlags.Value) & (int)MessageFlags.MfUnmodified,
                519,
                @"[In PidTagMessageFlags Property] [mfUnmodified (0x00000002)] The message has not been modified since it was first saved (if unsent) or it was delivered (if sent).");

            #endregion
            #endregion

            #region Call ROPs to send a message that contains HTML body to mailbox.
            // Initialize message information, and set the handle for this message.
            // Create a message in Outbox folder.
            messageHandle = this.CreatedMessage(logonResponse.FolderIds[5], this.insideObjHandle);

            // Add a Recipient that message send to.
            this.AddRecipients(modifyRecipientRow, messageHandle, propertyTag);

            // Set the message properties.
            propertyListForSet = new List<PropertyObj>();
            title = Common.GenerateResourceName(this.Site, "Mail");
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(title)));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagBodyHtml, Common.GetBytesFromUnicodeString(TestDataOfPidTagBodyHtml)));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagHtml, PropertyHelper.GetBinaryFromGeneral(Encoding.ASCII.GetBytes(TestDataOfPidTagBodyHtml))));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagBodyContentId, Common.GetBytesFromUnicodeString(TestDataOfPidTagBodyContentId)));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagBodyContentLocation, Common.GetBytesFromUnicodeString(TestDataOfPidTagBodyContentLocation)));

            this.SetPropertiesForMessage(messageHandle, propertyListForSet);

            saveMessageResp = this.SaveMessage(messageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveMessageResp.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(submitMessageRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            submitMessageResponse = (RopSubmitMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, submitMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Receive the message sent by step 7.
            messageId = 0;
            isMessageReceived = this.WaitEmailBeDelivered(title, logonResponse.FolderIds[4], this.insideObjHandle, out messageId);

            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            propertiesOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNativeBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfCompressed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfInSync],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetCodepage],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyContentId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyContentLocation],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyHtml]
            };

            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertiesOfMessage);

            pidTagNativeBody = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNativeBody);
            PropertyObj pidTagHtml = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagHtml);

            if (Common.IsRequirementEnabled(1714, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R131");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R131
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000003,
                    Convert.ToInt32(pidTagNativeBody.Value),
                    2061,
                    @"[In PidTagNativeBody Property] [The value 0x00000003 indicates the best available format for storing the message body is] HTML body.");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2082");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2082
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagHtml.Value,
                2082,
                @"[In PidTagHtml Property] [Type] The PidTagHtml property ([MS-OXPROPS] section 2.722) contains the message body text in HTML format.");

            PropertyObj pidTagBodyContentLocation = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagBodyContentLocation);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2076, the value of PidTagBodyContentLocation is {0}.", pidTagBodyContentLocation.Value.ToString());

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2076
            bool isVerifiedR2076 = Uri.IsWellFormedUriString(pidTagBodyContentLocation.Value.ToString(), UriKind.RelativeOrAbsolute);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR2076,
                2076,
                @"[In PidTagBodyContentLocation Property] The PidTagBodyContentLocation property ([MS-OXPROPS] section 2.611) contains a globally unique URI that serves as a label for the current message body.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2077, the value of PidTagBodyContentLocation is {0}.", pidTagBodyContentLocation.Value.ToString());

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2077
            bool isVerifiedR2077 = Uri.IsWellFormedUriString(pidTagBodyContentLocation.Value.ToString(), UriKind.RelativeOrAbsolute);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR2077,
                2077,
                @"[In PidTagBodyContentLocation Property] The URI can be either absolute or relative.");
            #endregion

            #region Call ROPs to send a message that contains a Clear-signed body to mailbox.
            #region Call RopCreateMessage to create Message object.
            messageHandle = this.CreatedMessage(logonResponse.FolderIds[5], this.insideObjHandle);
            #endregion

            #region Call RopModifyRecipients to add recipient to message created by step2
            this.AddRecipients(modifyRecipientRow, messageHandle, propertyTag);
            #endregion

            #region Call RopSetProperties to set properties of created message.
            propertyListForSet = new List<PropertyObj>();
            title = Common.GenerateResourceName(this.Site, "Mail");
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(title)));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagMessageClass, Common.GetBytesFromUnicodeString("IPM.Note.SMIME.MultipartSigned")));
            propertyListForSet.Add(new PropertyObj(PropertyNames.PidTagBody, Common.GetBytesFromUnicodeString(TestDataOfPidTagBody)));

            this.SetPropertiesForMessage(messageHandle, propertyListForSet);
            #endregion

            #region RopCreateAttachment success response
            RopCreateAttachmentRequest createAttachmentRequest = new RopCreateAttachmentRequest()
            {
                RopId = (byte)RopId.RopCreateAttachment,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(createAttachmentRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopCreateAttachmentResponse createAttachmentResponse = (RopCreateAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, createAttachmentResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            uint attachmentHandle = this.ResponseSOHs[0][createAttachmentResponse.OutputHandleIndex];
            #endregion

            propertyListForSet = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagAttachMethod, BitConverter.GetBytes(0x00000001)),
                new PropertyObj(PropertyNames.PidTagAttachMimeTag, Common.GetBytesFromUnicodeString("multipart/signed")),
                new PropertyObj(PropertyNames.PidTagAttachFilename, Common.GetBytesFromUnicodeString("SMIME.p7m")),
                new PropertyObj(PropertyNames.PidTagAttachLongFilename, Common.GetBytesFromUnicodeString("SMIME.p7m")),
                new PropertyObj(PropertyNames.PidTagDisplayName, Common.GetBytesFromUnicodeString("SMIME.p7m"))
            };

            this.SetPropertiesForMessage(attachmentHandle, propertyListForSet);

            #region RopSaveChangesAttachment success response
            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest = new RopSaveChangesAttachmentRequest()
            {
                RopId = (byte)RopId.RopSaveChangesAttachment,
                LogonId = CommonLogonId,
                ResponseHandleIndex = CommonOutputHandleIndex, // This index specifies the location in the Server object handle table that is referenced in the response.
                InputHandleIndex = CommonInputHandleIndex,
                SaveFlags = 0x0C // ForceSave
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(saveChangesAttachmentRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse = (RopSaveChangesAttachmentResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesAttachmentResponse.ReturnValue, "Call RopSaveChangesAttachment should success.");
            #endregion RopSaveChangesAttachment success response

            #region Call RopSaveChangesMessage to save created message.
            saveMessageResp = this.SaveMessage(messageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveMessageResp.ReturnValue, "Call RopSaveChangesMessage should success.");
            #endregion

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(submitMessageRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            submitMessageResponse = (RopSubmitMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, submitMessageResponse.ReturnValue, "Call RopSubmitMessage should success.");
            #endregion

            #region Receive the message sent by step 12.
            messageId = 0;
            isMessageReceived = this.WaitEmailBeDelivered(title, logonResponse.FolderIds[4], this.insideObjHandle, out messageId);

            Site.Assert.IsTrue(isMessageReceived, "The message should be received.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            propertiesOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNativeBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfCompressed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfInSync],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetCodepage],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyContentId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyContentLocation],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags]
            };

            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertiesOfMessage);

            pidTagNativeBody = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNativeBody);
            #endregion

            #region Create an undefined body message in inbox folder
            messageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);

            saveMessageResp = this.SaveMessage(messageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveMessageResp.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            propertiesOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNativeBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfCompressed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfInSync],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetCodepage]
            };

            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveMessageResp.MessageId, this.insideObjHandle, propertiesOfMessage);

            #region Verify requirements
            if (Common.IsRequirementEnabled(1714, this.Site))
            {
                pidTagNativeBody = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNativeBody);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R128");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R128
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000000,
                    Convert.ToInt32(pidTagNativeBody.Value),
                    2058,
                    @"[In PidTagNativeBody Property] [The value 0x00000000 indicates the best available format for storing the message body is] Undefined body.");
            }
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(messageHandle);
            #endregion

            #region Create a new Rich Text Format (RTF) compressed body message in inbox folder.
            messageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);

            propertyListForSet = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRtfCompressed, PropertyHelper.GetBinaryFromGeneral(this.rtfCompressedText))
            };
            this.SetPropertiesForMessage(messageHandle, propertyListForSet);

            saveMessageResp = this.SaveMessage(messageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveMessageResp.ReturnValue, "Call RopSaveChangesMessage should success.");
            #endregion

            #region Call RopGetPropertiesSpecific to get the specific properties of created message.
            propertiesOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNativeBody],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagBodyHtml],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfCompressed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRtfInSync],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetCodepage],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagObjectType],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRecordKey]
            };

            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveMessageResp.MessageId, this.insideObjHandle, propertiesOfMessage);

            #region Verify requirements
            if (Common.IsRequirementEnabled(1714, this.Site))
            {
                pidTagNativeBody = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNativeBody);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R130");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R130
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000002,
                    Convert.ToInt32(pidTagNativeBody.Value),
                    2060,
                    @"[In PidTagNativeBody Property] [The value 0x00000002 indicates the best available format for storing the message body is] RTF compressed body.");
            }

            PropertyObj pidTagRtfInSync = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRtfInSync);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R138");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R138
            this.Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(pidTagRtfInSync.Value),
                2068,
                @"[In PidTagRtfInSync Property] The PidTagRtfInSync property ([MS-OXPROPS] section 2.931) is set to ""TRUE"" (0x01) if the RTF body has been synchronized with the contents in the PidTagBody property (section 2.2.1.56.1).");
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(messageHandle);
            #endregion
        }

        /// <summary>
        /// This test case validates the retention properties and archive properties on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC03_RetentionAndArchivePropertiesOnMessageObject()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call SetPropertiesSpecific to set the retention property of the created message.
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagArchiveTag, PropertyHelper.GetBinaryFromGeneral(Guid.NewGuid().ToByteArray())),
                new PropertyObj(PropertyNames.PidTagPolicyTag, PropertyHelper.GetBinaryFromGeneral(Guid.NewGuid().ToByteArray())),
                new PropertyObj(PropertyNames.PidTagRetentionPeriod, BitConverter.GetBytes(0x00000001))
            };
            List<byte> lstBytes = new List<byte>();
            lstBytes.AddRange(BitConverter.GetBytes(0x00000001));
            lstBytes.AddRange(BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagStartDateEtc, PropertyHelper.GetBinaryFromGeneral(lstBytes.ToArray())));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagRetentionDate, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc())));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000002)));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagArchivePeriod, BitConverter.GetBytes(0x00000002)));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagArchiveDate, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc())));
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, "Call RopSaveChangesMessage should success.");
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set PropertySizeLimit,which specifies the maximum size allowed for a property value returned,
                // as specified in [MS-OXCROPS] section 2.2.8.4.1.
                PropertySizeLimit = 0xFFFF,
                WantUnicode = 1
            };

            // In process, call capture code to verify adapter requirement
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            RopGetPropertiesAllResponse getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetPropertiesAll should success.");
            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);

            #region Verify requirements
            PropertyObj pidTagArchiveTag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagArchiveTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R579");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R579
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagArchiveTag,
                2194,
                @"[In PidTagArchiveTag Property] The PidTagArchiveTag property can be present on Message objects.");

            PropertyObj pidTagPolicyTag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagPolicyTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R581");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R581
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagPolicyTag,
                2113,
                @"[In PidTagPolicyTag Property] The PidTagPolicyTag property can be present on Message objects.");

            PropertyObj pidTagRetentionPeriod = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRetentionPeriod);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1658");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1658
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRetentionPeriod,
                2119,
                @"[In PidTagRetentionPeriod Property] The PidTagRetentionPeriod property can be present on Message objects.");

            PropertyObj pidTagStartDateEtc = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagStartDateEtc);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R176");
            byte[] startDateEtc = (byte[])pidTagStartDateEtc.Value;
            int count = BitConverter.ToInt16(startDateEtc, 0);
            byte[] defaultRetentionPeriod = new byte[4];
            byte[] startDate = new byte[startDateEtc.Length - 4 - 2];
            Array.Copy(startDateEtc, 2, defaultRetentionPeriod, 0, 4);
            Array.Copy(startDateEtc, 6, startDate, 0, startDate.Length);
            DateTime start = DateTime.FromFileTimeUtc(BitConverter.ToInt64(startDate, 0));
            Site.Assert.IsNotNull(start, "The start date should not null.");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R176
            this.Site.CaptureRequirementIfAreEqual<int>(
                count - 4,
                startDate.Length,
                2131,
                @"[In PidTagStartDateEtc Property] [The length of] Start date [is 8 bytes], [which contains] The date, in UTC, from which the age of the Message object is calculated.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R177");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R177
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagStartDateEtc,
                2133,
                @"[In PidTagStartDateEtc Property] The PidTagStartDateEtc property can be present only on Message objects.");

            PropertyObj pidTagRetentionDate = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRetentionDate);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R181");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R181
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRetentionDate,
                2136,
                @"[In PidTagRetentionDate Property] The PidTagRetentionDate property can be present only on Message objects.");

            PropertyObj pidTagRetentionFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R585");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R585
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRetentionFlags,
                2149,
                @"[In PidTagRetentionFlags Property] The PidTagRetentionFlags property can be present on Message objects.");

            PropertyObj pidTagArchivePeriod = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagArchivePeriod);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R207");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R207
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagArchivePeriod,
                2165,
                @"[In PidTagArchivePeriod Property] The PidTagArchivePeriod property can be present on Message objects.");

            PropertyObj pidTagArchiveDate = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagArchiveDate);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R214");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R214
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagArchiveDate,
                2174,
                @"[In PidTagArchiveDate Property] The PidTagArchiveDate property can be present on only Message objects, not on folders.");
            #endregion
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case validates the properties and archive properties on folder. 
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC04_RetentionAndArchivePropertiesOnFolder()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);

            #region Call RopOpenFolder to open an existing folder
            uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopCreateFolder to create a new folder.
            ulong folderId;
            this.CreateSubFolder(openedFolderHandle, out folderId);
            #endregion

            #region Call RopOpenFolder to open a folder create by above step.
            uint openSubfolderHandle = this.OpenSpecificFolder(folderId, this.insideObjHandle);
            #endregion

            #region Call RopSetProperties to set properties of retention and archive.
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagArchiveTag, PropertyHelper.GetBinaryFromGeneral(Guid.NewGuid().ToByteArray())),
                new PropertyObj(PropertyNames.PidTagPolicyTag, PropertyHelper.GetBinaryFromGeneral(Guid.NewGuid().ToByteArray())),
                new PropertyObj(PropertyNames.PidTagRetentionPeriod, BitConverter.GetBytes(0x00000001))
            };
            List<byte> lstBytes = new List<byte>();
            lstBytes.AddRange(BitConverter.GetBytes(0x00000001));
            lstBytes.AddRange(BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagStartDateEtc, PropertyHelper.GetBinaryFromGeneral(lstBytes.ToArray())));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagRetentionDate, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc())));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000002)));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagArchivePeriod, BitConverter.GetBytes(0x00000002)));
            propertyList.Add(new PropertyObj(PropertyNames.PidTagArchiveDate, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc())));

            this.SetPropertiesForMessage(openSubfolderHandle, propertyList);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of the created folder.
            RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set PropertySizeLimit,which specifies the maximum size allowed for a property value returned,
                // as specified in [MS-OXCROPS] section 2.2.8.4.1.
                PropertySizeLimit = 0xFFFF,
                WantUnicode = 1
            };

            // In process, call capture code to verify adapter requirement
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, openSubfolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesAllResponse getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetPropertiesAll should success.");
            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);

            #region Verify requirements
            PropertyObj pidTagArchiveTag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagArchiveTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1818");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1818
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagArchiveTag,
                2195,
                @"[In PidTagArchiveTag Property] [In PidTagArchiveTag Property] The PidTagArchiveTag property can be present on folders.");

            PropertyObj pidTagPolicyTag = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagPolicyTag);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1819");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1819
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagPolicyTag,
                2114,
                @"[In PidTagPolicyTag Property] The PidTagPolicyTag property can be present on folders.");

            PropertyObj pidTagRetentionPeriod = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRetentionPeriod);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1820");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1820
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRetentionPeriod,
                2120,
                @"[In PidTagRetentionPeriod Property] The PidTagRetentionPeriod property can be present on folders.");

            PropertyObj pidTagRetentionFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1822");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1822
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagRetentionFlags,
                2150,
                @"[In PidTagRetentionFlags Property] The PidTagRetentionFlags property can be present on folders.");

            PropertyObj pidTagArchivePeriod = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagArchivePeriod);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1823");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1823
            this.Site.CaptureRequirementIfIsNotNull(
                pidTagArchivePeriod,
                2166,
                @"[In PidTagArchivePeriod Property] The PidTagArchivePeriod property can be present on folders.");
            #endregion
            #endregion

            #region Call RopDeleteFolder to delete the folder created by above step.
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest()
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete | (byte)DeleteFolderFlags.DelMessages,
                FolderId = folderId // Folder to be deleted
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteFolderRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteFolderResponse deleteFolderresponse = (RopDeleteFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteFolderresponse.ReturnValue, "Call RopDeleteFolder should success.");
            #endregion

            #region Call RopRelease to release all resources
            this.ReleaseRop(openSubfolderHandle);
            this.ReleaseRop(openedFolderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the value of PidTagRetentionFlags property on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC05_PidTagRetentionFlagsProperty()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to ExplicitTag.
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000001))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            RopGetPropertiesAllRequest getPropertiesAllRequest = new RopGetPropertiesAllRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesAll,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,

                // Set PropertySizeLimit,which specifies the maximum size allowed for a property value returned,
                // as specified in [MS-OXCROPS] section 2.2.8.4.1.
                PropertySizeLimit = 0xFFFF,
                WantUnicode = 1
            };

            // In process, call capture code to verify adapter requirement
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            RopGetPropertiesAllResponse getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            List<PropertyObj> propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            PropertyObj pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R945");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R945
            this.Site.CaptureRequirementIfAreEqual<int>(
                Convert.ToInt32(pidTagRetentionFlags.Value),
                0x00000001,
                2154,
                @"[In PidTagRetentionFlags Property] [ExplicitTag (0x00000001)] The retention tag on the folder is explicitly set.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to UserOverride.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000002))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R946");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R946
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2155,
                @"[In PidTagRetentionFlags Property] [UserOverride (0x00000002)] The retention tag was not changed by the end user.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to AutoTag.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000004))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R947");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R947
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000004,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2156,
                @"[In PidTagRetentionFlags Property] [AutoTag (0x00000004)] The retention tag on the Message object is an autotag, which is predicted by the system.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to PersonalTag.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000008))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R948");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R948
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000008,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2157,
                @"[In PidTagRetentionFlags Property] [PersonalTag (0x00000008)] The retention tag on the folder is of a personal type and can be made available to the end user.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to ExplicitArchiveTag.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000010))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R949");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R949
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000010,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2158,
                @"[In PidTagRetentionFlags Property] [ExplicitArchiveTag (0x00000010)] The archive tag on the folder is explicitly set.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to KeepInPlace.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000020))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R950");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R950
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000020,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2159,
                @"[In PidTagRetentionFlags Property] [KeepInPlace (0x00000020)] The Message object remains in place and is not archived.");
            #endregion

            #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to SystemData.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000040))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesAll to get all properties of created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
            getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
            pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R951");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R951
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000040,
                Convert.ToInt32(pidTagRetentionFlags.Value),
                2160,
                @"[In PidTagRetentionFlags Property] [SystemData (0x00000040)] The Message object or folder is system data.");
            #endregion

            if (Common.IsRequirementEnabled(1909, this.Site))
            {
                #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to NeedsRescan.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000080))
                };

                this.SetPropertiesForMessage(targetMessageHandle, propertyList);
                #endregion

                #region Call RopSaveChangesMessage to save created message.
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
                #endregion

                #region Call RopGetPropertiesAll to get all properties of created message.
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
                getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
                pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1909");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1909
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000080,
                    Convert.ToInt32(pidTagRetentionFlags.Value),
                    1909,
                    @"[In Appendix A: Product Behavior] Implementation does support the NeedsRescan flag. (Exchange 2010 SP2 and above follows this behavior.)");
                #endregion
            }

            if (Common.IsRequirementEnabled(1911, this.Site))
            {
                #region Call SetPropertiesSpecific to set PidTagRetentionFlags property to PendingRescan.
                propertyList = new List<PropertyObj>
                {
                    new PropertyObj(PropertyNames.PidTagRetentionFlags, BitConverter.GetBytes(0x00000100))
                };

                this.SetPropertiesForMessage(targetMessageHandle, propertyList);
                #endregion

                #region Call RopSaveChangesMessage to save created message.
                saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
                #endregion

                #region Call RopGetPropertiesAll to get all properties of created message.
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesAllRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.MessageProperties);
                getPropertiesAllResponse = (RopGetPropertiesAllResponse)this.response;
                Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesAllResponse.ReturnValue, "Call RopGetProperties should success.");

                propertyValues = PropertyHelper.GetPropertyObjFromBuffer(getPropertiesAllResponse);
                pidTagRetentionFlags = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagRetentionFlags);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1911");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1911
                this.Site.CaptureRequirementIfAreEqual<int>(
                    0x00000100,
                    Convert.ToInt32(pidTagRetentionFlags.Value),
                    1911,
                    @"[In Appendix A: Product Behavior] Implementation does support the PendingRescan flag. (Exchange 2010 SP2 and above follow this behavior.)");
                #endregion
            }

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the PidLidSideEffects value on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC06_PidLidSideEffectsProperty()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyNameObject> propertyNameList = new List<PropertyNameObject>
            {
                new PropertyNameObject(PropertyNames.PidLidSideEffects, (uint)PropertyLID.PidLidSideEffects, PropertySet.PSETIDCommon, PropertyType.PtypInteger32)
            };

            PropertyNameObject propertyName = new PropertyNameObject(PropertyNames.PidLidSideEffects, (uint)PropertyLID.PidLidSideEffects, PropertySet.PSETIDCommon, PropertyType.PtypInteger32);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Set PidLidSideEffects property to 0x00000001 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000001));
            #endregion

            #region Get PidLidSideEffects property value.
            Dictionary<PropertyNames, byte[]> propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2008");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2008
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2008,
                @"[In PidLidSideEffects Property] [The value of flag seOpenToDelete is] 0x00000001.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000008 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000008));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2009");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2009
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000008,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2009,
                @"[In PidLidSideEffects Property] [The value of flag seNoFrame is] 0x00000008.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000010 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000010));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2010");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2010
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000010,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2010,
                @"[In PidLidSideEffects Property] [The value of flag seCoerceToInbox is] 0x00000010.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000020 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000020));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2011");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2011
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000020,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2011,
                @"[In PidLidSideEffects Property] [The value of flag seOpenToCopy is] 0x00000020.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000040 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000040));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2012");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2012
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000040,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2012,
                @"[In PidLidSideEffects Property] [The value of flag seOpenToMove is] 0x00000040.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000100 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000100));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2013");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2013
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000100,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2013,
                @"[In PidLidSideEffects Property] [The value of flag seOpenForCtxMenu is] 0x00000100.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000400 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000400));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2014");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2014
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000400,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2014,
                @"[In PidLidSideEffects Property] [The value of flag seCannotUndoDelete is] 0x00000400.");
            #endregion

            #region Set PidLidSideEffects property to 0x00000800 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000800));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2015");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2015
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000800,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2015,
                @"[In PidLidSideEffects Property] [The value of flag seCannotUndoCopy is] 0x00000800.");
            #endregion

            #region Set PidLidSideEffects property to 0x00001000 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00001000));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2016");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2016
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00001000,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2016,
                @"[In PidLidSideEffects Property] [The value of flag seCannotUndoMove is] 0x00001000.");
            #endregion

            #region Set PidLidSideEffects property to 0x00002000 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00002000));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2017");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2017
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00002000,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2017,
                @"[In PidLidSideEffects Property] [The value of flag seHasScript is] 0x00002000.");
            #endregion

            #region Set PidLidSideEffects property to 0x00004000 in specified Message created by step 2.
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00004000));
            #endregion

            #region Get PidLidSideEffects property value.
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2018");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2018
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00004000,
                BitConverter.ToInt32(propertyValues[PropertyNames.PidLidSideEffects], 0),
                2018,
                @"[In PidLidSideEffects Property] [The value of flag seOpenToPermDelete is] 0x00004000.");
            #endregion

            this.ReleaseRop(targetMessageHandle);
        }

        /// <summary>
        /// This test case tests the named properties on Message object.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC07_NamedPropertiesOnMessageObject()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            List<PropertyNameObject> propertyNameList = new List<PropertyNameObject>
            {
                new PropertyNameObject(PropertyNames.PidLidPrivate, (uint)PropertyLID.PidLidPrivate, PropertySet.PSETIDCommon, PropertyType.PtypBoolean),
                new PropertyNameObject(PropertyNames.PidLidClassification, (uint)PropertyLID.PidLidClassification, PropertySet.PSETIDCommon, PropertyType.PtypString),
                new PropertyNameObject(PropertyNames.PidLidClassificationDescription, (uint)PropertyLID.PidLidClassificationDescription, PropertySet.PSETIDCommon, PropertyType.PtypString),
                new PropertyNameObject(PropertyNames.PidLidClassified, (uint)PropertyLID.PidLidClassified, PropertySet.PSETIDCommon, PropertyType.PtypBoolean),
                new PropertyNameObject(PropertyNames.PidLidInfoPathFormName, (uint)PropertyLID.PidLidInfoPathFormName, PropertySet.PSETIDCommon, PropertyType.PtypString),
                new PropertyNameObject(PropertyNames.PidLidCurrentVersion, (uint)PropertyLID.PidLidCurrentVersion, PropertySet.PSETIDCommon, PropertyType.PtypInteger32), new PropertyNameObject(PropertyNames.PidLidCurrentVersionName, (uint)PropertyLID.PidLidCurrentVersionName, PropertySet.PSETIDCommon, PropertyType.PtypString),
                new PropertyNameObject(PropertyNames.PidLidAgingDontAgeMe, (uint)PropertyLID.PidLidAgingDontAgeMe, PropertySet.PSETIDCommon, PropertyType.PtypBoolean),
                new PropertyNameObject(PropertyNames.PidLidCommonStart, (uint)PropertyLID.PidLidCommonStart, PropertySet.PSETIDCommon, PropertyType.PtypTime),
                new PropertyNameObject(PropertyNames.PidLidCommonEnd, (uint)PropertyLID.PidLidCommonEnd, PropertySet.PSETIDCommon, PropertyType.PtypTime),
                new PropertyNameObject(PropertyNames.PidLidSmartNoAttach, (uint)PropertyLID.PidLidSmartNoAttach, PropertySet.PSETIDCommon, PropertyType.PtypBoolean)
            };

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Set specified properties of message created by above step.
            PropertyNameObject propertyName = new PropertyNameObject(PropertyNames.PidLidPrivate, (uint)PropertyLID.PidLidPrivate, PropertySet.PSETIDCommon, PropertyType.PtypBoolean);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(true));
            propertyName = new PropertyNameObject(PropertyNames.PidLidClassification, (uint)PropertyLID.PidLidClassification, PropertySet.PSETIDCommon, PropertyType.PtypString);
            this.SetNamedProperty(targetMessageHandle, propertyName, Common.GetBytesFromUnicodeString("Classification"));
            propertyName = new PropertyNameObject(PropertyNames.PidLidClassificationDescription, (uint)PropertyLID.PidLidClassificationDescription, PropertySet.PSETIDCommon, PropertyType.PtypString);
            this.SetNamedProperty(targetMessageHandle, propertyName, Common.GetBytesFromUnicodeString("Classification description"));
            propertyName = new PropertyNameObject(PropertyNames.PidLidClassified, (uint)PropertyLID.PidLidClassified, PropertySet.PSETIDCommon, PropertyType.PtypBoolean);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(true));
            propertyName = new PropertyNameObject(PropertyNames.PidLidInfoPathFormName, (uint)PropertyLID.PidLidInfoPathFormName, PropertySet.PSETIDCommon, PropertyType.PtypString);
            this.SetNamedProperty(targetMessageHandle, propertyName, Common.GetBytesFromUnicodeString(TestDataOfInfoPath));
            propertyName = new PropertyNameObject(PropertyNames.PidLidCurrentVersion, (uint)PropertyLID.PidLidCurrentVersion, PropertySet.PSETIDCommon, PropertyType.PtypInteger32);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(0x00000001));
            propertyName = new PropertyNameObject(PropertyNames.PidLidCurrentVersionName, (uint)PropertyLID.PidLidCurrentVersionName, PropertySet.PSETIDCommon, PropertyType.PtypString);
            this.SetNamedProperty(targetMessageHandle, propertyName, Common.GetBytesFromUnicodeString("Current version name"));
            propertyName = new PropertyNameObject(PropertyNames.PidLidSmartNoAttach, (uint)PropertyLID.PidLidSmartNoAttach, PropertySet.PSETIDCommon, PropertyType.PtypBoolean);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(false));
            propertyName = new PropertyNameObject(PropertyNames.PidLidCommonStart, (uint)PropertyLID.PidLidCommonStart, PropertySet.PSETIDCommon, PropertyType.PtypTime);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()));
            propertyName = new PropertyNameObject(PropertyNames.PidLidCommonEnd, (uint)PropertyLID.PidLidCommonEnd, PropertySet.PSETIDCommon, PropertyType.PtypTime);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(DateTime.Parse(TestDataOfDateTime).ToFileTimeUtc()));
            propertyName = new PropertyNameObject(PropertyNames.PidLidAgingDontAgeMe, (uint)PropertyLID.PidLidAgingDontAgeMe, PropertySet.PSETIDCommon, PropertyType.PtypBoolean);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(true));
            #endregion

            #region Get properties value in message
            Dictionary<PropertyNames, byte[]> propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1654");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1654
            this.Site.CaptureRequirementIfIsTrue(
                BitConverter.ToBoolean(propertyValues[PropertyNames.PidLidAgingDontAgeMe], 0),
                1654,
                @"[In PidLidAgingDontAgeMe Property] This property [PidLidAgingDontAgeMe] is set to ""TRUE"" if the message will not be automatically archived.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R91");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R91
            this.Site.CaptureRequirementIfIsFalse(
                 BitConverter.ToBoolean(propertyValues[PropertyNames.PidLidSmartNoAttach], 0),
                91,
                @"[In PidLidSmartNoAttach Property] If this property is unset, a default value of FALSE (0x00) is used.");

            #endregion

            #region Set specified properties of message created by above step.
            propertyName = new PropertyNameObject(PropertyNames.PidLidAgingDontAgeMe, (uint)PropertyLID.PidLidAgingDontAgeMe, PropertySet.PSETIDCommon, PropertyType.PtypBoolean);
            this.SetNamedProperty(targetMessageHandle, propertyName, BitConverter.GetBytes(false));

            #endregion

            #region Get properties value in message
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1655");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1655
            this.Site.CaptureRequirementIfIsFalse(
                BitConverter.ToBoolean(propertyValues[PropertyNames.PidLidAgingDontAgeMe], 0),
                1655,
                @"[In PidLidAgingDontAgeMe Property] otherwise [the message will be automatically archived], [this property is set to] ""FALSE"".");
            #endregion

            this.ReleaseRop(targetMessageHandle);

            propertyNameList = new List<PropertyNameObject>();
            PropertyNameObject contentBaseName = new PropertyNameObject(PropertyNames.PidNameContentBase, "Content-Base", PropertySet.PSINTERNETHEADERS, PropertyType.PtypString);
            propertyNameList.Add(contentBaseName);
            PropertyNameObject acceptLanguageName = new PropertyNameObject(PropertyNames.PidNameAcceptLanguage, "Accept-Language", PropertySet.PSINTERNETHEADERS, PropertyType.PtypString);
            propertyNameList.Add(acceptLanguageName);
            PropertyNameObject contentClassName = new PropertyNameObject(PropertyNames.PidNameContentClass, "Content-Class", PropertySet.PSINTERNETHEADERS, PropertyType.PtypString);
            propertyNameList.Add(contentClassName);
            PropertyNameObject contentType = new PropertyNameObject(PropertyNames.PidNameContentType, "Content-Type", PropertySet.PSINTERNETHEADERS, PropertyType.PtypString);
            propertyNameList.Add(contentType);
            PropertyNameObject keywords = new PropertyNameObject(PropertyNames.PidNameKeywords, "Keywords", PropertySet.PSPUBLICSTRINGS, PropertyType.PtypMultipleString);
            propertyNameList.Add(keywords);

            #region Call RopCreateMessage to create Message object.
            targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Set specified properties of message created by above step.
            this.SetNamedProperty(targetMessageHandle, contentBaseName, Common.GetBytesFromUnicodeString(TestDataOfPidNameContentBase));
            this.SetNamedProperty(targetMessageHandle, acceptLanguageName, Common.GetBytesFromUnicodeString(TestDataOfPidNameAcceptLanguage));
            this.SetNamedProperty(targetMessageHandle, contentClassName, Common.GetBytesFromUnicodeString(TestDataOfPidNameContentClass));
            this.SetNamedProperty(targetMessageHandle, contentType, Common.GetBytesFromUnicodeString(TestDataOfPidNameContentType));
            this.SetNamedProperty(targetMessageHandle, keywords, Common.GetBytesFromMutiUnicodeString(new string[] { TestDataOfPidNameKeywords }));
            #endregion

            #region Get properties value in message
            propertyValues = this.MSOXCMSGAdapter.GetNamedPropertyValues(propertyNameList, targetMessageHandle);
            string contentBaseValue = Encoding.Unicode.GetString(propertyValues[PropertyNames.PidNameContentBase], 0, propertyValues[PropertyNames.PidNameContentBase].Length - 2);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1208, The value of PidNameContentBase is {0}.", contentBaseValue);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1208
            bool isVerifiedR1208 = Uri.IsWellFormedUriString(contentBaseValue, UriKind.RelativeOrAbsolute);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1208,
                1208,
                @"[In PidNameContentBase Property] The PidNameContentBase property ([MS-OXPROPS] section 2.406) specifies the value of the Content-Base header (2), which defines the base Uniform Resource Identifier (URI) for resolving relative URLs contained within the message body.");

            string acceptLanguageValue = Encoding.Unicode.GetString(propertyValues[PropertyNames.PidNameAcceptLanguage], 0, propertyValues[PropertyNames.PidNameAcceptLanguage].Length - 2);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1210");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1210
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidNameAcceptLanguage,
                acceptLanguageValue,
                1210,
                @"[In PidNameAcceptLanguage Property] The PidNameAcceptLanguage property ([MS-OXPROPS] section 2.366) contains the value of the Accept-Language header (2), which defines the natural languages in which the sender prefers to receive a response.");

            string contentTypeValue = Encoding.Unicode.GetString(propertyValues[PropertyNames.PidNameContentType], 0, propertyValues[PropertyNames.PidNameContentType].Length - 2);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2038");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2038
            this.Site.CaptureRequirementIfAreEqual<string>(
                TestDataOfPidNameContentType,
                contentTypeValue,
                2038,
                @"[In PidNameContentType Property] The PidNameContentType property ([MS-OXPROPS] section 2.408) contains the value of the Content-Type header (2), which defines the type of the body part's content.");           
        
            #endregion

            #region Call RopSaveMessage to commit the message changes
            RopSaveChangesMessageResponse saveChangesResposne = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesResposne.ReturnValue, "Call RopSaveChangesMessage ROP should success.");
            ulong messageId = saveChangesResposne.MessageId;
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is read-write to open created message.
            RopOpenMessageResponse openMessageResp;
            this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.ReadOnly, out openMessageResp);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1799");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1799
            this.Site.CaptureRequirementIfIsTrue(
                Convert.ToBoolean(openMessageResp.HasNamedProperties),
                1799,
                @"[In PidTagHasNamedProperties Property] The PidTagHasNamedProperties property is set to ""TRUE"" if this Message object supports named properties.");
            #endregion

            this.ReleaseRop(targetMessageHandle);
        }

        /// <summary>
        /// This test case tests that the value of PidTagMessageClass property MUST be case-insensitive.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC08_PidTagMessageClassCaseInsensitive()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetProperties to set PidTagMessageClass and the value is lower case string.
            string messageClass = "IPM.Note";
            RopSetPropertiesResponse setPropertiesResponse;
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagMessageClass, Common.GetBytesFromUnicodeString(messageClass.ToLower()))
            };
            this.SetPropertiesForMessage(targetMessageHandle, propertyList, out setPropertiesResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, "Call RopSetProperties should success.");
            #endregion

            #region Call RopSetProperties to set PidTagMessageClass and the value is upper case string.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagMessageClass, Common.GetBytesFromUnicodeString(messageClass.ToUpper()))
            };
            this.SetPropertiesForMessage(targetMessageHandle, propertyList, out setPropertiesResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, "Call RopSetProperties should success.");

            // R1605 can be verified because RopSetProperties is called with lower case and upper case PidTagMessageClass value.
            this.Site.CaptureRequirement(
                1605,
                @"[In PidTagMessageClass Property] Any equality or matching operations performed against the value of this property [PidTagMessageClass] MUST be case-insensitive.");
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests the value of PidTagImportance, PidTagPriority and PidTagSensitivity.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S04_TC09_PropertiesForMessageSetting()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetProperties to set properties of message created by step 2.
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagSubjectPrefix, Common.GetBytesFromUnicodeString(string.Empty)),
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject)),
                new PropertyObj(PropertyNames.PidTagImportance, BitConverter.GetBytes(0x00000002)),
                new PropertyObj(PropertyNames.PidTagPriority, BitConverter.GetBytes(0x00000001)),
                new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000002)),
                new PropertyObj(PropertyNames.PidTagResponsibility, BitConverter.GetBytes(false)),
                new PropertyObj(PropertyNames.PidTagTrustSender, BitConverter.GetBytes(0x00000000))
            };

            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save message created by step 2.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopGetPropertiesSpecific to get all properties of message created by step 2.
            List<PropertyTag> propertyTagsOfMessage = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagImportance],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagPriority],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSensitivity],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSubject],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNormalizedSubject],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSubjectPrefix],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagResponsibility],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagTrustSender]
            };

            List<PropertyObj> propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTagsOfMessage);

            PropertyObj pidTagImportance = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagImportance);
            PropertyObj pidTagPriority = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagPriority);
            PropertyObj pidTagSensitivity = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);
            PropertyObj pidTagSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubject);
            PropertyObj pidTagNormalizedSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNormalizedSubject);

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R74");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R74
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(pidTagImportance.Value),
                74,
                @"[In PidTagImportance Property] [The value 0x00000002 indicates the level of importance assigned by the end user to the Message object is] High importance.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R78");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R78
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000001,
                Convert.ToInt32(pidTagPriority.Value),
                78,
                @"[In PidTagPriority Property] [The value 0x00000001 indicates the client's request for the priority is] Urgent.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R87");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R87
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(pidTagSensitivity.Value),
                87,
                @"[In PidTagSensitivity Property] [The value 0x00000002 indicates the sender's assessment of the sensitivity of the Message object is] Private.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1802");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1802
            this.Site.CaptureRequirementIfAreEqual<string>(
                pidTagNormalizedSubject.Value.ToString(),
                pidTagSubject.Value.ToString(),
                1802,
                @"[In PidTagSubject Property] [If the PidTagSubjectPrefix property] is set to an empty string, then the values of the PidTagSubject and PidTagNormalizedSubject properties are equal.");

            PropertyObj pidTagResponsibility = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagResponsibility);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1199");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1199
            this.Site.CaptureRequirementIfIsFalse(
                Convert.ToBoolean(pidTagResponsibility.Value),
                1199,
                @"[In PidTagResponsibility Property] otherwise [another agent hasn't accepted responsibility], [property PidTagResponsibility is set to] ""FALSE"".");

            PropertyObj pidTagTrustSender = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagTrustSender);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1235");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1235
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                Convert.ToInt32(pidTagTrustSender.Value),
                1235,
                @"[In PidTagTrustSender] The value 0x00000000 indicates that Message was not delivered through a trusted transport channel.");
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetProperties to set properties of message created by step 8.
            propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject)),
                new PropertyObj(PropertyNames.PidTagPriority, BitConverter.GetBytes(0x00000000)),
                new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000003))
            };
            this.SetPropertiesForMessage(targetMessageHandle, propertyList);
            #endregion

            #region Call RopSaveChangesMessage to save message created by step 8.
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopGetPropertiesSpecific to get all properties of message created by step 8.
            propertyValues = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, propertyTagsOfMessage);

            pidTagPriority = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagPriority);
            pidTagSensitivity = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);
            pidTagSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSubject);
            pidTagNormalizedSubject = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagNormalizedSubject);

            #region Verify MS-OXCMSG_R79, MS-OXCMSG_R88, MS-OXCMSG_R1801
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R79");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R79
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000000,
                Convert.ToInt32(pidTagPriority.Value),
                79,
                @"[In PidTagPriority Property] [The value 0x00000000 indicates  the client's request for the priority is] Normal.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R88");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R88
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000003,
                Convert.ToInt32(pidTagSensitivity.Value),
                88,
                @"[In PidTagSensitivity Property] [The value 0x00000003 indicates the sender's assessment of the sensitivity of the Message object is] Confidential.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1801");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1801
            this.Site.CaptureRequirementIfAreEqual<string>(
                pidTagNormalizedSubject.Value.ToString(),
                pidTagSubject.Value.ToString(),
                1801,
                @"[In PidTagSubject Property] If the PidTagSubjectPrefix property is not set, the values of the PidTagSubject and PidTagNormalizedSubject properties are equal.");
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        #region Private methods
        /// <summary>
        /// Get a PropertyTag objects list Of Message object.
        /// </summary>
        /// <returns>A PropertyTag objects list Of Message object.</returns>
        private List<PropertyTag> GetPropertyTagListOfMessageObject()
        {
            List<PropertyTag> propertyList = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasAttachments],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageClass],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageCodepage],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageLocaleId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageFlags],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageSize],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageStatus],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSubjectPrefix],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagNormalizedSubject],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagImportance],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagPriority],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSensitivity],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAutoForwardComment],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagInternetReferences],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMimeSkeleton],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagTnefCorrelationKey],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAddressBookDisplayNamePrintable],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagCreatorEntryId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModifierEntryId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAlternateRecipientAllowed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagResponsibility],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRowid],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagHasNamedProperties],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRecipientOrder],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagPurportedSenderDomain],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagStoreEntryId],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagTrustSender],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagSubject],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagLocalCommitTime],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAutoForwarded],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAlternateRecipientAllowed],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagResponsibility],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagRecipientDisplayName]
            };

            return propertyList;
        }

        /// <summary>
        ///  Create message general  property tags for capture code  Set value
        /// </summary>
        /// <returns>Property tags array</returns>
        private List<PropertyObj> SetGeneralPropertiesOfMessage()
        {
            List<PropertyObj> propertyList = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagMessageClass, Common.GetBytesFromUnicodeString("IPM.Note")),
                new PropertyObj(PropertyNames.PidTagSubjectPrefix, Common.GetBytesFromUnicodeString(TestDataOfPidTagSubjectPrefix)),
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject)),
                new PropertyObj(PropertyNames.PidTagImportance, BitConverter.GetBytes(0x00000000)),
                new PropertyObj(PropertyNames.PidTagPriority, resultUnurgent),
                new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000001)),
                new PropertyObj(PropertyNames.PidTagAutoForwarded, BitConverter.GetBytes(true)),
                new PropertyObj(PropertyNames.PidTagAutoForwardComment, Common.GetBytesFromUnicodeString(TestDataOfPidTagAutoForwardComment)),
                new PropertyObj(PropertyNames.PidTagAddressBookDisplayNamePrintable, Common.GetBytesFromUnicodeString(TestDataOfPidTagAddressBookDisplayNamePrintable)),
                new PropertyObj(PropertyNames.PidTagRecipientOrder, BitConverter.GetBytes(0x00000001)),
                new PropertyObj(PropertyNames.PidTagPurportedSenderDomain, Common.GetBytesFromUnicodeString(Common.GetConfigurationPropertyValue("Domain", this.Site))),
                new PropertyObj(PropertyNames.PidTagAutoForwarded, BitConverter.GetBytes(false)),
                new PropertyObj(PropertyNames.PidTagAlternateRecipientAllowed, BitConverter.GetBytes(true)),
                new PropertyObj(PropertyNames.PidTagResponsibility, BitConverter.GetBytes(true)),
                new PropertyObj(PropertyNames.PidTagInternetReferences, Common.GetBytesFromUnicodeString(TestDataOfPidTagInternetReferences)),
                new PropertyObj(PropertyNames.PidTagRecipientDisplayName, Common.GetBytesFromUnicodeString(Common.GetConfigurationPropertyValue("AdminUserName", this.Site)))
            };

            return propertyList;
        }

        /// <summary>
        /// Set the value of properties identified by long ID or name in message.
        /// </summary>
        /// <param name="messageHandle">The specified message handle.</param>
        /// <param name="property">The PropertyName of specified property.</param>
        /// <param name="value">The value of specified property.</param>
        private void SetNamedProperty(uint messageHandle, PropertyNameObject property, byte[] value)
        {
            #region Call RopGetPropertyIdsFromNames to get property ID.
            PropertyName[] propertyNames = new PropertyName[1];
            propertyNames[0] = property.PropertyName;

            RopGetPropertyIdsFromNamesRequest getPropertyIdsFromNamesRequest = new RopGetPropertyIdsFromNamesRequest()
            {
                RopId = (byte)RopId.RopGetPropertyIdsFromNames,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                Flags = (byte)GetPropertyIdsFromNamesFlags.Create,
                PropertyNameCount = (ushort)propertyNames.Length,
                PropertyNames = propertyNames,
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertyIdsFromNamesRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertyIdsFromNamesResponse getPropertyIdsFromNamesResponse = (RopGetPropertyIdsFromNamesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertyIdsFromNamesResponse.ReturnValue, "Call RopGetPropertyIdsFromNames should success.");
            #endregion

            #region Set property value.

            List<TaggedPropertyValue> taggedPropertyValues = new List<TaggedPropertyValue>();

            int valueSize = 0;
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = getPropertyIdsFromNamesResponse.PropertyIds[0].ID,
                PropertyType = (ushort)property.PropertyType
            };
            TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue
            {
                PropertyTag = propertyTag,
                Value = value
            };
            valueSize += taggedPropertyValue.Size();
            taggedPropertyValues.Add(taggedPropertyValue);

            RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertyValueCount = (ushort)taggedPropertyValues.Count,
                PropertyValueSize = (ushort)(valueSize + 2),
                PropertyValues = taggedPropertyValues.ToArray()
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            RopSetPropertiesResponse rpmSetResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, rpmSetResponse.PropertyProblemCount, "If ROP succeeds, the PropertyProblemCount of its response is 0(success).");
            #endregion
        }

        #endregion
    }
}