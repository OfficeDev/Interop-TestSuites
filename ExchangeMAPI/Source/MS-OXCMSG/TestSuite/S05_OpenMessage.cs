namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that Verifies the requirements related to RopOpenMessage.
    /// </summary>
    [TestClass]
    public class S05_OpenMessage : TestSuiteBase
    {
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
        /// This test case validates opening and saving the same message in different transactions.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC01_TransactionOnMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            // Create a message in InBox
            this.MessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call SetPropertiesSpecific to set PidTagPriority property to 0x00000001 in created message.
            // Set PidTagPriority to 0x00000001
            List<PropertyObj> propertyValues = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000001))
            };
            this.SetPropertiesForMessage(this.MessageHandle, propertyValues);
            #endregion

            #region Call RopSaveChangesMessage to save the created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(this.MessageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            ulong messageId = saveChangesMessageResponse.MessageId;
            this.ReleaseRop(this.MessageHandle);
            #endregion

            #region Call RopOpenMessage to open created message on different transactions.
            uint messageHandleFirst = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.BestAccess);
            uint messageHandleSecond = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.BestAccess);

            #region Verify MS-OXCMSG_R703
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R703");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R703
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                messageHandleFirst,
                messageHandleSecond,
                703,
                @"[In Receiving a RopSaveChangesMessage ROP Request] When the server receives multiple requests to open the same Message object, it [server] returns a different handle and maintains a separate transaction (3) for each.");
            #endregion
            #endregion

            #region Call RopSetProperties to set PidTagSensitivity to 0x00000000 of created message on first transaction.
            propertyValues.Clear();
            propertyValues.Add(new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000000)));
            this.SetPropertiesForMessage(messageHandleFirst, propertyValues);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagPriority property of created message on second transaction.
            uint messageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.BestAccess);
            PropertyTag[] tagArray = new PropertyTag[1];
            tagArray[0] = PropertyHelper.PropertyTagDic[PropertyNames.PidTagSensitivity];

            // Get specific property for created message
            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the propert
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj pidTagSensitivityBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);

            this.ReleaseRop(messageHandle);

            #region Verify MS-OXCMSG_R321
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R321");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R321
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                getPropertiesSpecificResponse.ReturnValue,
                321,
                @"[In Receiving a RopOpenMessage ROP Request] The Message object returned by the RopOpenMessage ROP ([MS-OXCROPS] section 2.2.6.1) is used in subsequent ROPs, such as a RopGetPropertiesSpecific ROP request ([MS-OXCROPS] section 2.2.8.3).");
            #endregion
            #endregion

            #region Call RopSetProperties to set PidTagSensitivity to 0x00000002 of created message on second transaction.
            propertyValues.Clear();
            propertyValues.Add(new PropertyObj(PropertyNames.PidTagSensitivity, BitConverter.GetBytes(0x00000002)));
            this.SetPropertiesForMessage(messageHandleSecond, propertyValues);
            #endregion

            #region Call RopSaveChangesMessage to save the created message on first transaction.
            saveChangesMessageResponse = this.SaveMessage(messageHandleFirst, (byte)SaveFlags.KeepOpenReadWrite);

            messageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.BestAccess);

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, messageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            PropertyObj pidTagSensitivityAfterSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);
            this.ReleaseRop(messageHandle);

            Site.Assert.AreNotEqual<int>(0x00000000, Convert.ToInt32(pidTagSensitivityBeforeSave.Value), "The second transaction should not visible the changes that be generated by first transaction changed before first transaction committed.");

            Site.Assert.AreEqual<int>(0x00000000, Convert.ToInt32(pidTagSensitivityAfterSave.Value), "The second transaction should not visible the changes that be generated by first transaction changed after first transaction committed.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R704");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R704
            // Because the above steps have verified that any changes made on one transaction MUST NOT be visible to another transaction until the changes are committed via the RopSaveChangesMessage ROP.
            // R704 will be direct verified.
            this.Site.CaptureRequirement(
                704,
                @"[In Receiving a RopSaveChangesMessage ROP Request] Any changes made on one transaction (3) MUST NOT be visible to another transaction (3) until the changes are committed via the RopSaveChangesMessage ROP.");
            #endregion

            #region Call RopSaveChangesMessage to save the created message on second transaction with not force save.
            if (Common.IsRequirementEnabled(1643, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(messageHandleSecond, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1643");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1643
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1643,
                    @"[In Appendix A: Product Behavior] Implementation does return Success for RopSaveChangesMessage ROP requests ([MS-OXCROPS] section 2.2.6.3) when a previous request has already been committed against the Message object, even though the changes to the object are not actually committed to the server message store. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1916, this.Site))
            {
                saveChangesMessageResponse = this.SaveMessage(messageHandleSecond, (byte)SaveFlags.KeepOpenReadWrite);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1916");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1916
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    TestSuiteBase.Success,
                    saveChangesMessageResponse.ReturnValue,
                    1916,
                    @"[In Appendix A: Product Behavior] Implementation doesn't return Success for RopSaveChangesMessage ROP requests when a previous request has already been committed against the Message object, even though the changes to the object are not actually committed to the server message store. (Exchange 2007 follows this behavior.)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1053");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R1053
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040109,
                    saveChangesMessageResponse.ReturnValue,
                    1053,
                    @"[In Receiving a RopSaveChangesMessage ROP Request] The value of the ecObjectModified is 0x80040109, which indicates that the underlying data for this Message object was changed through another transaction (3) context.");
            }
            #endregion

            #region Call RopSaveChangesMessage to save the created message on second transaction with force save.
            saveChangesMessageResponse = this.SaveMessage(messageHandleSecond, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopSaveChangesMessage should success.");

            #region Call RopGetPropertiesSpecific to get the PidTagPriority property of created message on second transaction.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, messageHandleSecond, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, "Call RopGetPropertiesSpecific should success.");

            propertyValues = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);
            pidTagSensitivityBeforeSave = PropertyHelper.GetPropertyByName(propertyValues, PropertyNames.PidTagSensitivity);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R724");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R724
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000002,
                Convert.ToInt32(pidTagSensitivityBeforeSave.Value),
                724,
                @"[In RopSaveChangesMessage ROP Request Buffer] [SaveFlags] [ForceSave (0x04)] The ecObjectModified error code is not valid when this flag [ForceSave] is set; the server overwrites any changes instead.");
            #endregion
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(messageHandleFirst);
            this.ReleaseRop(messageHandleSecond);
            #endregion
        }

        /// <summary>
        /// This test case validates that the RopOpenMessage ROP ignores the invalid OpenModeFlags value.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC02_IgnoreOpenModeFlagsOnRopOpenMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage and OpenModeFlags is 0x08 to open the created message.
            RopOpenMessageRequest openMessageRequestFirst = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Open the message in INBOX folder in which message is created.
                OpenModeFlags = 0x08, // The OpenModeFlags other bit are set, server should ignore this bit.
                MessageId = saveChangesMessageResponse.MessageId // Open the saved message
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequestFirst, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponseFirst = (RopOpenMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponseFirst.ReturnValue, "Call RopOpenMessage should success.");

            uint openedMessageHandleFirst = this.ResponseSOHs[0][openMessageResponseFirst.OutputHandleIndex];
            #endregion

            #region Call RopOpenMessage and OpenModeFlags is 0x0B to open the created message.
            RopOpenMessageRequest openMessageRequestSecond = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Open the message in INBOX folder in which message is created.
                OpenModeFlags = 0x0B, // The OpenModeFlags other bit are set, server should ignore this bit.
                MessageId = saveChangesMessageResponse.MessageId // Open the saved message
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequestSecond, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponseSecond = (RopOpenMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponseSecond.ReturnValue, "Call RopOpenMessage should success.");

            uint openedMessageHandleSecond = this.ResponseSOHs[0][openMessageResponseSecond.OutputHandleIndex];
            #endregion

            #region Verify MS-OXCMSG_R1838, MS-OXCMSG_R1475
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1838");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1838
            this.Site.CaptureRequirementIfAreEqual<RopOpenMessageResponse>(
                openMessageResponseFirst,
                openMessageResponseSecond,
                1838,
                @"[In RopOpenMessage ROP Request Buffer] The server responses are same when OpenModeFlags are set with two different values, and the two different values are not one of 0x00, 0x01, 0x03 and 0x04.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1475");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1475
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                openedMessageHandleFirst,
                openedMessageHandleSecond,
                1475,
                @"[In Receiving a RopOpenMessage ROP Request] When the server receives multiple requests to open the same Message object, it returns a different handle and maintains a separate transaction (3) for each.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(openedMessageHandleFirst);
            this.ReleaseRop(openedMessageHandleSecond);
            #endregion
        }

        /// <summary>
        /// This test case validates the error code ecNotFound (0x8004010F) in RopOpenMessage ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC03_RopOpenMessageFailure()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong targetMessageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopOpenMessage to open a message which does not exist.
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Open the message in INBOX folder in which message is created.
                OpenModeFlags = 0x00, // The message will be opened as read-only
                MessageId = targetMessageId - 1 // The specified ID does not exist. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.response;

            #region Verify MS-OXCMSG_R323, MS-OXCMSG_R1477 and MS-OXCMSG_R333
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R323");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R323
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.Success,
                openMessageResponse.ReturnValue,
                323,
                @"[In Receiving a RopOpenMessage ROP Request] A RopOpenMessage ROP MUST NOT succeed if a Message object with the specified ID does not exist.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1477");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1477
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openMessageResponse.ReturnValue,
                1477,
                @"[In Receiving a RopOpenMessage ROP Request] [ecNotFound (0x8004010F)] The folder corresponding to the FID ([MS-OXCDATA] section 2.2.1.1) entered in the ROP request buffer does not contain a message with the entered MID.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R333");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R333
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openMessageResponse.ReturnValue,
                333,
                @"[In Receiving a RopOpenMessage ROP Request] [ecNotFound (0x8004010F)] The MID ([MS-OXCDATA] section 2.2.1.2) does not correspond to a message in the database.");
            #endregion
            #endregion

            #region Call RopOpenMessage to open a message and the value of the InputHandleIndex field on which this ROP was called does not refer to a Folder object or a Store object.
            openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonOutputHandleIndex, // the value of the InputHandleIndex field does not refer to a Folder object or a Store object.
                OutputHandleIndex = CommonOutputHandleIndex, 
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Open the message in INBOX folder in which message is created.
                OpenModeFlags = 0x00, // The message will be opened as read-only
                MessageId = targetMessageId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openMessageResponse = (RopOpenMessageResponse)this.response;

            #region Verify MS-OXCMSG_R1478
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1478");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1478
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004b9,
                openMessageResponse.ReturnValue,
                1478,
                @"[In Receiving a RopOpenMessage ROP Request] ecNullObject (0x000004b9)] The value of the InputHandleIndex field on which this ROP was called does not refer to a Folder object or a Store object");
            #endregion
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case validates opening the soft delete message by using RopOpenMessage ROP.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC04_OpenSoftDeletedMessage()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open inbox folder
            uint openedFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create Message object.
            uint targetMessageHandleForDelete = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandleForDelete, (byte)SaveFlags.ForceSave);
            ulong messageIdForDelete = saveChangesMessageResponse.MessageId;
            #endregion

            this.ReleaseRop(targetMessageHandleForDelete);

            #region Call RopDeleteMessages to soft delete created message.
            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest()
            {
                RopId = (byte)RopId.RopDeleteMessages,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                WantAsynchronous = 0x00,

                // The server does not generate a non-read receipt for the deleted messages.
                NotifyNonRead = 0x00,
                MessageIdCount = 1,
                MessageIds = new ulong[] { messageIdForDelete },
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteMessagesRequest, openedFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteMessagesResponse deleteMessagesResponse = (RopDeleteMessagesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteMessagesResponse.ReturnValue, "Call RopDeleteMessages should success.");
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is not OpenSoftDeleted to open soft deleted message.
            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = logonResponse.FolderIds[4], // Open the message in INBOX folder in which message is created.
                OpenModeFlags = (byte)MessageOpenModeFlags.ReadWrite,
                MessageId = messageIdForDelete // Open the saved message
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.response;

            #region Verify MS-OXCMSG_R335, MS-OXCMSG_R326
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R335");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R335
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                openMessageResponse.ReturnValue,
                335,
                @"[In Receiving a RopOpenMessage ROP Request] [ecNotFound (0x8004010F)] The message is soft deleted and the client has not specified the OpenSoftDeleted flag as part of the OpenModeFlag field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R326");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R326
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.Success,
                openMessageResponse.ReturnValue,
                326,
                @"[In Receiving a RopOpenMessage ROP Request] If the OpenSoftDeleted flag is not included, the server MUST NOT provide access to soft deleted Message objects.");
            #endregion
            #endregion

            #region Call RopRelease to release the created message
            this.ReleaseRop(openedFolderHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests RopOpenMessage with OpenModeFlags set to ReadOnly.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC05_RopOpenMessageAsReadOnly()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyObj> propertyValues = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject))
            };

            #region Call RopLogon to logon the private mailbox with administrator.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is read-only to open created message.
            uint openedMessageReadOnlyHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.ReadOnly);
            #endregion

            #region Call RopSetProperties to set PidTagNormalizedSubject property of created message.
            if (Common.IsRequirementEnabled(661, this.Site))
            {
                List<TaggedPropertyValue> taggedPropertyValueList = new List<TaggedPropertyValue>();

                int valueSize = 0;
                foreach (PropertyObj propertyObj in propertyValues)
                {
                    PropertyTag propertyTag = new PropertyTag
                    {
                        PropertyId = (ushort)propertyObj.PropertyID,
                        PropertyType = (ushort)propertyObj.ValueTypeCode
                    };

                    TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue
                    {
                        PropertyTag = propertyTag,
                        Value = (byte[])propertyObj.Value
                    };
                    valueSize += taggedPropertyValue.Size();

                    taggedPropertyValueList.Add(taggedPropertyValue);
                }

                RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
                {
                    RopId = (byte)RopId.RopSetProperties,
                    LogonId = CommonLogonId,
                    InputHandleIndex = CommonInputHandleIndex,
                    PropertyValueCount = (ushort)taggedPropertyValueList.Count,
                    PropertyValueSize = (ushort)(valueSize + 2),
                    PropertyValues = taggedPropertyValueList.ToArray()
                };
                uint returnValue;
                this.response = null;
                this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, openedMessageReadOnlyHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None, out returnValue);

                #region Verify MS-OXCMSG_R661
                if (this.response == null)
                {
                    Site.Assert.AreNotEqual<uint>(0, returnValue, "Call RopSetProperties should fail.");
                }
                else
                {
                    RopSetPropertiesResponse rpmSetResponse = (RopSetPropertiesResponse)this.response;
                    Site.Assert.AreNotEqual<uint>(0, rpmSetResponse.ReturnValue, "Call RopSetProperties should fail.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R661");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R661
                // Because when call RopSetProperties to set property of specified Message object that opened by read-only, server will fail and return error code.
                // R661 will be verified.
                this.Site.CaptureRequirement(
                    661,
                    @"[In RopOpenMessage ROP Request Buffer] [OpenModeFlags] [ReadOnly (0x00)] Message will be opened as read only.");
                #endregion
            }
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(openedMessageReadOnlyHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests RopOpenMessage with OpenModeFlags set to ReadWrite.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC06_RopOpenMessageAsReadWrite()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox with administrator.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            List<PropertyObj> propertyValues = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject))
            };

            #region Call RopCreateMessage to create a Message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is ReadWrite to open created message.
            uint openedMessageReadWriteHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite);
            #endregion

            #region Call RopSetProperties to set PidTagNormalizedSubject property of created message.
            List<TaggedPropertyValue> taggedPropertyValueList = new List<TaggedPropertyValue>();

            int valueSize = 0;
            foreach (PropertyObj propertyObj in propertyValues)
            {
                PropertyTag propertyTag = new PropertyTag
                {
                    PropertyId = (ushort)propertyObj.PropertyID,
                    PropertyType = (ushort)propertyObj.ValueTypeCode
                };

                TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue
                {
                    PropertyTag = propertyTag,
                    Value = (byte[])propertyObj.Value
                };
                valueSize += taggedPropertyValue.Size();

                taggedPropertyValueList.Add(taggedPropertyValue);
            }

            RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertyValueCount = (ushort)taggedPropertyValueList.Count,
                PropertyValueSize = (ushort)(valueSize + 2),
                PropertyValues = taggedPropertyValueList.ToArray()
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, openedMessageReadWriteHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetPropertiesResponse rpmSetResponse = (RopSetPropertiesResponse)this.response;
            #region Verify MS-OXCMSG_R663
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R663");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R663
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                rpmSetResponse.ReturnValue,
                663,
                @"[In RopOpenMessage ROP Request Buffer] [OpenModeFlags] [ReadWrite (0x01)] Message will be opened for both reading and writing.");
            #endregion
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(openedMessageReadWriteHandle);
            #endregion
        }

        /// <summary>
        /// This test case tests RopOpenMessage with OpenModeFlags set to BestAccess.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC07_RopOpenMessageAsBestAccess()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to logon the private mailbox with administrator.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            List<PropertyObj> propertyValues = new List<PropertyObj>
            {
                new PropertyObj(PropertyNames.PidTagNormalizedSubject, Common.GetBytesFromUnicodeString(TestDataOfPidTagNormalizedSubject))
            };
            string commonUser = Common.GetConfigurationPropertyValue("CommonUser", Site);
            string commonUserPassword = Common.GetConfigurationPropertyValue("CommonUserPassword", Site);
            string commonUserEssdn = Common.GetConfigurationPropertyValue("CommonUserEssdn", Site);
            uint pidTagMemberRights;
            RopSetPropertiesResponse rpmSetResponse;

            #region Call RopOpenFolder to open the inbox folder.
            ulong parentFolderId = logonResponse.FolderIds[4];
            uint openedInboxFolderHandle = this.OpenSpecificFolder(parentFolderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateFolder to create a subfolder in inbox folder.
            ulong firstSubfolderId;
            uint firstSubFolderHandle = this.CreateSubFolder(openedInboxFolderHandle, out firstSubfolderId);
            LongTermId firstSubfolderLongTermID = this.GetLongTermIdFormID(firstSubfolderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage and RopSaveChangesMessage to create a Message object in subfolder created by step 3.
            // Create a message in InBox
            this.MessageHandle = this.CreatedMessage(firstSubfolderId, this.insideObjHandle);

            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(this.MessageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            ulong firstMessageId = saveChangesMessageResponse.MessageId;
            LongTermId firstMessageLongTermID = this.GetLongTermIdFormID(firstMessageId, this.insideObjHandle);
            this.ReleaseRop(this.MessageHandle);
            #endregion

            #region Call RopCreateFolder to create a new subfolder.
            ulong secondSubfolderId;
            System.Threading.Thread.Sleep(1); // Sleep 1 millisecond to generate different named folder
            uint secondSubFolderHandle = this.CreateSubFolder(openedInboxFolderHandle, out secondSubfolderId);
            LongTermId secondSubfolderLongTermID = this.GetLongTermIdFormID(secondSubfolderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage and RopSaveChangesMessage to create a Message object in subfolder created by step 5.
            // Create a message in InBox
            this.MessageHandle = this.CreatedMessage(secondSubfolderId, this.insideObjHandle);

            saveChangesMessageResponse = this.SaveMessage(this.MessageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            ulong secondMessageId = saveChangesMessageResponse.MessageId;
            LongTermId secondMessageLongTermID = this.GetLongTermIdFormID(secondMessageId, this.insideObjHandle);
            this.ReleaseRop(this.MessageHandle);
            #endregion

            #region Add Read permission to "CommonUser" on inbox folder.
            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible | (uint)PidTagMemberRights.ReadAny;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, openedInboxFolderHandle);
            #endregion

            #region Add Read permission to "CommonUser" on subfolder created by step 3.
            // Add folder visible permission for the subfolder1.
            pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible | (uint)PidTagMemberRights.ReadAny;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, firstSubFolderHandle);
            #endregion

            #region Add Read and write permission to "CommonUser" on subfolder by step 5
            pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible | (uint)PidTagMemberRights.ReadAny | (uint)PidTagMemberRights.EditAny;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, secondSubFolderHandle);
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
                Essdn = Encoding.ASCII.GetBytes(userDN),
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(logonRequest, this.insideObjHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            logonResponse = (RopLogonResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, logonResponse.ReturnValue, "Call RopLogon should success.");

            uint objHandle = this.ResponseSOHs[0][logonResponse.OutputHandleIndex];
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is BestAccess to open the message in subfolder created by step 3.
            firstSubfolderId = this.GetObjectIdFormLongTermID(firstSubfolderLongTermID, objHandle);
            firstMessageId = this.GetObjectIdFormLongTermID(firstMessageLongTermID, objHandle);

            uint openedMessageBestReadOnlyHandle = this.OpenSpecificMessage(firstSubfolderId, firstMessageId, objHandle, MessageOpenModeFlags.BestAccess);
            #endregion

            #region Call RopSetProperties to set PidTagNormalizedSubject property of message in subfolder created by step 3.
            List<TaggedPropertyValue> taggedPropertyValueList = new List<TaggedPropertyValue>();

            int valueSize = 0;
            foreach (PropertyObj propertyObj in propertyValues)
            {
                PropertyTag propertyTag = new PropertyTag
                {
                    PropertyId = (ushort)propertyObj.PropertyID,
                    PropertyType = (ushort)propertyObj.ValueTypeCode
                };

                TaggedPropertyValue taggedPropertyValue = new TaggedPropertyValue
                {
                    PropertyTag = propertyTag,
                    Value = (byte[])propertyObj.Value
                };
                valueSize += taggedPropertyValue.Size();

                taggedPropertyValueList.Add(taggedPropertyValue);
            }

            RopSetPropertiesRequest rpmSetRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                PropertyValueCount = (ushort)taggedPropertyValueList.Count,
                PropertyValueSize = (ushort)(valueSize + 2),
                PropertyValues = taggedPropertyValueList.ToArray(),
            };
            uint returnValue;
            this.response = null;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, openedMessageBestReadOnlyHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None, out returnValue);
            
            if(Common.IsRequirementEnabled(3009,this.Site))
            {
                rpmSetResponse = (RopSetPropertiesResponse)this.response;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3009");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3009
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    rpmSetResponse.ReturnValue,
                    3009,
                    @"[In Appendix A: Product Behavior] BestAccess is read/write if the user don't have write permissions. (<12> Section 2.2.3.1.1:  Exchange 2010 and above follow this behavior.)");
            }

            if(Common.IsRequirementEnabled(3010,this.Site))
            {
                if (this.response == null)
                {
                    Site.Assert.AreNotEqual<uint>(0, returnValue, "Call RopSetProperties should fail.");
                }
                else
                {
                    rpmSetResponse = (RopSetPropertiesResponse)this.response;
                    Site.Assert.AreNotEqual<uint>(0, rpmSetResponse.ReturnValue, "Call RopSetProperties should fail.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3010");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R661
                // Because when call RopSetProperties to set property of specified Message object that opened by read-only, server will fail and return error code.
                // R661 will be verified.
                this.Site.CaptureRequirement(
                    3010,
                    @"[In Appendix A: Product Behavior] BestAccess is read-only if the user don't have write permissions. (Exchange 2007 follows this behavior.)");
            }
            #endregion

            #region Call RopRelease to release created message in subfolder created by step 3.
            this.ReleaseRop(openedMessageBestReadOnlyHandle);
            #endregion

            #region Call RopOpenMessage which OpenModeFlags is BestAccess to open the message in subfolder created by step 5.
            secondSubfolderId = this.GetObjectIdFormLongTermID(secondSubfolderLongTermID, objHandle);
            secondMessageId = this.GetObjectIdFormLongTermID(secondMessageLongTermID, objHandle);
            uint openedMessageBestReadWrietHandle = this.OpenSpecificMessage(secondSubfolderId, secondMessageId, objHandle, MessageOpenModeFlags.BestAccess);
            #endregion

            #region Call RopSetProperties to set PidTagSubjectPrefix property of message in subfolder created by step 5.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(rpmSetRequest, openedMessageBestReadWrietHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            rpmSetResponse = (RopSetPropertiesResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R665");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R665
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                rpmSetResponse.ReturnValue,
                665,
                @"[In RopOpenMessage ROP Request Buffer] [OpenModeFlags] [BestAccess (0x03)] Open for read/write if the user has write permissions for the folder.");
            #endregion

            #region Call RopRelease to release created message in subfolder created by step 5.
            this.ReleaseRop(openedMessageBestReadWrietHandle);
            #endregion

            #region Call RopLogon to logon the private mailbox with administrator
            this.rawData = null;
            this.insideObjHandle = 0;
            this.response = null;
            this.ResponseSOHs = null;
            this.MSOXCMSGAdapter.RpcDisconnect();
            this.MSOXCMSGAdapter.Reset();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            openedInboxFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopDeleteFolder to delete the subfolder created by 3.
            firstSubfolderId = this.GetObjectIdFormLongTermID(firstSubfolderLongTermID, this.insideObjHandle);
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest()
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete | (byte)DeleteFolderFlags.DelMessages,
                FolderId = firstSubfolderId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteFolderRequest, openedInboxFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteFolderResponse deleteFolderresponse = (RopDeleteFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteFolderresponse.ReturnValue, "Call RopDeleteFolder should success.");
            #endregion

            #region Call RopDeleteFolder to delete the subfolder created by 5
            secondSubfolderId = this.GetObjectIdFormLongTermID(secondSubfolderLongTermID, this.insideObjHandle);
            deleteFolderRequest.FolderId = secondSubfolderId;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteFolderRequest, openedInboxFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            deleteFolderresponse = (RopDeleteFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteFolderresponse.ReturnValue, "Call RopDeleteFolder should success.");
            #endregion

            this.ReleaseRop(openedInboxFolderHandle);
        }

        /// <summary>
        /// This test case tests that user does not have rights to open message.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S05_TC08_RopOpenMessageWithoutRight()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            string commonUser = Common.GetConfigurationPropertyValue("CommonUser", Site);
            string commonUserPassword = Common.GetConfigurationPropertyValue("CommonUserPassword", Site);
            string commonUserEssdn = Common.GetConfigurationPropertyValue("CommonUserEssdn", Site);
            uint pidTagMemberRights;

            #region Call RopLogon to logon the private mailbox with administrator.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopOpenFolder to open the inbox folder.
            ulong parentFolderId = logonResponse.FolderIds[4];
            uint openedInboxFolderHandle = this.OpenSpecificFolder(parentFolderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateFolder to create a new subfolder.
            ulong thirdSubfolderId;
            uint thirdSubFolderHandle = this.CreateSubFolder(openedInboxFolderHandle, out thirdSubfolderId);
            LongTermId thirdSubfolderLongTermID = this.GetLongTermIdFormID(thirdSubfolderId, this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage and RopSaveChangesMessage to create a Message object in subfolder created.
            // Create a message in InBox
            this.MessageHandle = this.CreatedMessage(thirdSubfolderId, this.insideObjHandle);

            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(this.MessageHandle, (byte)SaveFlags.KeepOpenReadWrite);
            ulong thirdMessageId = saveChangesMessageResponse.MessageId;
            LongTermId thirdMessageLongTermID = this.GetLongTermIdFormID(thirdMessageId, this.insideObjHandle);
            this.ReleaseRop(this.MessageHandle);
            #endregion

            #region Add Read permission to "CommonUser" on inbox folder.
            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible | (uint)PidTagMemberRights.ReadAny;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, openedInboxFolderHandle);
            #endregion

            #region Add Read and write permission to "CommonUser" on subfolder
            pidTagMemberRights = (uint)PidTagMemberRights.FolderVisible;
            this.AddPermission(commonUserEssdn, pidTagMemberRights, thirdSubFolderHandle);
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

            #region Call RopOpenMessage to open a message that the user does not have rights to the message.
            thirdSubfolderId = this.GetObjectIdFormLongTermID(thirdSubfolderLongTermID, objHandle);
            thirdMessageId = this.GetObjectIdFormLongTermID(thirdMessageLongTermID, objHandle);

            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                OutputHandleIndex = CommonOutputHandleIndex,
                CodePageId = 0x0FFF, // Code page of Logon object is used
                FolderId = thirdSubfolderId,
                OpenModeFlags = (byte)MessageOpenModeFlags.ReadWrite,
                MessageId = thirdMessageId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openMessageRequest, objHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenMessageResponse openMessageResponse = (RopOpenMessageResponse)this.response;

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R324");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R324
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                TestSuiteBase.Success,
                openMessageResponse.ReturnValue,
                324,
                @"[in Receiving a RopOpenMessage ROP Request] RopOpenMessage MUST NOT succeed if the client has insufficient access rights to the folder in which the Message object is stored.");
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2184");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2184
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                openMessageResponse.ReturnValue,
                2184,
                @"[In Receiving a RopOpenMessage ROP Request] [ecAccessDenied(0x80070005)] The user does not have rights to the message.");
            #endregion
            #endregion

            #region Call RopLogon to logon the private mailbox with administrator
            this.rawData = null;
            this.insideObjHandle = 0;
            this.response = null;
            this.ResponseSOHs = null;
            this.MSOXCMSGAdapter.RpcDisconnect();
            this.MSOXCMSGAdapter.Reset();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);
            logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            openedInboxFolderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopDeleteFolder to delete the subfolder created
            thirdSubfolderId = this.GetObjectIdFormLongTermID(thirdSubfolderLongTermID, this.insideObjHandle);
            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest()
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete | (byte)DeleteFolderFlags.DelMessages,
                FolderId = thirdSubfolderId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(deleteFolderRequest, openedInboxFolderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopDeleteFolderResponse deleteFolderresponse = (RopDeleteFolderResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, deleteFolderresponse.ReturnValue, "Call RopDeleteFolder should success.");
            #endregion
        }
    }
}