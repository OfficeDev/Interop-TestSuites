namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test Rops related to EmbeddedMessage  
    /// </summary>
    [TestClass]
    public class S09_RopEmbeddedMessage : TestSuiteBase
    {
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

        /// <summary>
        /// This case is used to test positive response of RopOpenEmbeddedMessage.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S09_TC01_RopOpenEmbeddedMessageSuccessfully()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            // Set property size 
            int size = 0;
            TaggedPropertyValue[] taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArrays(out size, PidTagAttachMethodFlags.afEmbeddedMessage);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopCreateAttachment to create an embedded attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(targetMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod property, that is the attachment is the embedded attachment.
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopSetProperties.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertyValueSize = (ushort)(size + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setPropertiesRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagAttachMethod of created Attachment
            List<PropertyTag> tagArray = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagAttachMethod]
            };

            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse;
            getPropertiesSpecificResponse = this.GetSpecificPropertiesOfMessage(attachmentHandle, tagArray);
            List<PropertyObj> pts = PropertyHelper.GetPropertyObjFromBuffer(tagArray.ToArray(), getPropertiesSpecificResponse);

            // Parse property response get Property Value to verify test  case requirement
            PropertyObj pidTagAttachMethod = PropertyHelper.GetPropertyByName(pts, PropertyNames.PidTagAttachMethod);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R596");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R596
            this.Site.CaptureRequirementIfAreEqual<int>(
                0x00000005,
                Convert.ToInt32(pidTagAttachMethod.Value),
                596,
                @"[In PidTagAttachMethod Property] [afEmbeddedMessage (0x00000005)] The attachment is an embedded message that is accessed via the RopOpenEmbeddedMessage ROP ([MS-OXCROPS] section 2.2.6.16).");

            #region Call RopSaveChangesAttachment to save the attachment changes.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 to create the attachment if it doesn't exist, and expect to get a successful response
            RopOpenEmbeddedMessageRequest openEmbeddedMessageRequest = new RopOpenEmbeddedMessageRequest()
            {
                RopId = (byte)RopId.RopOpenEmbeddedMessage,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopOpenEmbeddedMessage.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                CodePageId = 0x0FFF, // Code page of Logon object is used
                OpenModeFlags = 0x02 // Create the attachment if it does not already exist and open the message for both reading and writing
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openEmbeddedMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            uint embeddedMessageHandle = this.ResponseSOHs[0][openEmbeddedMessageResponse.OutputHandleIndex];
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            saveChangesMessageResponse = this.SaveMessage(embeddedMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R913");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R913
            // The embedded message hasn't been created, so MS-OXCMSG_R913 can be verified if the response of calling RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 is successful.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                openEmbeddedMessageResponse.ReturnValue,
                913,
                @"[In RopOpenEmbeddedMessage ROP Request Buffer] [OpenModeFlags] [Create (0x02)] Create the attachment if it does not already exist and open the message for both reading and writing.");

            #region Call RopRelease to release the embedded message
            this.ReleaseRop(embeddedMessageHandle);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x00 to open the embedded message as read-only.
            openEmbeddedMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            openEmbeddedMessageRequest.OpenModeFlags = 0x00; // Open the message as read-only.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openEmbeddedMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            embeddedMessageHandle = this.ResponseSOHs[0][openEmbeddedMessageResponse.OutputHandleIndex];
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R881");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R881
            // MS-OXCMSG_R881 can be verified when the handle returned from calling RopOpenEmbeddedMessage is not null.
            this.Site.CaptureRequirementIfIsNotNull(
                embeddedMessageHandle,
                881,
                @"[In RopOpenEmbeddedMessage ROP] The RopOpenEmbeddedMessage ROP ([MS-OXCROPS] section 2.2.6.16) retrieves a handle to a Message object from the given Attachment object.");

            #region Call RopModifyRecipients to add recipient to the read-only embedded message and expect to get the failure response
            // Initialize TestUser1
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(TestUser1, 0)
            };

            RopModifyRecipientsResponse modifyRecipientsResponse;
            this.AddRecipients(modifyRecipientRow, embeddedMessageHandle, propertyTag, out modifyRecipientsResponse);
            #endregion

            if (Common.IsRequirementEnabled(3014, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3014");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3014
                // The response of RopModifyRecipients is not successful because the embedded message is opened as read-only, so MS-OXCMSG_R3014 can be verified.
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    TestSuiteBase.Success,
                    modifyRecipientsResponse.ReturnValue,
                    3014,
                    @"[In Appendix A: Product Behavior] [OpenModeFlags] [ReadOnly (0x00)] Message will be opened as read only. (Exchange 2007 follows this behavior.)");
            }

            if(Common.IsRequirementEnabled(3013,this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R3013");

                // Verify MS-OXCMSG requirement: MS-OXCMSG_R3013
                // The response of RopModifyRecipients is successful because the embedded message is opened as read/write, so MS-OXCMSG_R3013 can be verified.
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    TestSuiteBase.Success,
                    modifyRecipientsResponse.ReturnValue,
                    3013,
                    @"[In Appendix A: Product Behavior]  [OpenModeFlags] [ReadOnly (0x00)] Message will be opened as read/write. (<17> Section 2.2.3.16.1:  Exchange 2010, Exchange 2013, and Exchange 2016 follow this behavior.)");
            }

            #region Call RopRelease to release the embedded message.
            this.ReleaseRop(embeddedMessageHandle);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x01 to open the embedded message as read/write
            openEmbeddedMessageRequest.InputHandleIndex = 0x00;
            openEmbeddedMessageRequest.OpenModeFlags = 0x01; // Open the message for both reading and writing.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            embeddedMessageHandle = this.ResponseSOHs[0][openEmbeddedMessageResponse.OutputHandleIndex];
            #endregion

            #region Call RopModifyRecipients to add recipient to the read/write embedded message and expect to get the successful response.
            this.AddRecipients(modifyRecipientRow, embeddedMessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R912");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R912
            // The response of RopModifyRecipients is success because the embedded message is opened as read/write, so MS-OXCMSG_R912 can be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                modifyRecipientsResponse.ReturnValue,
                912,
                @"[In RopOpenEmbeddedMessage ROP Request Buffer] [OpenModeFlags] [ReadWrite (0x01)] Message will be opened for both reading and writing.");

            #region Call RopRelease to release the created message and the created attachment.
            this.ReleaseRop(embeddedMessageHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test negative response of RopOpenEmbeddedMessage.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S09_TC02_RopOpenEmbeddedMessageFailed()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            // Set property size 
            int size = 0;
            TaggedPropertyValue[] taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArrays(out size, PidTagAttachMethodFlags.afEmbeddedMessage);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an embedded attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod property, that is the attachment is the embedded attachment.
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopSetProperties.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertyValueSize = (ushort)(size + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setPropertiesRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment changes.
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 to create the attachment if it doesn't exist, and expect to get a successful response
            RopOpenEmbeddedMessageRequest openEmbeddedMessageRequest = new RopOpenEmbeddedMessageRequest()
            {
                RopId = (byte)RopId.RopOpenEmbeddedMessage,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopOpenEmbeddedMessage.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                CodePageId = 0x0FFF, // Code page of Logon object is used
                OpenModeFlags = 0x02 // Create the attachment if it does not already exist and open the message for both reading and writing
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openEmbeddedMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            uint embeddedMessageHandle = this.ResponseSOHs[0][openEmbeddedMessageResponse.OutputHandleIndex];
            #endregion

            #region Call RopRelease to release the embedded message
            this.ReleaseRop(embeddedMessageHandle);
            #endregion

            #region Call RopOpenEmbeddedMessage with CodePageId set to 0x000F and error code 0x000003ef is expected
            openEmbeddedMessageRequest.CodePageId = 0x000F;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1068");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1068
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000003ef,
                openEmbeddedMessageResponse.ReturnValue,
                1068,
                @"[In Receiving a RopOpenEmbeddedMessage ROP Request] [ecUnknownCodePage (0x000003ef)] The code page is unknown.");

            #region Call RopOpenEmbeddedMessage with InputHandleIndex set to 0x01 to open the embedded message and expect to get the error code 0x000004B9 (ecNullObject).
            openEmbeddedMessageRequest.CodePageId = 0x0FFF; // Code page of Logon object is used
            openEmbeddedMessageRequest.InputHandleIndex = 0x01;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1525");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1525
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                openEmbeddedMessageResponse.ReturnValue,
                1525,
                @"[In Receiving a RopOpenEmbeddedMessage ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopOpenEmbeddedMessage] was called does not refer to an Attachment object.");

            #region Call RopRelease to release the embedded message.
            this.ReleaseRop(embeddedMessageHandle);
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test transaction in embedded message. 
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S09_TC03_Transaction()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            // Set property size 
            int size = 0;
            TaggedPropertyValue[] taggedPropertyValueArray = this.CreateMessageTaggedPropertyValueArrays(out size, PidTagAttachMethodFlags.afEmbeddedMessage);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the newly created message.
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
             #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopCreateAttachment to create an embedded attachment.
            RopCreateAttachmentResponse createAttachmentResponse;
            uint attachmentId;
            uint attachmentHandle = this.CreateAttachment(openedMessageHandle, out createAttachmentResponse, out attachmentId);
            #endregion

            #region Call RopSetProperties to set PidTagAttachMethod property, that is the attachment is the embedded attachment.
            // Setting PidTagAttachMethod property means the attachment is the embedded message
            RopSetPropertiesRequest setPropertiesRequest = new RopSetPropertiesRequest()
            {
                RopId = (byte)RopId.RopSetProperties,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopSetProperties.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertyValueSize = (ushort)(size + 2),
                PropertyValueCount = (ushort)taggedPropertyValueArray.Length,
                PropertyValues = taggedPropertyValueArray
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setPropertiesRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetPropertiesResponse setPropertiesResponse = (RopSetPropertiesResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setPropertiesResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSaveChangesAttachment to save the attachment changes
            RopSaveChangesAttachmentResponse saveChangesAttachmentResponse;
            this.SaveAttachment(attachmentHandle, out saveChangesAttachmentResponse);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x02 to create an embedded message.
            RopOpenEmbeddedMessageRequest openEmbeddedMessageRequest = new RopOpenEmbeddedMessageRequest()
            {
                RopId = (byte)RopId.RopOpenEmbeddedMessage,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopOpenEmbeddedMessage.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                OutputHandleIndex = CommonOutputHandleIndex, // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                CodePageId = 0x0FFF, // Code page of Logon object is used
                OpenModeFlags = 0x02 // Create the attachment if it does not already exist and open the message for both reading and writing
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.response;
            uint embeddedMessageHandle = this.ResponseSOHs[0][openMessageResponse.OutputHandleIndex];
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openEmbeddedMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x01 to open the embedded message which is created in step 9 and expect a failure response
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponseFirst;
            openEmbeddedMessageRequest.OpenModeFlags = 0x01; // Try to open the embedded message as read/write.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponseFirst = (RopOpenEmbeddedMessageResponse)this.response;
            #endregion

            #region Call RopSaveChangesMessage to save the embedded message created in step 9.
            saveChangesMessageResponse = this.SaveMessage(embeddedMessageHandle, (byte)SaveFlags.ForceSave);
            #endregion

            #region Call ReleaseRop to release the embedded message created in step 9.
            this.ReleaseRop(embeddedMessageHandle);
            #endregion

            #region Call RopOpenEmbeddedMessage with OpenModeFlags set to 0x01 to open the embedded message which is created in step 9 and expect a successful response.
            RopOpenEmbeddedMessageResponse openEmbeddedMessageResponseSecond;
            openEmbeddedMessageRequest.OpenModeFlags = 0x01; // Try to open the embedded message as read/write.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(openEmbeddedMessageRequest, attachmentHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            openEmbeddedMessageResponseSecond = (RopOpenEmbeddedMessageResponse)this.response;
            #endregion

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R481. The return value of RopOpenEmbeddedMessage before calling RopSaveChangesMessage is {0}, the return value of RopOpenEmbeddedMessage after calling RopSaveChangesMessage is {1}.", openEmbeddedMessageResponseFirst.ReturnValue, openEmbeddedMessageResponseSecond.ReturnValue);
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R481
            bool isVerifiedR481 = openEmbeddedMessageResponseFirst.ReturnValue != TestSuiteBase.Success && openEmbeddedMessageResponseSecond.ReturnValue == TestSuiteBase.Success;

            // The server doesn't commit the Message object to the containing Attachment object until the RopSaveChangesMessage ROP is called if isVerifiedR481 is true, then MS-OXCMSG_R481 can be verified.
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR481,
                481,
                @"[In Receiving a RopOpenEmbeddedMessage ROP Request] The server MUST NOT commit the Message object to the containing Attachment object until the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) is called with the Embedded Message object's handle.");

            #region Call RopRelease to release the message created in step 2.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }
    }
}