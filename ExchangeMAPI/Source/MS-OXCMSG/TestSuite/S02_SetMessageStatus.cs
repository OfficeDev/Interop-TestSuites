namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that Verifies the requirements related to RopSetMessageStatus and RopGetMessageStatus operations.
    /// </summary>
    [TestClass]
    public class S02_SetMessageStatus : TestSuiteBase
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
        /// This test case is used to validate the RopSetMessageStatus and RopGetMessageStatus ROP operations.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S02_TC01_SetAndGetMessageStatus()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            List<PropertyTag> propertyTags = new List<PropertyTag>
            {
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagMessageStatus],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagChangeKey],
                PropertyHelper.PropertyTagDic[PropertyNames.PidTagLastModificationTime]
            };

            #region Call RopLogon to logon the private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a new message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to commit the new message object.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRelease to release all resources.
            this.ReleaseRop(targetMessageHandle);
            #endregion

            #region Call RopOpenFolder to open the inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagMessageStatus,PidTagChangeKey and PidTagLastModificationTime property before set MessagStatus.
            List<PropertyObj> ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageStatusBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageStatus);
            PropertyObj pidTagChangeKeyBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeBeforeSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);
            #endregion

            #region Call RopSetMessageStatus to set the MessageStatusFlags property of specific message to 0x00001000.
            RopSetMessageStatusRequest setMessageStatusRequest = new RopSetMessageStatusRequest()
            {
                RopId = (byte)RopId.RopSetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId,
                MessageStatusFlags = (uint)MessageStatusFlags.MsRemoteDownload,
                MessageStatusMask = (uint)MessageStatusFlags.MsInConflict | (uint)MessageStatusFlags.MsRemoteDownload | (uint)MessageStatusFlags.MsRemoteDelete
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetMessageStatusResponse setMessageStatusResponse = (RopSetMessageStatusResponse)this.response;

            #region Verify MS-OXCMSG_R773 and MS-OXCMSG_R549
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R773");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R773
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                setMessageStatusResponse.ReturnValue,
                773,
                @"[In RopSetMessageStatus ROP] The RopSetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.8) sets the PidTagMessageStatus property ([MS-OXPROPS] section 2.792) on a message in a folder without the need to open or save the Message object.");

            // Because the MS-OXCMSG_R773 has been captured and verify that the RopSetMessthatageStatus ROP sets the PidTagMessageStatus property.
            // So MS-OXCMSG_R549 can be captured directly.
            this.Site.CaptureRequirement(
                549,
                @"[In Sending a RopSetProperties ROP Request] Instead, the client calls the RopSetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.8), as specified in section 2.2.3.8.");
            #endregion
            #endregion

            #region Call RopGetMessageStatus to get the message status of specific message created step 2.
            RopGetMessageStatusRequest getMessageStatusRequest = new RopGetMessageStatusRequest()
            {
                RopId = (byte)RopId.RopGetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            // The response buffers for RopGetMessageStatus are the same as those for RopSetMessageStatus.
            RopSetMessageStatusResponse getMessageStatusRespnse = (RopSetMessageStatusResponse)this.response;

            #region MS-OXCMSG_R786, MS-OXCMSG_R2028, MS-OXCMSG_R421
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R786, the MessageStatusFlags is {0}.", getMessageStatusRespnse.MessageStatusFlags);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R786
            bool isVerifiedR786 = getMessageStatusRespnse.ReturnValue == TestSuiteBase.Success && getMessageStatusRespnse.MessageStatusFlags == 0x00001000;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR786,
                786,
                @"[In RopGetMessageStatus ROP] The RopGetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.9) gets the message status of a message in a folder.");

            // Because the MS-OXCMSG_R786 has been captured and verify that the RopGetMessageStatus gets the message status of a message.
            // So OXCMSG_R55 can be captured directly.
            this.Site.CaptureRequirement(
                55,
                @"[In Sending a RopGetPropertiesSpecific ROP Request] Instead, the client calls the RopGetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.9), as specified in section 2.2.3.9.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2028");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2028
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00001000,
                getMessageStatusRespnse.MessageStatusFlags,
                2028,
                @"[In PidTagMessageStatus Property] [The value of flag msRemoteDownload is] 0x00001000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R421");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R421
            // Because the message handle has been release in above step.
            // If RopGetMessageStatus execute successfully then R421 will be verified.
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                getMessageStatusRespnse.ReturnValue,
                421,
                @"[In Receiving a RopGetMessageStatus ROP Request] When processing the RopGetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.9), the server MUST NOT require the Message object to be opened.");
            #endregion
            #endregion

            #region Call RopGetPropertiesSpecific to get the PidTagMessageStatus,PidTagChangeKey and PidTagLastModificationTime property after set MessagStatus.
            ps = this.GetSpecificPropertiesOfMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, propertyTags);
            PropertyObj pidTagMessageStatusAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagMessageStatus);
            PropertyObj pidTagChangeKeyAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagChangeKey);
            PropertyObj pidTagLastModificationTimeAfterSet = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagLastModificationTime);

            #region Verify MS-OXCMSG_R1501
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1501");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1501
            bool isVerifiedR1501 =
                Common.CompareByteArray((byte[])pidTagChangeKeyBeforeSet.Value, (byte[])pidTagChangeKeyAfterSet.Value) == true &&
                pidTagLastModificationTimeBeforeSet.Value.Equals(pidTagLastModificationTimeAfterSet.Value) &&
                pidTagMessageStatusBeforeSet.Value != pidTagMessageStatusAfterSet.Value;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1501,
                1501,
                @"[In Receiving a RopSetMessageStatus ROP Request] The server immediately commits the changes to the Message object as if the Message object had been opened and the RopSaveChangesMessage ROP ([MS-OXCROPS] section 2.2.6.3) had been called, except that it [server] changes only the PidTagMessageStatus property, not the PidTagChangeKey property ([MS-OXCFXICS] section 2.2.1.2.7), the PidTagLastModificationTime property (section 2.2.2.2), or any other property that is modified during the RopSaveChangesMessage ROP request.");
            #endregion
            #endregion

            #region Call RopSetMessageStatus to set the MessageStatusFlags property of specific message to 0x00000800.
            setMessageStatusRequest = new RopSetMessageStatusRequest()
            {
                RopId = (byte)RopId.RopSetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId,
                MessageStatusFlags = (uint)MessageStatusFlags.MsInConflict,
                MessageStatusMask = (uint)MessageStatusFlags.MsInConflict | (uint)MessageStatusFlags.MsRemoteDownload | (uint)MessageStatusFlags.MsRemoteDelete
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageStatusResponse = (RopSetMessageStatusResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageStatusResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #region Verify MS-OXCMSG_R784
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R784");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R784
            this.Site.CaptureRequirementIfAreEqual<uint>(
                getMessageStatusRespnse.MessageStatusFlags,
                setMessageStatusResponse.MessageStatusFlags,
                784,
                @"[In RopSetMessageStatus ROP Response Buffer] MessageStatusFlags: 4 bytes indicating the status flags that were set on the Message object before processing this request.");
            #endregion
            #endregion

            #region Call RopGetMessageStatus to get the message status of specific message created step 2.
            // RopGetMessageStatusResponse getMessageStatusResponse;
            getMessageStatusRequest.MessageId = saveChangesMessageResponse.MessageId;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getMessageStatusRespnse = (RopSetMessageStatusResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageStatusResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #region Verify requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2029");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2029
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000800,
                getMessageStatusRespnse.MessageStatusFlags,
                2029,
                @"[In PidTagMessageStatus Property] [The value of flag msInConflict is] 0x00000800.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R416, the MessageStatusFlags value is {0}.", setMessageStatusResponse.MessageStatusFlags);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R416
            bool isVerifiedR416 = getMessageStatusRespnse.MessageStatusFlags == (setMessageStatusRequest.MessageStatusMask & setMessageStatusRequest.MessageStatusFlags);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR416,
                416,
                @"[In Receiving a RopSetMessageStatus ROP Request] When processing the RopSetMessageStatus ROP ([MS-OXCROPS] section 2.2.6.8), the server modifies the bits on the PidTagMessageStatus property (section 2.2.1.8) specified by the MessageStatusMask field, preserving only those flags that are set in both the MessageStatusMask field and the MessageStatusFlags field, and clearing any other flags set only in the MessageStatusMask field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1500");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1500
            bool isVerifiedR1500 = (getMessageStatusRespnse.MessageStatusFlags & (~setMessageStatusRequest.MessageStatusFlags & setMessageStatusRequest.MessageStatusMask)) == 0x00000000;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR1500,
                1500,
                @"[In Receiving a RopSetMessageStatus ROP Request] [When processing the RopSetMessageStatus ROP, the server] clearing any other flags set only in the MessageStatusMask field.");
            #endregion
            #endregion

            #region Call RopSetMessageStatus to set the MessageStatusFlags property of specific message to 0x00002000.
            setMessageStatusRequest = new RopSetMessageStatusRequest()
            {
                RopId = (byte)RopId.RopSetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId,
                MessageStatusFlags = (uint)MessageStatusFlags.MsRemoteDelete,
                MessageStatusMask = (uint)MessageStatusFlags.MsInConflict | (uint)MessageStatusFlags.MsRemoteDownload | (uint)MessageStatusFlags.MsRemoteDelete
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            setMessageStatusResponse = (RopSetMessageStatusResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageStatusResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetMessageStatus get the message status of specific message created step 2.
            getMessageStatusRequest.MessageId = saveChangesMessageResponse.MessageId;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getMessageStatusRequest, folderHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            getMessageStatusRespnse = (RopSetMessageStatusResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, setMessageStatusResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            #region MS-MS-OXCMSG_R2030
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R2030");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R2030
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00002000,
                getMessageStatusRespnse.MessageStatusFlags,
                2030,
                @"[In PidTagMessageStatus Property] [The value of flag msRemoteDelete is] 0x00002000.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case tests the error code in RopSetMessageStatus and RopGetMessageStatus operations.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S02_TC02_ErrorCodeOfMessageStatus()
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
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenFolder to open the inbox folder.
            uint folderHandle = this.OpenSpecificFolder(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSetMessageStatus to set message status. The value of the InputHandleIndex field on which this ROP was called does not refer to a Folder object.
            RopSetMessageStatusRequest setMessageStatusRequest = new RopSetMessageStatusRequest()
            {
                RopId = (byte)RopId.RopSetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId,
                MessageStatusFlags = (uint)MessageStatusFlags.MsRemoteDownload,
                MessageStatusMask = (uint)MessageStatusFlags.MsRemoteDownload
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(setMessageStatusRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopSetMessageStatusResponse setMessageStatusResponse = (RopSetMessageStatusResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1505");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1505
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                setMessageStatusResponse.ReturnValue,
                1505,
                @"[In Receiving a RopSetMessageStatus ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopSetMessageStatus] was called does not refer to a Folder object.");
            #endregion

            #region Call RopGetMessageStatus to get message status.The value of the InputHandleIndex field on which this ROP was called does not refer to a Folder object.
            RopGetMessageStatusRequest getMessageStatusRequest = new RopGetMessageStatusRequest
            {
                RopId = (byte)RopId.RopGetMessageStatus,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                MessageId = saveChangesMessageResponse.MessageId
            };

            // RopGetMessageStatusResponse getMessageStatusResponse;
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getMessageStatusRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);

            // The response buffers for RopGetMessageStatus are the same as those for RopSetMessageStatus
            setMessageStatusResponse = (RopSetMessageStatusResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1507");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1507
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                setMessageStatusResponse.ReturnValue,
                1507,
                @"[In Receiving a RopGetMessageStatus ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopGetMessageStatus] was called does not refer to a Folder object.");
            #endregion

            #region Call RopRelease to release created message.
            this.ReleaseRop(targetMessageHandle);
            this.ReleaseRop(folderHandle);
            #endregion
        }
    }
}