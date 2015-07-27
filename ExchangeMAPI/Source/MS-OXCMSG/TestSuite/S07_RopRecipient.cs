//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCMSG
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// Test Rops related to Recipients
    /// </summary>
    [TestClass]
    public class S07_RopRecipient : TestSuiteBase
    {
        #region Const definitions for test
        /// <summary>
        /// The name of Open Specification MS-OXCDATA.
        /// </summary>
        private const string CdataShortName = "MS-OXCDATA";

        /// <summary>
        /// Constant string for the user name.
        /// </summary>
        private const string UserName = "user2";

        /// <summary>
        /// Constant string for user name of TestUser2.
        /// </summary>
        private const string TestUser2 = "TestUser2";

        /// <summary>
        /// Constant string for user name of TestUser3.
        /// </summary>
        private const string TestUser3 = "TestUser3";

        /// <summary>
        /// Constant string for user name of TestUser4.
        /// </summary>
        private const string TestUser4 = "TestUser4";

        /// <summary>
        /// Constant string for user name of TestUser5.
        /// </summary>
        private const string TestUser5 = "TestUser5";

        /// <summary>
        /// Constant string for user name of TestUser6.
        /// </summary>
        private const string TestUser6 = "TestUser6";

        /// <summary>
        /// Constant string for user name of TestUser7.
        /// </summary>
        private const string TestUser7 = "TestUser7";

        /// <summary>
        /// Constant string for user name of TestUser8.
        /// </summary>
        private const string TestUser8 = "TestUser8";

        /// <summary>
        /// Constant string for user name of TestUser9.
        /// </summary>
        private const string TestUser9 = "TestUser9";
        #endregion

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
        /// This case is used to test calling the RopModifyRecipients and RopRemoveAllRecipients operations successfully.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S07_TC01_RopModifyRemoveRecipientSuccessfully()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRemoveAllRecipients to remove all recipients of the newly created message in step 2, and expected to get a successful response.
            RopRemoveAllRecipientsRequest removeAllRecipientsRequest = new RopRemoveAllRecipientsRequest()
            {
                RopId = (byte)RopId.RopRemoveAllRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopRemoveAllRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                Reserved = 0x00000000
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(removeAllRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopRemoveAllRecipientsResponse removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1491");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1491
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                removeAllRecipientsResponse.ReturnValue,
                1491,
                @"[In Receiving a RopRemoveAllRecipients ROP Request] The call to the RopRemoveAllRecipients ROP succeeds even if the Message object on which it is executed has no recipients (2).");
            #endregion

            #region Call RopModifyRecipients to modify the recipient and expect a successful response.
            PropertyTag[] propertyTag = null;
            ModifyRecipientRow[] modifyRecipientRow = null;
            this.CreateRecipientColumnsAndRecipientRows(out propertyTag, out modifyRecipientRow);

            RopModifyRecipientsRequest modifyRecipientsRequest = new RopModifyRecipientsRequest()
            {
                RopId = (byte)RopId.RopModifyRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopModifyRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. In example, the value is 0x08
                ColumnCount = Convert.ToUInt16(propertyTag.Length),
                RowCount = Convert.ToUInt16(modifyRecipientRow.Length),
                RecipientColumns = propertyTag,
                RecipientRows = modifyRecipientRow
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(modifyRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopModifyRecipientsResponse modifyRecipientsResponse = (RopModifyRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSaveChangesMessage to save the changes of the message.
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRemoveAllRecipients to remove all recipients.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(removeAllRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, removeAllRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopGetPropertiesSpecific to get property PidTagRowid.

            // Prepare property Tag 
            PropertyTag[] tagArray = this.GetModifiedProperties();

            RopGetPropertiesSpecificRequest getPropertiesSpecificRequest = new RopGetPropertiesSpecificRequest()
            {
                RopId = (byte)RopId.RopGetPropertiesSpecific,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopGetPropertiesSpecific.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                PropertySizeLimit = 0xFFFF, // This value specifies the maximum number of the property
                PropertyTagCount = (ushort)tagArray.Length,
                PropertyTags = tagArray,
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(getPropertiesSpecificRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopGetPropertiesSpecificResponse getPropertiesSpecificResponse = (RopGetPropertiesSpecificResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, getPropertiesSpecificResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            List<PropertyObj> ps = PropertyHelper.GetPropertyObjFromBuffer(tagArray, getPropertiesSpecificResponse);

            // Parse getPropertiesSpecificResponse to get the value of properties which are modified when calling RopModifyRecipients to verify MS-OXCMSG_R381.
            PropertyObj pidTagRowid = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRowid);
            PropertyObj pidTagDisplayType = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagDisplayType);
            PropertyObj pidTagAddressBookDisplayNamePrintable = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagAddressBookDisplayNamePrintable);
            PropertyObj pidTagSmtpAddress = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagSmtpAddress);
            PropertyObj pidTagSendInternetEncoding = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagSendInternetEncoding);
            PropertyObj pidTagDisplayTypeEx = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagDisplayTypeEx);
            PropertyObj pidTagRecipientDisplayName = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientDisplayName);
            PropertyObj pidTagRecipientFlags = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientFlags);
            PropertyObj pidTagRecipientTrackStatus = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientTrackStatus);
            PropertyObj pidTagRecipientResourceState = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientResourceState);
            PropertyObj pidTagRecipientOrder = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientOrder);
            PropertyObj pidTagRecipientEntryId = PropertyHelper.GetPropertyByName(ps, PropertyNames.PidTagRecipientEntryId);

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R381. pidTagRowid = {0}, pidTagDisplayType = {1}, pidTagAddressBookDisplayNamePrintable = {2}, pidTagSmtpAddress = {3}, pidTagSendInternetEncoding = {4}, pidTagDisplayTypeEx = {5}, pidTagRecipientDisplayName = {6}, pidTagRecipientFlags = {7}, pidTagRecipientTrackStatus {8}, pidTagRecipientResourceState = {9}, pidTagRecipientOrder = {10}, pidTagRecipientEntryId = {11}",
                pidTagRowid,
                pidTagDisplayType,
                pidTagAddressBookDisplayNamePrintable,
                pidTagSmtpAddress,
                pidTagSendInternetEncoding,
                pidTagDisplayTypeEx,
                pidTagRecipientDisplayName,
                pidTagRecipientFlags,
                pidTagRecipientTrackStatus,
                pidTagRecipientResourceState,
                pidTagRecipientOrder,
                pidTagRecipientEntryId);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R381
            bool isVerifiedR381 = (pidTagRowid == null || BitConverter.ToUInt32((byte[])pidTagRowid.Value, 0) == 0x8004010F)
                && (pidTagDisplayType == null || BitConverter.ToUInt32((byte[])pidTagDisplayType.Value, 0) == 0x8004010F)
                && (pidTagAddressBookDisplayNamePrintable == null || BitConverter.ToUInt32((byte[])pidTagAddressBookDisplayNamePrintable.Value, 0) == 0x8004010F)
                && (pidTagSmtpAddress == null || BitConverter.ToUInt32((byte[])pidTagSmtpAddress.Value, 0) == 0x8004010F)
                && (pidTagSendInternetEncoding == null || BitConverter.ToUInt32((byte[])pidTagSendInternetEncoding.Value, 0) == 0x8004010F)
                && (pidTagDisplayTypeEx == null || BitConverter.ToUInt32((byte[])pidTagDisplayTypeEx.Value, 0) == 0x8004010F)
                && (pidTagRecipientDisplayName == null || BitConverter.ToUInt32((byte[])pidTagRecipientDisplayName.Value, 0) == 0x8004010F)
                && (pidTagRecipientFlags == null || BitConverter.ToUInt32((byte[])pidTagRecipientFlags.Value, 0) == 0x8004010F)
                && (pidTagRecipientTrackStatus == null || BitConverter.ToUInt32((byte[])pidTagRecipientTrackStatus.Value, 0) == 0x8004010F)
                && (pidTagRecipientResourceState == null || BitConverter.ToUInt32((byte[])pidTagRecipientResourceState.Value, 0) == 0x8004010F)
                && (pidTagRecipientOrder == null || BitConverter.ToUInt32((byte[])pidTagRecipientOrder.Value, 0) == 0x8004010F)
                && (pidTagRecipientEntryId == null || BitConverter.ToUInt32((byte[])pidTagRecipientEntryId.Value, 0) == 0x8004010F);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR381,
                381,
                @"[In Receiving a RopRemoveAllRecipients ROP Request] Until the server receives a RopSaveChangesMessage ROP request ([MS-OXCROPS] section 2.2.6.3) from the client, the server adheres to the following: The PidTagRowid property (section 2.2.1.38) and associated data of removed recipients (2) MUST NOT be returned as part of any subsequent handling of ROPs for the opened Message object on the same Message object handle.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test the RowSize structure in the response of the RopModifyRecipient.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S07_TC02_RecipientRowSize()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message.
            this.MessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopReadRecipients to read the recipient of the created message and expect the error code 0x8004010F (ecNotFound) is returned.
            RopReadRecipientsRequest readRecipientsRequest = new RopReadRecipientsRequest()
            {
                RopId = (byte)RopId.RopReadRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopReadRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                RowId = 0x00000000, // Starting index for the recipients to be retrieved
                Reserved = 0x0000 // Set the Reserved value to 0x0000 as specified in Open Specification. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, this.MessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse = (RopReadRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R403");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R403
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010f,
                readRecipientsResponse.ReturnValue,
                403,
                @"[In Receiving a RopReadRecipients ROP Request] If the message does not have recipients (2), the server returns the error ecNotFound.");
            #endregion

            #region Call RopModifyRecipient to modify the recipient.
            // Initialize TestUser1 TestUser2 TestUser5 TestUser4 TestUser3
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(TestUser1, 0, RecipientType.PrimaryRecipient),
                this.CreateModifyRecipientRow(TestUser2, 1, RecipientType.CcRecipient),
                this.CreateModifyRecipientRow(TestUser5, 2, RecipientType.BccRecipient),
                this.CreateModifyRecipientRow(TestUser4, 3),
                this.CreateModifyRecipientRow(TestUser3, 4)
            };

            RopModifyRecipientsResponse modifyRecipientsResponse;
            this.AddRecipients(modifyRecipientRow, this.MessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipient to read the recipient of the created message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, this.MessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            readRecipientsResponse = (RopReadRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R737");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R737
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                readRecipientsResponse.ReturnValue,
                737,
                @"[In RopModifyRecipients ROP] The RopModifyRecipients ROP ([MS-OXCROPS] section 2.2.6.5) modifies recipients (2) associated with the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R753");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R753
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                readRecipientsResponse.ReturnValue,
                753,
                @"[In RopReadRecipients ROP] The RopReadRecipients ROP ([MS-OXCROPS] section 2.2.6.6) retrieves the recipients (2) associated with the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R394");

            // Five recipients have been added in the above step, so MS-OXCMSG_R394 can be verified if the RowCount == 5.
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R394
            this.Site.CaptureRequirementIfAreEqual<byte>(
                5,
                readRecipientsResponse.RowCount,
                394,
                @"[In Receiving a RopModifyRecipients ROP Request] 2. Any changes made to the recipients (2) MUST be included in the response buffer for any subsequent ROP requests that apply to recipients (2) for the same Message object handle.");
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(this.MessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R679");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R679
            // 5 recipients have been added above, so MS-OXCMSG_R679 can be verified if recipient count is 5.
            this.Site.CaptureRequirementIfAreEqual<ushort>(
                5,
                openMessageResponse.RecipientCount,
                679,
                @"[In RopOpenMessage ROP Response Buffer] RecipientCount: A 2-byte unsigned integer containing the number of recipients (2) associated with the Message object.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R977");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R977
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x01,
                openMessageResponse.RecipientRows[0].RecipientType,
                977,
                @"[In RopOpenMessage ROP Response Buffer] [RecipientRows] The value 0x01 means Primary recipient.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R978");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R978
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x02,
                openMessageResponse.RecipientRows[1].RecipientType,
                978,
                @"[In RopOpenMessage ROP Response Buffer] [RecipientRows] The value 0x02 means Carbon copy (Cc) recipient.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R979");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R979
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x03,
                openMessageResponse.RecipientRows[2].RecipientType,
                979,
                @"[In RopOpenMessage ROP Response Buffer] [RecipientRows] The value 0x03 means Blind carbon copy (Bcc) recipient.");

            int rowIdFir = Convert.ToInt32(readRecipientsResponse.RecipientRows[0].RowId);
            int rowIdSec = Convert.ToInt32(readRecipientsResponse.RecipientRows[1].RowId);
            int rowIdThr = Convert.ToInt32(readRecipientsResponse.RecipientRows[2].RowId);

            bool isAscOrder = false;
            bool isDecOrder = false;
            bool isVerifyR328 = false;

            // Ascending order
            if (rowIdFir > rowIdSec)
            {
                if (rowIdSec > rowIdThr)
                {
                    isAscOrder = true;
                }
            }

            // Descending order
            if (rowIdFir < rowIdSec)
            {
                if (rowIdSec < rowIdThr)
                {
                    isDecOrder = true;
                }
            }

            if (isAscOrder || isDecOrder)
            {
                isVerifyR328 = true;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R328");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R328
            Site.CaptureRequirementIfIsTrue(
                isVerifyR328,
                328,
                @"[In Receiving a RopOpenMessage ROP Request] In addition, the server returns data for as many recipients (2) as will fit in the response buffer, in the order of the value of the RowId field.");
            #endregion

            #region Call RopReadRecipients with RowId set to an un-existing one and expect the error code 0x8004010F (ecNotFound) is returned.
            readRecipientsRequest.RowId = 0x0000000F; // Set RowId to an un-existing one.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            readRecipientsResponse = (RopReadRecipientsResponse)this.response;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1055");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1055
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                readRecipientsResponse.ReturnValue,
                1055,
                @"[In Receiving a RopReadRecipients ROP Request] [ecNotFound (0x8004010F)] Recipient row RowId does not exist on the message.");
            #endregion

            #region Call RopModifyRecipient to modify the recipient of the message.
            modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(TestUser5, 0),
                this.CreateModifyRecipientRow(TestUser4, 1),
                this.CreateModifyRecipientRow(TestUser3, 2),
                this.CreateModifyRecipientRow(TestUser1, 3),
                this.CreateModifyRecipientRow(TestUser2, 4),
                this.CreateModifyRecipientRow(TestUser6, 5),
                this.CreateModifyRecipientRow(TestUser9, 6),
                this.CreateModifyRecipientRow(TestUser7, 7),
                this.CreateModifyRecipientRow(TestUser8, 8)
            };

            this.AddRecipients(modifyRecipientRow, this.MessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients to read recipients and expect a successful response.
            readRecipientsRequest.RowId = 0x00000000; // Starting index for the recipients to be retrieved
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, this.MessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            readRecipientsResponse = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, readRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ReadRecipientRow[] recipientRowsNew = readRecipientsResponse.RecipientRows;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R389, the length of RecipientRows is {0}.", recipientRowsNew.Length);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R389
            bool isVerifiedR389 = recipientRowsNew.Length > 5;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR389,
                389,
                @"[In Receiving a RopModifyRecipients ROP Request] If the recipient (2) indicated by the value of the RowId field does not exist, the server creates a new recipient (2) with that RowId field value and applies the data from the request.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCMSG_R390, the recipients are {0}, {1}, {2}, {3}, {4}.",
                Encoding.Unicode.GetString(recipientRowsNew[0].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()),
                Encoding.Unicode.GetString(recipientRowsNew[1].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()),
                Encoding.Unicode.GetString(recipientRowsNew[2].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()),
                Encoding.Unicode.GetString(recipientRowsNew[3].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()),
                Encoding.Unicode.GetString(recipientRowsNew[4].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()));

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R390
            bool isVerifiedR390 = Encoding.Unicode.GetString(recipientRowsNew[0].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()) == TestUser5
                                && Encoding.Unicode.GetString(recipientRowsNew[1].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()) == TestUser4
                                && Encoding.Unicode.GetString(recipientRowsNew[2].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()) == TestUser3
                                && Encoding.Unicode.GetString(recipientRowsNew[3].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()) == TestUser1
                                && Encoding.Unicode.GetString(recipientRowsNew[4].RecipientRow.DisplayName).TrimEnd(TestSuiteBase.NullTerminatorString.ToCharArray()) == TestUser2;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR390,
                390,
                @"[In Receiving a RopModifyRecipients ROP Request] If the recipient (2) currently exists on the Message object and the value of RecipientRowSize field in the request buffer is nonzero, the server replaces all existing properties of the recipient (2) with the property values supplied in the request.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1495, the count of recipients is {0}.", readRecipientsResponse.RowCount);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1495
            // 9 recipients have been added in the above steps, so MS-OXCMSG_R1495 can be verified if readRecipientsResponse.RowCount == 9.
            this.Site.CaptureRequirementIfAreEqual<byte>(
                9,
                readRecipientsResponse.RowCount,
                1495,
                @"[In Receiving a RopReadRecipients ROP Request] When the value of the RowId field is 0x00000000, the server returns all recipients (2) for the message, beginning with the first recipient (2) and filling the response buffer with as many RecipientRow structures ([MS-OXCDATA] section 2.8.3) as will fit.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R400, the count of recipients is {0}.", readRecipientsResponse.RowCount);

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R400
            // 9 recipients have been added in the above steps, so MS-OXCMSG_R400 can be verified if readRecipientsResponse.RowCount == 9.
            this.Site.CaptureRequirementIfAreEqual<byte>(
                9,
                readRecipientsResponse.RowCount,
                400,
                @"[In Receiving a RopReadRecipients ROP Request] The RopReadRecipients ROP ([MS-OXCROPS] section 2.2.6.6) is used to obtain information for all recipients (2) in the Message object, regardless of the number of recipients (2) on the message.");

            // Test requirements gathered from MS-OXCDATA.
            this.TestMSOXCDATARequirements(recipientRowsNew);
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(openedMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test calling the RopModifyRecipients, RopReadRecipients and RopRemoveAllRecipients operations unsuccessfully. 
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S07_TC03_ErrorCodesOfReadModifyRemoveAllRecipients()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopModifyRecipient to modify the recipient and expect error code 0x000004B9 is returned.
            // Initialize TestUser1 
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(TestUser1, 0)
            };

            RopModifyRecipientsResponse modifyRecipientsResponse;

            RopModifyRecipientsRequest modifyRecipientsRequest = new RopModifyRecipientsRequest
            {
                RopId = (byte)RopId.RopModifyRecipients,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                ColumnCount = Convert.ToUInt16(propertyTag.Length),
                RowCount = Convert.ToUInt16(modifyRecipientRow.Count),
                RecipientColumns = propertyTag,
                RecipientRows = modifyRecipientRow.ToArray()
            };

            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(modifyRecipientsRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            modifyRecipientsResponse = (RopModifyRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1982");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1982
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                modifyRecipientsResponse.ReturnValue,
                1982,
                @"[In Receiving a RopModifyRecipients ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP was called does not refer to a Message object.");
            #endregion

            #region Call RopModifyRecipient to modify the recipient.
            // Initialize TestUser1 
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestSuiteBase.TestUser1, 0));
            this.AddRecipients(modifyRecipientRow, targetMessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients with InputHandleIndex set to 0x01 (which does not refer to a Message object) to read the recipient of the created message and expect the error code 0x000004B9 (ecNullObject) is returned.
            RopReadRecipientsRequest readRecipientsRequest = new RopReadRecipientsRequest()
            {
                RopId = (byte)RopId.RopReadRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopReadRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                RowId = 0x00000000, // Starting index for the recipients to be retrieved
                Reserved = 0x0000 // Reserved value set to 0x0000 as indicated in MS-OXCMSG. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse = (RopReadRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1057");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1057
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                readRecipientsResponse.ReturnValue,
                1057,
                @"[In Receiving a RopReadRecipients ROP Request] [ecNullObject (0x000004B9)] The InputHandleIndex on which this ROP [RopReadRecipients] was called does not refer to a Message object.");
            #endregion

            #region Call RopModifyRecipients and set ModifyRecipientRow.RecipientRowSize to 0x0000 and set RowId to an existing value.
            List<ModifyRecipientRow> modifyRecipientRowNew = new List<ModifyRecipientRow>
            {
                this.ChangeRecipientRowSize(TestUser1, 0, RecipientType.PrimaryRecipient, 0x0000)
            };

            this.AddRecipients(modifyRecipientRowNew, targetMessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients to read the recipient of the opened message.
            readRecipientsRequest.InputHandleIndex = 0x00; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, targetMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            readRecipientsResponse = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, readRecipientsResponse.ReturnValue, "Can't find any recipient of the message");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R393");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R393
            // The RowId field and associated data are not returned as part of subsequent handling of ROPs for the opened Message handle if RecipientRows is null, and then MS-OXCMSG_R393 can be verified.
            this.Site.CaptureRequirementIfIsNull(
                readRecipientsResponse.RecipientRows,
                393,
                @"[In Receiving a RopModifyRecipients ROP Request] 1. If a recipient (2) was deleted, its RowId field and associated data MUST NOT be returned as part of any subsequent handling of ROPs for the opened Message object.");
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message
            RopOpenMessageResponse openMessageResponseNew;
            openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponseNew);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponseNew.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R391. The recipient number after adding one recipient is {0}, the recipient number after adding recipient with RecipientRowSize set to 0x0000 is {1}.", openMessageResponse.RowCount, openMessageResponseNew.RowCount);

            // MS-OXCMSG_R391 can be verified if the recipient number after adding one recipient is 1 and the recipient number after adding recipient with RecipientRowSize set to 0x0000 is 0.
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R391
            bool isVerifiedR391 = openMessageResponse.RowCount == 1 && openMessageResponseNew.RowCount == 0;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR391,
                391,
                @"[In Receiving a RopModifyRecipients ROP Request] If the value of the RecipientRowSize field in the ModifyRecipientRow structure within the RecipientRows field of the request buffer is 0x0000 then the server deletes the recipient (2) from the Message object.");
            #endregion

            #region Call RopRemoveAllRecipients with InputHandleIndex set to a nonexisting one and expect error code 0x000004B9 is returned.
            RopRemoveAllRecipientsRequest removeAllRecipientsRequest = new RopRemoveAllRecipientsRequest()
            {
                RopId = (byte)RopId.RopRemoveAllRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopRemoveAllRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                Reserved = 0x00000000
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(removeAllRecipientsRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopRemoveAllRecipientsResponse removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)this.response;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1493");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1493
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004B9,
                removeAllRecipientsResponse.ReturnValue,
                1493,
                @"[In Receiving a RopRemoveAllRecipients ROP Request] [ecNullObject (0x000004B9)] The value of the InputHandleIndex field on which this ROP [RopRemoveAllRecipients] was called does not refer to a Message object.");
            #endregion

            #region Call RopModifyRecipient to modify the recipient.
            // Initialize TestUser1 TestUser2 TestUser5 TestUser4 TestUser3
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestSuiteBase.TestUser1, 0, RecipientType.PrimaryRecipient));
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestUser2, 1, RecipientType.CcRecipient));
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestUser5, 2, RecipientType.BccRecipient));
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestUser4, 3));
            modifyRecipientRow.Add(this.CreateModifyRecipientRow(TestUser3, 4));

            this.AddRecipients(modifyRecipientRow, openedMessageHandle, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRemoveAllRecipients to remove all recipients of the opened message
            removeAllRecipientsRequest.InputHandleIndex = 0x0; // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(removeAllRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, removeAllRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients to read the recipient of the opened message.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            readRecipientsResponse = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, readRecipientsResponse.ReturnValue, "Can't find any recipient of the message");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R729");
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R729
            // All recipients added above have been removed if RecipientRows is null, and then MS-OXCMSG_R729 can be verified.
            this.Site.CaptureRequirementIfIsNull(
                readRecipientsResponse.RecipientRows,
                729,
                @"[In RopRemoveAllRecipients ROP] The client sends the RopRemoveAllRecipients ROP request ([MS-OXCROPS] section 2.2.6.4) to delete all recipients (2) from a message.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// This case is used to test the transactions between RopModifyRecipients,RopReadRecipients and RopSaveChangesMessage. 
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S07_TC04_Transaction()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a message
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            RopSaveChangesMessageResponse saveChangesMessageResponse;
            saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message for the first time.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle1 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message for the second time.
            uint openedMessageHandle2 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopModifyRecipient to modify the recipient.
            // Initialize TestUser1 TestUser2 TestUser5 TestUser4 TestUser3
            PropertyTag[] propertyTag = this.CreateRecipientColumns();
            List<ModifyRecipientRow> modifyRecipientRow = new List<ModifyRecipientRow>
            {
                this.CreateModifyRecipientRow(TestUser1, 0),
                this.CreateModifyRecipientRow(TestUser2, 1),
                this.CreateModifyRecipientRow(TestUser5, 2),
                this.CreateModifyRecipientRow(TestUser4, 3),
                this.CreateModifyRecipientRow(TestUser3, 4)
            };

            RopModifyRecipientsResponse modifyRecipientsResponse;
            this.AddRecipients(modifyRecipientRow, openedMessageHandle1, propertyTag, out modifyRecipientsResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, modifyRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients to read the recipient of the created message and expect the error code 0x8004010F (ecNotFound) is returned.
            RopReadRecipientsRequest readRecipientsRequest = new RopReadRecipientsRequest()
            {
                RopId = (byte)RopId.RopReadRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopReadRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                RowId = 0x00000000, // Starting index for the recipients to be retrieved
                Reserved = 0x0000 // Set the Reserved value to 0x0000 as indicated in Open Specification. 
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle2, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse1 = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, readRecipientsResponse1.ReturnValue, "Can't find any recipients of the message.");
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle1, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message for the third time.
            uint openedMessageHandle3 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopReadRecipients to read the recipient of the created message again and expect a successful response
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle3, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse2 = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, readRecipientsResponse2.ReturnValue, TestSuiteBase.ROPSucceedMsg);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R395, readRecipientsResponse1.ReturnValue is {0}, readRecipientsResponse2.ReturnValue is {1}.", readRecipientsResponse1.ReturnValue, readRecipientsResponse2.ReturnValue);
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R395
            bool isVerifiedR395 = readRecipientsResponse1.ReturnValue != 0x00000000 && readRecipientsResponse2.ReturnValue == 0x00000000;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR395,
                395,
                @"[In Receiving a RopModifyRecipients ROP Request] 3. The changes made to the recipients (2) MUST NOT be included in the response buffer returned for ROP requests that apply to recipients (2) on different Message object handles.");
            #endregion

            #region Call RopOpenMessage to open the message for the fourth time
            uint openedMessageHandle4 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopRemoveAllRecipients to remove all recipients of the newly created message in step 2, and expected to get the successful response.
            RopRemoveAllRecipientsRequest removeAllRecipientsRequest = new RopRemoveAllRecipientsRequest()
            {
                RopId = (byte)RopId.RopRemoveAllRecipients,
                LogonId = CommonLogonId, // The logonId 0x00 is associated with RopRemoveAllRecipients.
                InputHandleIndex = CommonInputHandleIndex, // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                Reserved = 0x00000000
            };
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(removeAllRecipientsRequest, openedMessageHandle3, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopRemoveAllRecipientsResponse removeAllRecipientsResponse = (RopRemoveAllRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, removeAllRecipientsResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Use the message handle 4 to call RopReadRecipients to get all recipients of the message created in step 2.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle4, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse3 = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, readRecipientsResponse3.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopSaveChangesMessage to save the message
            saveChangesMessageResponse = this.SaveMessage(openedMessageHandle3, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Call RopOpenMessage to open the message for the fifth time
            uint openedMessageHandle5 = this.OpenSpecificMessage(logonResponse.FolderIds[4], saveChangesMessageResponse.MessageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, openMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            #endregion

            #region Use the message handle 5 to call RopReadRecipients to get all recipients of the message created in step 2.
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(readRecipientsRequest, openedMessageHandle5, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReadRecipientsResponse readRecipientsResponse4 = (RopReadRecipientsResponse)this.response;
            Site.Assert.AreNotEqual<uint>(TestSuiteBase.Success, readRecipientsResponse4.ReturnValue, "Can't find any recipients of the message.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R382. readRecipientsResponse3.ReturnValue is {0}, readRecipientsResponse3.ReturnValue is {1}.", readRecipientsResponse3.ReturnValue, readRecipientsResponse4.ReturnValue);
        
            // Verify MS-OXCMSG requirement: MS-OXCMSG_R382
            bool isVerifiedR382 = readRecipientsResponse3.ReturnValue == 0x00000000 && readRecipientsResponse4.ReturnValue != 0x00000000;
            
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR382,
                382,
                @"[In Receiving a RopRemoveAllRecipients ROP Request] [Until the server receives a RopSaveChangesMessage ROP request ([MS-OXCROPS] section 2.2.6.3) from the client, the server adheres to the following:] The changes made to the recipients (2) MUST NOT be included in the response buffer returned for ROP requests that apply to recipients (2) on different Message object handles.");
            #endregion

            #region Call RopRelease to release the created message.
            this.ReleaseRop(targetMessageHandle);
            #endregion
        }

        /// <summary>
        /// Create Recipient Array 
        /// </summary>
        /// <param name="name">Recipient name value</param>
        /// <param name="rowId">RowId value</param>
        /// <param name="recipientType">RecipientType value</param>
        /// <param name="recipientRowSize">RecipientRowSize value</param>
        /// <returns>Return ModifyRecipientRow</returns>
        protected ModifyRecipientRow ChangeRecipientRowSize(string name, uint rowId, RecipientType recipientType, ushort recipientRowSize)
        {
            PropertyRow propertyRow = this.CreateRecipientColumns(name);

            RecipientRow recipientRow = new RecipientRow
            {
                RecipientFlags = 0x065B,
                DisplayName = Common.GetBytesFromUnicodeString(name)
            };
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            recipientRow.EmailAddress = Common.GetBytesFromUnicodeString(string.Format("{0}{1}{2}", name, TestSuiteBase.At, domainName));
            recipientRow.SimpleDisplayName = Common.GetBytesFromUnicodeString(TestSuiteBase.PrefixOfDisplayName + name);
            recipientRow.RecipientColumnCount = 0x000C; // Matches ColummnCount
            recipientRow.RecipientProperties = propertyRow;

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = rowId,
                RecipientType = (byte)recipientType,
                RecipientRowSize = recipientRowSize,
                RecptRow = recipientRow.Serialize()
            };

            return modifyRecipientRow;
        }

        /// <summary>
        /// This method is used to verify requirements gathered from MS-OXCDATA.
        /// </summary>
        /// <param name="recipientRows">ReadRecipientRow structure.</param>
        private void TestMSOXCDATARequirements(ReadRecipientRow[] recipientRows)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R177. The DisplayName is {0}.", recipientRows[0].RecipientRow.DisplayName);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R177
            Site.CaptureRequirementIfAreEqual<int>(
                recipientRows[0].RecipientRow.RecipientColumnCount,
                recipientRows[0].RecipientRow.RecipientProperties.PropertyValues.Count,
                CdataShortName,
                177,
                @"[In RecipientRow] RecipientColumnCount (2 bytes): This value [RecipientColumnCount] specifies the number of columns from the RecipientColumns field that are included in the RecipientProperties field.");

            bool isFieldRExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0080);

            if (isFieldRExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R105. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R105
                // R flag exists if RecipientFlags & 0x0080 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    105,
                    @"[In RecipientFlags Field] R (1 bit): (mask 0x0080).");
            }

            bool isFieldSExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0040);

            if (isFieldSExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R107. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R107
                // S flag exists if RecipientFlags & 0x0040 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    107,
                    @"[In RecipientFlags Field] S (1 bit): (mask 0x0040).");
            }

            bool isFieldDExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0010);

            if (isFieldDExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R111. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R111
                // D flag exists if RecipientFlags & 0x0010 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    111,
                    @"[In RecipientFlags Field] D (1 bit): (mask 0x0010).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R112. The DisplayName is {0}.", recipientRows[0].RecipientRow.DisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R112
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.DisplayName,
                    CdataShortName,
                    112,
                    @"[In RecipientFlags Field] D (1 bit): If this flag is b'1', the DisplayName field is included.");
                
                int lengthOfDisplayName = recipientRows[0].RecipientRow.DisplayName.Length;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R164. The DisplayName is {0}.", recipientRows[0].RecipientRow.DisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R164
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    recipientRows[0].RecipientRow.DisplayName[lengthOfDisplayName - 1],
                    CdataShortName,
                    164,
                    @"[In RecipientRow Structure] DisplayName (optional) (variable): A null-terminated string.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R165. The DisplayName is {0}.", recipientRows[0].RecipientRow.DisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R165
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.DisplayName,
                    CdataShortName,
                    165,
                    @"[In RecipientRow Structure] DisplayName (optional) (variable): This field [DisplayName (optional) (variable)] MUST be present when the D flag of the RecipientsFlags field is set.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R166. The DisplayName is {0}.", recipientRows[0].RecipientRow.DisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R166
                bool isVerifiedR166 = Common.IsUtf16LEString(recipientRows[0].RecipientRow.DisplayName);
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR166,
                    CdataShortName,
                    166,
                    @"[In RecipientRow Structure] DisplayName (optional) (variable): This field [DisplayName] MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and in the 8-bit character.");
            }

            bool isFieldEExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0008);

            if (isFieldEExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R113. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R113
                // E flag exists if RecipientFlags & 0x0008 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    113,
                    @"[In RecipientFlags Field] E (1 bit): (mask 0x0008).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R114. The EmailAddress is {0}.", recipientRows[0].RecipientRow.EmailAddress);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R112
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.EmailAddress,
                    CdataShortName,
                    114,
                    @"[In RecipientFlags Field] E (1 bit): If this flag is b'1', the EmailAddress field is included.");

                 int lengthOfEmailAddress = recipientRows[0].RecipientRow.EmailAddress.Length;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R160. The EmailAddress is {0}.", recipientRows[0].RecipientRow.EmailAddress);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R160
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    recipientRows[0].RecipientRow.EmailAddress[lengthOfEmailAddress - 1],
                    CdataShortName,
                    160,
                    @"[In RecipientRow Structure] EmailAddress (optional) (variable): A null-terminated string.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R161. The EmailAddress is {0}.", recipientRows[0].RecipientRow.EmailAddress);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R161
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.EmailAddress,
                    CdataShortName,
                    161,
                    @"[In RecipientRow Structure] EmailAddress (optional) (variable): This field [EmailAddress (optional) (variable)] MUST be present when the E flag of the RecipientsFlags field is set.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R162. The EmailAddress is {0}.", recipientRows[0].RecipientRow.EmailAddress);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R162
                bool isVerifiedR162 = Common.IsUtf16LEString(recipientRows[0].RecipientRow.EmailAddress);
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR162,
                    CdataShortName,
                    162,
                    @"[In RecipientRow Structure] EmailAddress (optional) (variable): This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and in the 8-bit character set.");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R115. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R115
            bool isFieldTypeExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0007);

            Site.CaptureRequirementIfIsTrue(
                isFieldTypeExist,
                CdataShortName,
                115,
                @"[In RecipientFlags Field] Type (3 bits): (mask 0x0007).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R117. The value of Type is {0}.", (int)recipientRows[0].RecipientRow.RecipientFlags & 0x0007);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R117
            int valueOfTypeField = (int)recipientRows[0].RecipientRow.RecipientFlags & 0x0007;
            bool isTypeValid = (valueOfTypeField & 0x0) == 0x0
                || (valueOfTypeField & 0x1) == 0x1
                || (valueOfTypeField & 0x2) == 0x2
                || (valueOfTypeField & 0x3) == 0x3
                || (valueOfTypeField & 0x4) == 0x4
                || (valueOfTypeField & 0x5) == 0x5
                || (valueOfTypeField & 0x6) == 0x6
                || (valueOfTypeField & 0x7) == 0x7;

            Site.CaptureRequirementIfIsTrue(
                isTypeValid,
                CdataShortName,
                117,
                @"[In  RecipientFlags Field] Type (3 bits):The valid types are: [NoType (0x0), X500DN (0x1), MsMail (0x2), SMTP (0x3), Fax (0x4), ProfessionalOfficeSystem (0x5), PersonalDistributionList1 (0x6), PersonalDistributionList2 (0x7)].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R120.");

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R108
            Site.CaptureRequirementIfAreEqual<int>(
                0x0000,
                (int)recipientRows[0].RecipientRow.RecipientFlags & 0x7800,
                CdataShortName,
                120,
                @"[In  RecipientFlags Field] Reserved (4 bits):  (mask 0x7800) The server MUST set this to b'0000'.");

            bool isFieldIExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0400);

            if (isFieldIExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R121. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R121
                // I flag exists if RecipientFlags & 0x0400 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    121,
                    @"[In  RecipientFlags Field] I (1 bit): (mask 0x0400).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R122. The SimpleDisplayName is {0}.", recipientRows[0].RecipientRow.SimpleDisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R122
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.SimpleDisplayName,
                    CdataShortName,
                    122,
                    @"[In RecipientFlags Field] I (1 bit): If this flag is b'1', the SimpleDisplayName field is included.");

                  int lengthOfSimpleDisplayName = recipientRows[0].RecipientRow.SimpleDisplayName.Length;

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R168. The SimpleDisplayName is {0}.", recipientRows[0].RecipientRow.SimpleDisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R168
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    recipientRows[0].RecipientRow.SimpleDisplayName[lengthOfSimpleDisplayName - 1],
                    CdataShortName,
                    168,
                    @"[In RecipientRow Structure] SimpleDisplayName (optional) (variable): A null-terminated string.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R169. The SimpleDisplayName is {0}.", recipientRows[0].RecipientRow.SimpleDisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R169
                Site.CaptureRequirementIfIsNotNull(
                    recipientRows[0].RecipientRow.SimpleDisplayName,
                    CdataShortName,
                    169,
                    @"[In RecipientRow Structure] SimpleDisplayName (optional) (variable): This field [SimpleDisplayName (optional) (variable)] MUST be present when the I flag of the RecipientsFlags field is set.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R170. The SimpleDisplayName is {0}.", recipientRows[0].RecipientRow.SimpleDisplayName);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R170
                bool isVerifiedR170 = Common.IsUtf16LEString(recipientRows[0].RecipientRow.SimpleDisplayName);
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR170,
                    CdataShortName,
                    170,
                    @"[In RecipientRow Structure] SimpleDisplayName (optional) (variable): This field MUST be specified in Unicode characters if the U flag of the RecipientsFlags field is set and in the 8-bit character set.");
            }

            bool isFieldUExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x0200);

            if (isFieldUExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R123. The RecipientFlags is {0}.", recipientRows[0].RecipientRow.RecipientFlags);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R123
                // U flag exists if RecipientFlags & 0x0200 is not 0.
                Site.CaptureRequirement(
                    CdataShortName,
                    123,
                    @"[In  RecipientFlags Field] U (1 bit): (mask 0x0200).");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2764. The DisplayType is {0}.", recipientRows[0].RecipientRow.DisplayType);

            // Verify MS-OXCDATA requirement: MS-OXCDATA_R2764
            bool isDisplayTypeValid = (recipientRows[0].RecipientRow.DisplayType & 0x00) == 0x00
                || (recipientRows[0].RecipientRow.DisplayType & 0x01) == 0x01
                || (recipientRows[0].RecipientRow.DisplayType & 0x02) == 0x02
                || (recipientRows[0].RecipientRow.DisplayType & 0x03) == 0x03
                || (recipientRows[0].RecipientRow.DisplayType & 0x04) == 0x04
                || (recipientRows[0].RecipientRow.DisplayType & 0x05) == 0x05
                || (recipientRows[0].RecipientRow.DisplayType & 0x06) == 0x06;

            Site.CaptureRequirementIfIsTrue(
                isDisplayTypeValid,
                CdataShortName,
                2764,
                @"[In RecipientRow Structure] Valid values [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06] for this field [DisplayType (optional) (1 byte)] are specified in the following table.");

            if (((int)(recipientRows[0].RecipientRow.RecipientFlags & 0x0007) != 0x6) && ((int)(recipientRows[0].RecipientRow.RecipientFlags & 0x0007) != 0x7))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R142. The EntryIdSize is {0}.", recipientRows[0].RecipientRow.EntryIdSize);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R142
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    recipientRows[0].RecipientRow.EntryIdSize,
                    CdataShortName,
                    142,
                    @"[In RecipientRow Structure] EntryIdSize (optional) (2 bytes): This field MUST NOT be present otherwise. This value specifies the size of the EntryID field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R145. The EntryId is {0}.", recipientRows[0].RecipientRow.EntryId);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R145
                Site.CaptureRequirementIfIsNull(
                    recipientRows[0].RecipientRow.EntryId,
                    CdataShortName,
                    145,
                    @"[In RecipientRow Structure] EntryID (optional) (variable): This field [EntryID (optional) (variable)] MUST NOT be present otherwise.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R150. The SearchKeySize is {0}.", recipientRows[0].RecipientRow.SearchKeySize);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R150
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0,
                    recipientRows[0].RecipientRow.SearchKeySize,
                    CdataShortName,
                    150,
                    @"[In RecipientRow Structure] SearchKeySize (optional) (2 bytes): This field MUST NOT be present otherwise. This value specifies the size of the SearchKey field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R153. The SearchKey is {0}.", recipientRows[0].RecipientRow.SearchKey);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R150
                Site.CaptureRequirementIfIsNull(
                    recipientRows[0].RecipientRow.SearchKey,
                    CdataShortName,
                    153,
                    @"[In RecipientRow Structure] SearchKey (optional) (variable): This field [SearchKey (optional) (variable)] MUST NOT be present otherwise.");
            }

            bool isFieldOExist = Convert.ToBoolean(recipientRows[0].RecipientRow.RecipientFlags & 0x8000);

            if (!((int)(recipientRows[0].RecipientRow.RecipientFlags & 0x0007) == 0x0) || !isFieldOExist)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R158. The AddressType is {0}.", recipientRows[0].RecipientRow.AddressType);

                // Verify MS-OXCDATA requirement: MS-OXCDATA_R158
                Site.CaptureRequirementIfIsNull(
                    recipientRows[0].RecipientRow.AddressType,
                    CdataShortName,
                    158,
                    @"[In RecipientRow Structure] AddressType (optional) (variable):  This field [AddressType (optional) (variable)] MUST NOT be present otherwise.");
            }
        }

        /// <summary>
        /// Create properties array of get all properties after modifying all recipients.
        /// </summary>
        /// <returns>List of PropertyTag</returns>
        private PropertyTag[] GetModifiedProperties()
        {
            PropertyTag[] tags = new PropertyTag[12];
            PropertyTag tag;

            // PidTagDisplayType
            tag.PropertyId = 0x3900;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[0] = tag;

            // PidTagAddressBookDisplayNamePrintable
            tag.PropertyId = 0x39ff;
            tag.PropertyType = 0x001f; // PtypString
            tags[1] = tag;

            // PidTagSmtpAddress
            tag.PropertyId = 0x39fe;
            tag.PropertyType = 0x001f; // PtypString
            tags[2] = tag;

            // PidTagSendInternetEncoding
            tag.PropertyId = 0x3a71;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[3] = tag;

            // PidTagDisplayTypeEx
            tag.PropertyId = 0x3905;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[4] = tag;

            // PidTagRecipientDisplayName
            tag.PropertyId = 0x5ff6;
            tag.PropertyType = 0x001f; // PtypString
            tags[5] = tag;

            // PidTagRecipientFlags
            tag.PropertyId = 0x5ffd;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[6] = tag;

            // PidTagRecipientTrackStatus
            tag.PropertyId = 0x5fff;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[7] = tag;

            // PidTagRecipientResourceState
            tag.PropertyId = 0x5fde;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[8] = tag;

            // PidTagRecipientOrder
            tag.PropertyId = 0x5fdf;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[9] = tag;

            // PidTagRecipientEntryId
            tag.PropertyId = 0x5ff7;
            tag.PropertyType = 0x0102; // PtypBinary
            tags[10] = tag;

            // PidTagRowid 
            tag.PropertyId = 0x3000;
            tag.PropertyType = 0x0003; // PtypInteger32
            tags[11] = tag;

            return tags;
        }

        /// <summary>
        /// Create Recipient Array 
        /// </summary>
        /// <param name="recipientColumns">Recipient Columns </param>
        /// <param name="recipientRows">Recipient Rows</param>
        private void CreateRecipientColumnsAndRecipientRows(out PropertyTag[] recipientColumns, out ModifyRecipientRow[] recipientRows)
        {
            #region recipientColumns

            // The following sample data is from MS-OXCMSG
            PropertyTag[] sampleRecipientColumns = new PropertyTag[12];
            PropertyTag tag;

            // PidTagObjectType
            tag.PropertyId = 0x0ffe;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[0] = tag;

            // PidTagDisplayType
            tag.PropertyId = 0x3900;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[1] = tag;

            // PidTagAddressBookDisplayNamePrintable
            tag.PropertyId = 0x39ff;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[2] = tag;

            // PidTagSmtpAddress
            tag.PropertyId = 0x39fe;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[3] = tag;

            // PidTagSendInternetEncoding
            tag.PropertyId = 0x3a71;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[4] = tag;

            // PidTagDisplayTypeEx
            tag.PropertyId = 0x3905;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[5] = tag;

            // PidTagRecipientDisplayName
            tag.PropertyId = 0x5ff6;
            tag.PropertyType = 0x001f; // PtypString
            sampleRecipientColumns[6] = tag;

            // PidTagRecipientFlags
            tag.PropertyId = 0x5ffd;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[7] = tag;

            // PidTagRecipientTrackStatus
            tag.PropertyId = 0x5fff;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[8] = tag;

            // PidTagRecipientResourceState
            tag.PropertyId = 0x5fde;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[9] = tag;

            // PidTagRecipientOrder
            tag.PropertyId = 0x5fdf;
            tag.PropertyType = 0x0003; // PtypInteger32
            sampleRecipientColumns[10] = tag;

            // PidTagRecipientEntryId
            tag.PropertyId = 0x5ff7;
            tag.PropertyType = 0x0102; // PtypBinary
            sampleRecipientColumns[11] = tag;

            recipientColumns = sampleRecipientColumns;
            #endregion

            #region Configure a StandardPropertyRow: propertyRow, data is from Page 62 of MS-OXCMSG

            PropertyValue[] propertyValueArray = new PropertyValue[12];
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValueArray[i] = new PropertyValue();
            }

            propertyValueArray[0].Value = BitConverter.GetBytes(0x00000006); // PidTagObjectType
            propertyValueArray[1].Value = BitConverter.GetBytes(0x00000000); // PidTagDisplayType
            propertyValueArray[2].Value = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.NullTerminatorString); // PidTa7BitDisplayName
            string domainName = Common.GetConfigurationPropertyValue("Domain", this.Site);
            propertyValueArray[3].Value = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.At + domainName + TestSuiteBase.NullTerminatorString); // PidTagSmtpAddress
            propertyValueArray[4].Value = BitConverter.GetBytes(0x00000000); // PidTagSendInternetEncoding
            propertyValueArray[5].Value = BitConverter.GetBytes(0x40000000); // PidTagDisplayTypeEx
            propertyValueArray[6].Value = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.NullTerminatorString); // PidTagRecipientDisplayName
            propertyValueArray[7].Value = BitConverter.GetBytes(0x00000001); // PidTagRecipientFlags
            propertyValueArray[8].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientTrackStatus
            propertyValueArray[9].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientResourceState
            propertyValueArray[10].Value = BitConverter.GetBytes(0x00000000); // PidTagRecipientOrder

            // The following sample data (0x007c and the subsequent 124(0x7c) binary) 
            byte[] sampleData = 
            {
                0x7c, 0x00, 0x00, 0x00, 0x00, 0x00, 0xdc, 0xa7, 0x40,
                0xc8, 0xc0, 0x42, 0x10, 0x1a, 0xb4, 0xb9, 0x08, 0x00, 0x2b, 0x2f, 0xe1, 0x82, 0x01, 0x00, 0x00,
                0x00, 0x00, 0x00, 0x00, 0x00, 0x2f, 0x6f, 0x3d, 0x46, 0x69, 0x72, 0x73, 0x74, 0x20, 0x4f, 0x72,
                0x67, 0x61, 0x6e, 0x69, 0x7a, 0x61, 0x74, 0x69, 0x6f, 0x6e, 0x2f, 0x6f, 0x75, 0x3d, 0x45, 0x78,
                0x63, 0x68, 0x61, 0x6e, 0x67, 0x65, 0x20, 0x41, 0x64, 0x6d, 0x69, 0x6e, 0x69, 0x73, 0x74, 0x72,
                0x61, 0x74, 0x69, 0x76, 0x65, 0x20, 0x47, 0x72, 0x6f, 0x75, 0x70, 0x20, 0x28, 0x46, 0x59, 0x44,
                0x49, 0x42, 0x4f, 0x48, 0x46, 0x32, 0x33, 0x53, 0x50, 0x44, 0x4c, 0x54, 0x29, 0x2f, 0x63, 0x6e,
                0x3d, 0x52, 0x65, 0x63, 0x69, 0x70, 0x69, 0x65, 0x6e, 0x74, 0x73, 0x2f, 0x63, 0x6e, 0x3d, 0x75,
                0x73, 0x65, 0x72, 0x32, 0x00
            };
            propertyValueArray[11].Value = sampleData; // PidTagRecipientEntryId

            List<PropertyValue> propertyValues = new List<PropertyValue>();
            for (int i = 0; i < propertyValueArray.Length; i++)
            {
                propertyValues.Add(propertyValueArray[i]);
            }

            PropertyRow propertyRow = new PropertyRow
            {
                Flag = 0x01, PropertyValues = propertyValues
            };
            #endregion

            RecipientRow recipientRow = new RecipientRow
            {
                RecipientFlags = 0x065B,
                DisplayName = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.NullTerminatorString),
                EmailAddress = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.At + domainName + TestSuiteBase.NullTerminatorString),
                SimpleDisplayName = Encoding.Unicode.GetBytes(UserName + TestSuiteBase.NullTerminatorString),
                RecipientColumnCount = 0x000C,
                RecipientProperties = propertyRow
            };

            ModifyRecipientRow modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = 0x00000000,
                RecipientType = (byte)RecipientType.PrimaryRecipient,
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };

            ModifyRecipientRow[] sampleModifyRecipientRows = new ModifyRecipientRow[3];
            sampleModifyRecipientRows[0] = modifyRecipientRow;

            // Test add second recipient 
            ModifyRecipientRow testmodifyRecipientRow = new ModifyRecipientRow
            {
                RowId = 0x0000001,
                RecipientType = (byte)RecipientType.PrimaryRecipient,
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };
            sampleModifyRecipientRows[1] = testmodifyRecipientRow;

            ModifyRecipientRow test2modifyRecipientRow = new ModifyRecipientRow
            {
                RowId = 0x0000002,
                RecipientType = (byte)RecipientType.PrimaryRecipient,
                RecipientRowSize = (ushort)recipientRow.Size(),
                RecptRow = recipientRow.Serialize()
            };
            sampleModifyRecipientRows[2] = test2modifyRecipientRow;

            recipientRows = sampleModifyRecipientRows;
        }
    }
}