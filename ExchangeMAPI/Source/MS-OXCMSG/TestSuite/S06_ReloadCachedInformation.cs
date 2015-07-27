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
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario contains test cases that Verifies the requirements related to RopReloadCachedInformation.
    /// </summary>
    [TestClass]
    public class S06_ReloadCachedInformation : TestSuiteBase
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
        /// This test case validates the operation of RopReloadCachedInformation used to get information of the created message.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S06_TC01_RopReloadCachedInformation()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            uint logonHandle;
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out logonHandle);
            #endregion

            #region Call RopCreateMessage to create a new message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], logonHandle);
            #endregion

            #region Call RopSaveChangesMessage to commit the new message object.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            this.ReleaseRop(targetMessageHandle);

            #region Call RopOpenMessage to open the created message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, logonHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            #endregion

            #region Call RopReloadCachedInformation to get information of the specific message.
            RopReloadCachedInformationRequest reloadCachedInformationRequest = new RopReloadCachedInformationRequest()
            {
                RopId = (byte)RopId.RopReloadCachedInformation,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                Reserved = 0x0000
            };
            this.response = new RopReloadCachedInformationResponse();
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(reloadCachedInformationRequest, openedMessageHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReloadCachedInformationResponse reloadCachedInformationResponse = (RopReloadCachedInformationResponse)this.response;

            #region Verify MS-OXCMSG_R279, MS-OXCMSG_R766
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R279");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R279
            this.Site.CaptureRequirementIfAreEqual<uint>(
                TestSuiteBase.Success,
                reloadCachedInformationResponse.ReturnValue,
                279,
                @"[In Reload Message Object Header Info] A client retrieves the current state of the data returned in a RopOpenMessage ROP ([MS-OXCROPS] section 2.2.6.1) by sending a RopReloadCachedInformation ROP request ([MS-OXCROPS] section 2.2.6.7).");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R766");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R766
            bool isVerifiedR766 = this.CompareResponseOfOpenMessageAndRopReloadCachedInformation(openMessageResponse, reloadCachedInformationResponse);

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR766,
                766,
                @"[In RopReloadCachedInformation ROP] The RopReloadCachedInformation ROP ([MS-OXCROPS] section 2.2.6.7) retrieves the same information as RopOpenMessage ROP ([MS-OXCROPS] section 2.2.6.1) but operates on an already opened Message object.");
            #endregion
            #endregion

            #region Call RopRelease to release all resources
            this.ReleaseRop(openedMessageHandle);
            #endregion
        }

        /// <summary>
        /// This test case validates error code ecNullObject (0x000004b9) of RopReloadCachedInformation operation.
        /// </summary>
        [TestCategory("MSOXCMSG"), TestMethod()]
        public void MSOXCMSG_S06_TC02_RopReloadCachedInformationFailure()
        {
            this.CheckMapiHttpIsSupported();
            this.ConnectToServer(ConnectionType.PrivateMailboxServer);

            #region Call RopLogon to log on a private mailbox.
            RopLogonResponse logonResponse = this.Logon(LogonType.Mailbox, out this.insideObjHandle);
            #endregion

            #region Call RopCreateMessage to create a new message object.
            uint targetMessageHandle = this.CreatedMessage(logonResponse.FolderIds[4], this.insideObjHandle);
            #endregion

            #region Call RopSaveChangesMessage to commit the new message object.
            RopSaveChangesMessageResponse saveChangesMessageResponse = this.SaveMessage(targetMessageHandle, (byte)SaveFlags.ForceSave);
            Site.Assert.AreEqual<uint>(TestSuiteBase.Success, saveChangesMessageResponse.ReturnValue, TestSuiteBase.ROPSucceedMsg);
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            this.ReleaseRop(targetMessageHandle);

            #region Call RopOpenMessage to open the created message.
            RopOpenMessageResponse openMessageResponse;
            uint openedMessageHandle = this.OpenSpecificMessage(logonResponse.FolderIds[4], messageId, this.insideObjHandle, MessageOpenModeFlags.ReadWrite, out openMessageResponse);
            #endregion

            #region Call RopReloadCachedInformation which contains an InputHandleIndex that does not refer to a message object.
            RopReloadCachedInformationRequest reloadCachedInformationRequest = new RopReloadCachedInformationRequest()
            {
                RopId = (byte)RopId.RopReloadCachedInformation,
                LogonId = CommonLogonId,
                InputHandleIndex = CommonInputHandleIndex,
                Reserved = 0x0000
            };
            this.response = new RopReloadCachedInformationResponse();
            this.ResponseSOHs = this.MSOXCMSGAdapter.DoRopCall(reloadCachedInformationRequest, TestSuiteBase.InvalidInputHandle, ref this.response, ref this.rawData, GetPropertiesFlags.None);
            RopReloadCachedInformationResponse reloadCachedInformationResponse = (RopReloadCachedInformationResponse)this.response;
            
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCMSG_R1498");

            // Verify MS-OXCMSG requirement: MS-OXCMSG_R1498
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000004b9,
                reloadCachedInformationResponse.ReturnValue,
                1498,
                @"[In Receiving a RopReloadCachedInformation ROP Request] [ecNullObject (0x000004b9)] means the value of the InputHandleIndex field on which this ROP [RopReloadCachedInformation] was called does not refer to a Message object.");
            #endregion

            #region Call RopRelease to release all resources
            this.ReleaseRop(openedMessageHandle);
            #endregion
        }

        #region Private methods
        /// <summary>
        /// Compare the information of Message object from RopOpenMessageResponse and RopReloadCachedInformationResponse.
        /// </summary>
        /// <param name="openResp">The response of RopOpenMessage.</param>
        /// <param name="reloadResp">The response of RopReloadCachedInformation.</param>
        /// <returns>A Boolean indicates whether the information is same.</returns>
        private bool CompareResponseOfOpenMessageAndRopReloadCachedInformation(RopOpenMessageResponse openResp, RopReloadCachedInformationResponse reloadResp)
        {
            if (openResp.HasNamedProperties != reloadResp.HasNamedProperties)
            {
                return false;
            }

            if (!openResp.SubjectPrefix.Equals(reloadResp.SubjectPrefix))
            {
                return false;
            }

            if (!openResp.NormalizedSubject.Equals(reloadResp.NormalizedSubject))
            {
                return false;
            }

            if (!openResp.RecipientCount.Equals(reloadResp.RecipientCount))
            {
                return false;
            }

            if (!openResp.ColumnCount.Equals(reloadResp.ColumnCount))
            {
                return false;
            }

            if (!openResp.RowCount.Equals(reloadResp.RowCount))
            {
                return false;
            }

            if ((openResp.RecipientColumns == null && reloadResp.RecipientColumns != null) || (openResp.RecipientColumns != null && reloadResp.RecipientColumns == null))
            {
                return false;
            }

            if (openResp.RecipientColumns != null && reloadResp.RecipientColumns != null)
            {
                if (openResp.RecipientColumns.Length != reloadResp.RecipientColumns.Length)
                {
                    return false;
                }

                for (int i = 0; i < openResp.RecipientColumns.Length; i++)
                {
                    if (!openResp.RecipientColumns[i].Equals(reloadResp.RecipientColumns[i]))
                    {
                        return false;
                    }
                }
            }

            if ((openResp.RecipientRows == null && reloadResp.RecipientRows != null) || (openResp.RecipientRows != null && reloadResp.RecipientRows == null))
            {
                return false;
            }

            if (openResp.RecipientRows != null && reloadResp.RecipientRows != null)
            {
                if (openResp.RecipientRows.Length != reloadResp.RecipientRows.Length)
                {
                    return false;
                }

                for (int i = 0; i < openResp.RecipientRows.Length; i++)
                {
                    if (!openResp.RecipientRows[i].Equals(reloadResp.RecipientRows[i]))
                    {
                        return false;
                    }
                }
            }

            return true;
        }
        #endregion
    }
}