namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This class is used to verify the ROP operations that the client has insufficient rights to operate on the specified private mailbox folder. 
    /// </summary>
    [TestClass]
    public class S05_InsufficientRightsOnFolder : TestSuiteBase
    {
        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

        /// <summary>
        /// The index of inbox of Exchange Server. Specified in section 2.2.1.1.3 in [MS-OXCSTOR].
        /// </summary>
        private int inboxIndex;

        /// <summary>
        /// SUT(system under test) computer name.
        /// </summary>
        private string sutServer;

        /// <summary>
        /// Define the name of domain where server belongs to.
        /// </summary>
        private string domain;

        /// <summary>
        /// CommonUser is the credential user that the permissions list can be modified by administrator.
        /// </summary>
        private string commonUser;

        /// <summary>
        /// The password for the CommonUser.
        /// </summary>
        private string commonUserPassword;

        /// <summary>
        /// A string that identifies the user configured by "CommonUser" who is making the EcDoConnectEx call. On Windows platform, this value is the value of legacyExchangeDN property on the user.
        /// </summary>
        private string commonUserEssdn;

        #region Test Suite Initialization

        /// <summary>
        /// Initialize the test suite
        /// </summary>
        /// <param name="testContext">The test context instance</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext testContext)
        {
            TestClassBase.Initialize(testContext);
        }

        /// <summary>
        /// Reset the test environment
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case is designed to validate that the RopCreateFolder operation under the condition that the user has insufficient right to create a folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S05_TC01_CreateFolderNoCreateSubfolderPermission()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();
            uint pidTagMemberRights;

            #region Step 1. Assign access permission for common user on the inbox and the root folder.
            uint inboxHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, inboxHandle);

            // Add folder visible permission for the root folder.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, this.RootFolderHandle);
            #endregion

            #region Step 2. Use the common user to logon the private mailbox.
            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(this.sutServer, ConnectionType.PrivateMailboxServer, this.commonUserEssdn, this.domain, this.commonUser, this.commonUserPassword);
            uint logonHandle;
            this.Logon(LogonFlags.Private, out logonHandle, (uint)OpenFlags.UsePerMDBReplipMapping);
            #endregion

            #region Step 3. The common user open the root folder.

            // Find and open the root folder.
            ulong commonUserRootFolderId = this.GetSubfolderIDByName(this.DefaultFolderIds[this.inboxIndex], logonHandle, this.RootFolder);
            uint commonUserRootFolderHandle = 0;
            this.OpenFolder(logonHandle, commonUserRootFolderId, ref commonUserRootFolderHandle);
            #endregion

            #region Step 4. The common user creates [MSOXCFOLDSubfolder1] in the root folder.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = 0x1C,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = 0x01,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x01,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.Subfolder1),
                Comment = Encoding.ASCII.GetBytes(Constants.Subfolder1)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, commonUserRootFolderHandle, ref this.responseHandles);

            #region Verify ecNoCreateSubfolderRight.
            if (Common.IsRequirementEnabled(106602, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R106602");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R106602.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000502,
                    createFolderResponse.ReturnValue,
                    106602,
                    @"[In Appendix A: Product Behavior] Implementation returns ecNoCreateSubfolderRight if the client does not have permissions to create the folder. <11> Section 3.2.5.2: Exchange 2010 returns ecNoCreateSubfolderRight.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1074");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1074.
                // MS-OXCFOLD_R106602 captured error code '0x00000502', MS-OXCFOLD_R1074 can be captured directly.
                Site.CaptureRequirement(
                    1074,
                    @"[In Processing a RopCreateFolder ROP Request] The value of error code ecNoCreateSubfolderRight is 0x00000502.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1075");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1075.
                // MS-OXCFOLD_R106602 captured error code '0x00000502', MS-OXCFOLD_R1075 can be captured directly.
                Site.CaptureRequirement(
                    1075,
                    "[In Processing a RopCreateFolder ROP Request] When the error code is ecNoCreateSubfolderRight, it indicates the client does not have access rights to create the folder.");
            }
            #endregion

            #region Verify error code ecAccessdenied.
            if (Common.IsRequirementEnabled(106601, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R106601");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R106601.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    createFolderResponse.ReturnValue,
                    106601,
                    @"[In Appendix A: Product Behavior] Implementation returns ecAccessdenied if the client does not have permissions to create the folder. <11> Section 3.2.5.2: Exchange 2007, Exchange 2013 and Exchange 2016 return ecAccessdenied.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1071");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1071.
                // MS-OXCFOLD_R106601 captured error code '0x80070005', MS-OXCFOLD_R1071 can be captured directly.
                Site.CaptureRequirement(
                    1071,
                    @"[In Processing a RopCreateFolder ROP Request] The value of error code ecAccessdenied is 0x80070005.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R489");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R489.
                // MS-OXCFOLD_R106601 captured error code '0x80070005', MS-OXCFOLD_R489 can be captured directly.
                Site.CaptureRequirement(
                    489,
                    "[In Processing a RopCreateFolder ROP Request] When the error code is ecAccessDenied, it indicates the client does not have permissions to create the folder.");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the client operates messages and subfolders under the condition that the user has insufficient rights to delete or move messages.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S05_TC02_MessageDeletePartialCompleteValidation()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            uint logonHandle = 0;
            uint pidTagMemberRights = 0;

            this.GenericFolderInitialization();

            #region Step 1. Add permission entry for the user configured by "CommonUser" on the inbox and the root folder.
            uint inboxHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, inboxHandle);

            // Add folder visible permission for the root folder.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny | (uint)PidTagMemberRightsEnum.EditOwned | (uint)PidTagMemberRightsEnum.Create | (uint)PidTagMemberRightsEnum.CreateSubFolder | (uint)PidTagMemberRightsEnum.DeleteOwned;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, this.RootFolderHandle);
            #endregion

            #region Step 2. Create message and subfolder in the root folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);

            uint subfolderHandle = 0;
            ulong subfolderId = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId, ref subfolderHandle);
            #endregion

            #region Step 3. Logon the user configured by "CommonUser".
            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(this.sutServer, ConnectionType.PrivateMailboxServer, this.commonUserEssdn, this.domain, this.commonUser, this.commonUserPassword);
            this.Logon(LogonFlags.Private, out logonHandle, (uint)OpenFlags.UsePerMDBReplipMapping);
            #endregion

            #region Step 4. The "CommonUser" opens the root folder.
            this.OpenFolder(logonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            ulong rootFolderId = this.GetSubfolderIDByName(this.DefaultFolderIds[this.inboxIndex], logonHandle, this.RootFolder);
            uint rootFolderHandle = 0;
            RopOpenFolderResponse openFolderResponse = this.OpenFolder(logonHandle, rootFolderId, ref rootFolderHandle);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R46202");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R46202
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                openFolderResponse.ReturnValue,
                46202,
                @"[In Processing a RopOpenFolder ROP Request] If the folder with the specified ID actually exists and the client has sufficient access rights to view the folder, the RopOpenFolder ROP performs successfully.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R394");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R394.
            // MS-OXCFOLD_R462 was verified, MS-OXCFOLD_R462 can be verified directly.
            Site.CaptureRequirement(
                394,
                @"[In Opening a Folder] The client MUST have sufficient access rights to the folder for this operation to succeed.");

            ulong commonUserMessageId = 0;
            uint commonUserMessageHandle = 0;
            this.CreateSaveMessage(rootFolderHandle, rootFolderId, ref commonUserMessageId, ref commonUserMessageHandle);

            uint commonUserSubfolderHandle = 0;
            ulong commonUserSubfolderId = 0;
            this.CreateFolder(rootFolderHandle, Constants.Subfolder2, ref commonUserSubfolderId, ref commonUserSubfolderHandle);
            #endregion

            #region Step 5. The client calls RopDeleteMessages to delete the message created in step 2.

            ulong[] messageIds = new ulong[] { messageId, commonUserMessageId };
            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest();
            RopDeleteMessagesResponse deleteMessagesResponse = new RopDeleteMessagesResponse();
            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = Constants.CommonLogonId;
            deleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteMessagesRequest.WantAsynchronous = 0x00;

            // The server does not generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = 0x00;
            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;
            deleteMessagesResponse = this.Adapter.DeleteMessages(deleteMessagesRequest, rootFolderHandle, ref this.responseHandles);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R991");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R991
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                deleteMessagesResponse.PartialCompletion,
                991,
                @"[In RopDeleteMessages ROP Response Buffer] PartialCompletion (1 byte): If the ROP [RopDeleteMessages] fails for a subset of targets, the value of this field is nonzero (TRUE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1147");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1147
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                deleteMessagesResponse.PartialCompletion,
                1147,
                @"[In Processing a RopDeleteMessages ROP Request] If the server fails to delete any messages, it sets the PartialCompletion field of the RopDeleteMessages ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.11.2.");
            #endregion

            #region Step 6. The client calls RopHardDeleteMessages to delete the message created in step 2.
            RopHardDeleteMessagesRequest hardDeleteMessagesRequest = new RopHardDeleteMessagesRequest
            {
                RopId = 0x91,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                NotifyNonRead = 0x00,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds
            };
            RopHardDeleteMessagesResponse hardDeleteMessagesResponse = this.Adapter.HardDeleteMessages(hardDeleteMessagesRequest, rootFolderHandle, ref this.responseHandles);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R997");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R997.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                hardDeleteMessagesResponse.PartialCompletion,
                997,
                @"[In RopHardDeleteMessages ROP Response Buffer] If the ROP [RopHardDeleteMessages] fails for a subset of targets, the value of this field is nonzero (TRUE).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R115602");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R115602
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                hardDeleteMessagesResponse.PartialCompletion,
                115602,
                @"[In Processing a RopHardDeleteMessages ROP Request] In the server behavior, if the server fails to delete any messages, it sets the PartialCompletion field of the RopDeleteMessages ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.11.2.");
            #endregion

            #region Step 7. The client calls RopHardDeleteMessagesAndSubfolders to delete the message created in step 2.
            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };
            RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse = this.Adapter.HardDeleteMessagesAndSubfolders(hardDeleteMessagesAndSubfoldersRequest, rootFolderHandle, ref this.responseHandles);
            
            if (Common.IsRequirementEnabled(2721, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2721");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2721
                Site.CaptureRequirementIfAreNotEqual<byte>(
                    0,
                    hardDeleteMessagesAndSubfoldersResponse.PartialCompletion,
                    2721,
                    @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] PartialCompletion (1 byte): If the ROP [RopHardDeleteMessagesAndSubfolders] fails for a subset of targets, the value of this field is nonzero (TRUE).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R113803");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R113803.
                Site.CaptureRequirement(
                    113803,
                    @"[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] In the server behavior, if the server fails to delete any message or subfolder, it sets the PartialCompletion field of the RopHardDeleteMessagesAndSubfolders ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.10.2.");
            }
            #endregion

            #region Step 8. The client calls RopMoveCopyMessages to move the message from the root folder to the inbox.
            List<uint> handleList = new List<uint>
            {
                rootFolderHandle, inboxHandle
            };

            RopMoveCopyMessagesRequest moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
            {
                RopId = (byte)RopId.RopMoveCopyMessages,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds,
                WantAsynchronous = 0x00,
                WantCopy = 0xFF
            };

            // WantCopy is nonzero (TRUE) indicates this is a copy operation.
            RopMoveCopyMessagesResponse moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handleList, ref this.responseHandles);

            if (Common.IsRequirementEnabled(586, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R10163");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R10163.
                Site.CaptureRequirementIfAreNotEqual<byte>(
                    0,
                    moveCopyMessagesResponse.PartialCompletion,
                    10163,
                    @"[In RopMoveCopyMessages ROP Response Buffer] PartialCompletion (1 byte): If the ROP fails for a subset of targets, the value of this field is nonzero (TRUE).");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R586");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R586.
                // MS-OXCFOLD_R10163 is verified, MS-OXCFOLD_R586 can be verified directly.
                Site.CaptureRequirement(
                    586,
                    @"[In Processing a RopMoveCopyMessages ROP Request] If the server fails to move or copy any message, it sets the PartialCompletion field of the RopMoveCopyMessages ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.6.2.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopEmptyFolder operation under the condition that the user has insufficient rights to empty a folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S05_TC03_RopEmptyFolderPartialCompleteValidation()
        {
            if (!Common.IsRequirementEnabled(1131, this.Site))
            {
                this.NeedCleanup = false;
                Site.Assert.Inconclusive("The implementation does not support to set nonzero value to PartialCompletion field of the RopEmptyFolder ROP response if the server fails to delete any message or subfolder by RopEmptyFolder ROP.");
            }

            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            uint pidTagMemberRights = 0;
            uint logonHandle = 0;
            this.GenericFolderInitialization();

            #region Step 1. Assign access permission for common user on the inbox and the root folder.
            uint inboxHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, inboxHandle);

            // Add folder visible permission for the root folder.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, this.RootFolderHandle);
            #endregion

            #region Step 2. Create a subfolder [MSOXCFOLDSubfolder1] in the root folder and assign full permission for common user on the new folder.
            uint adminSubfolderHandle1 = 0;
            ulong adminSubfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref adminSubfolderId1, ref adminSubfolderHandle1);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny | (uint)PidTagMemberRightsEnum.DeleteOwned | (uint)PidTagMemberRightsEnum.CreateSubFolder;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle1);
            #endregion

            #region Step 3. Create a subfolder [MSOXCFOLDSubfolder2] in the root folder and assign full permission for common user on the new folder.
            uint adminSubfolderHandle2 = 0;
            ulong adminSubfolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref adminSubfolderId2, ref adminSubfolderHandle2);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FullPermission;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle2);
            #endregion

            #region Step 4. Create a subfolder [MSOXCFOLDSubfolder3] in the [MSOXCFOLDSubfolder1] and assign full permission for common user on the new folder.
            uint adminSubfolderHandle3 = 0;
            ulong adminSubfolderId3 = 0;
            this.CreateFolder(adminSubfolderHandle1, Constants.Subfolder3, ref adminSubfolderId3, ref adminSubfolderHandle3);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FullPermission;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle3);
            #endregion

            #region Step 5. Create a subfolder [MSOXCFOLDSubfolder4] in the root folder.
            uint adminSubfolderHandle4 = 0;
            ulong adminSubfolderId4 = 0;
            this.CreateFolder(adminSubfolderHandle1, Constants.Subfolder4, ref adminSubfolderId4, ref adminSubfolderHandle4);
            #endregion

            #region Step 6. Logon to the private mailbox use common user.
            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(this.sutServer, ConnectionType.PrivateMailboxServer, this.commonUserEssdn, this.domain, this.commonUser, this.commonUserPassword);
            RopLogonResponse logonResponse = this.Logon(LogonFlags.Private, out logonHandle, (uint)OpenFlags.UsePerMDBReplipMapping);
            #endregion

            #region Step 7. The common user open the root folder, [MSOXCFOLDSubfolder1], [MSOXCFOLDSubfolder2] and [MSOXCFOLDSubfolder3].

            // Find and open the root folder.
            ulong commonUserRootFolderId = this.GetSubfolderIDByName(logonResponse.FolderIds[this.inboxIndex], logonHandle, this.RootFolder);
            uint commonUserRootFolderHandle = 0;
            this.OpenFolder(logonHandle, commonUserRootFolderId, ref commonUserRootFolderHandle);

            // Find and open the folder named [MSOXCFOLDSubfolder1].
            ulong commonUserSubfolder1 = this.GetSubfolderIDByName(commonUserRootFolderId, commonUserRootFolderHandle, Constants.Subfolder1);
            uint commonUserRootSubfolderHandle1 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolder1, ref commonUserRootSubfolderHandle1);

            // Find and open the folder named [MSOXCFOLDSubfolder2].
            ulong commonUserSubfolder2 = this.GetSubfolderIDByName(commonUserRootFolderId, commonUserRootFolderHandle, Constants.Subfolder2);
            uint commonUserRootSubfolderHandle2 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolder2, ref commonUserRootSubfolderHandle2);
            #endregion

            #region Step 5. Create a subfolder [MSOXCFOLDSubfolder4] in the root folder.
            uint commonUserSubfolderHandle5 = 0;
            ulong commonUserSubfolderId5 = 0;
            this.CreateFolder(commonUserRootSubfolderHandle1, Constants.Subfolder5, ref commonUserSubfolderId5, ref commonUserSubfolderHandle5);
            #endregion

            #region Step 8. The client calls RopEmptyFolder clear the root folder.
            RopEmptyFolderRequest emptyFolderRequest = new RopEmptyFolderRequest
            {
                RopId = (byte)RopId.RopEmptyFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0x00
            };

            // Invoke RopEmptyFolder operation to soft delete Subfolder3 from Subfolder1 without deleting Subfolder1.
            RopEmptyFolderResponse emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, commonUserRootSubfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.ReturnValue, "The RopEmptyFolder Rop operation performs successfully on [MSOXCFOLDSubfolder1]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R980");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R980.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                emptyFolderResponse.PartialCompletion,
                980,
                @"[In RopEmptyFolder ROP Response Buffer] If the ROP [RopEmptyFolder] fails for a subset of targets, the value of this field [PartialCompletion] is nonzero (TRUE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1131");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1131
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                emptyFolderResponse.PartialCompletion,
                1131,
                @"[In Processing a RopEmptyFolder ROP Request] If the server fails to delete any message or subfolder, it sets the PartialCompletion field of the RopEmptyFolder ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.9.2.");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveFolder operation under the condition that the user has insufficient rights to move a folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S05_TC04_RopMoveFolderPartialCompleteValidation()
        {
            if (!Common.IsRequirementEnabled(1114, this.Site))
            {
                this.NeedCleanup = false;
                Site.Assert.Inconclusive("The implementation does not support to set nonzero value to PartialCompletion field of the RopMoveFolder ROP response if the server fails to move any folder, message, or subfolder by RopMoveFolder ROP.");
            }

            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            uint pidTagMemberRights = 0;
            uint logonHandle = 0;
            this.GenericFolderInitialization();

            #region Step 1. Assign access permission for common user on the inbox and the root folder.
            uint inboxHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, inboxHandle);

            // Add folder visible permission for the root folder.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.EditOwned;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, this.RootFolderHandle);
            #endregion

            #region Step 2. Create a subfolder [MSOXCFOLDSubfolder1] in the root folder and assign FolderVisible permission for common user on the new folder.
            uint adminSubfolderHandle1 = 0;
            ulong adminSubfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref adminSubfolderId1, ref adminSubfolderHandle1);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle1);
            #endregion

            #region Step 3. Create a subfolder [MSOXCFOLDSubfolder2] in the root folder and assign full permission for common user on the new folder.
            uint adminSubfolderHandle2 = 0;
            ulong adminSubfolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref adminSubfolderId2, ref adminSubfolderHandle2);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FullPermission;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle2);
            #endregion

            #region Step 4. Create a subfolder [MSOXCFOLDSubfolder3] in the [MSOXCFOLDSubfolder1] and assign full permission for common user on the new folder.
            uint adminSubfolderHandle3 = 0;
            ulong adminSubfolderId3 = 0;
            this.CreateFolder(adminSubfolderHandle1, Constants.Subfolder3, ref adminSubfolderId3, ref adminSubfolderHandle3);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle3);
            #endregion

            #region Step 5. Create a subfolder [MSOXCFOLDSubfolder4] in the root folder.
            uint adminSubfolderHandle4 = 0;
            ulong adminSubfolderId4 = 0;
            this.CreateFolder(adminSubfolderHandle3, Constants.Subfolder4, ref adminSubfolderId4, ref adminSubfolderHandle4);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.None;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle4);
            #endregion

            #region Step 6. Logon to the private mailbox use common user.
            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(this.sutServer, ConnectionType.PrivateMailboxServer, this.commonUserEssdn, this.domain, this.commonUser, this.commonUserPassword);
            RopLogonResponse logonResponse = this.Logon(LogonFlags.Private, out logonHandle, (uint)OpenFlags.UsePerMDBReplipMapping);
            #endregion

            #region Step 7. The common user open the root folder, [MSOXCFOLDSubfolder1], [MSOXCFOLDSubfolder2] and [MSOXCFOLDSubfolder3].

            // Find and open the root folder.
            ulong commonUserRootFolderId = this.GetSubfolderIDByName(logonResponse.FolderIds[this.inboxIndex], logonHandle, this.RootFolder);
            uint commonUserRootFolderHandle = 0;
            this.OpenFolder(logonHandle, commonUserRootFolderId, ref commonUserRootFolderHandle);

            // Find and open the folder named [MSOXCFOLDSubfolder1].
            ulong commonUserSubfolderId1 = this.GetSubfolderIDByName(commonUserRootFolderId, commonUserRootFolderHandle, Constants.Subfolder1);
            uint commonUserRootSubfolderHandle1 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolderId1, ref commonUserRootSubfolderHandle1);

            // Find and open the folder named [MSOXCFOLDSubfolder2].
            ulong commonUserSubfolderId2 = this.GetSubfolderIDByName(commonUserRootFolderId, commonUserRootFolderHandle, Constants.Subfolder2);
            uint commonUserRootSubfolderHandle2 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolderId2, ref commonUserRootSubfolderHandle2);
            #endregion

            #region Step 8. The client calls RopMoveFolder to move target folder [MSOXCFOLDSubfolder1] from the root folder to destination folder [MSOXCFOLDSubfolder2].

            // Initialize a server object handle table.
            List<uint> handleList = new List<uint>
            {
                commonUserRootFolderHandle, commonUserRootSubfolderHandle2
            };

            // Call the RopMoveFolder operation to move the folder.
            RopMoveFolderRequest moveFolderRequest = new RopMoveFolderRequest
            {
                RopId = (byte)RopId.RopMoveFolder,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                WantAsynchronous = 0x00,
                UseUnicode = 0x00,
                FolderId = commonUserSubfolderId1,
                NewFolderName = Encoding.ASCII.GetBytes(Constants.Subfolder5)
            };
            RopMoveFolderResponse moveFolderResponse = this.Adapter.MoveFolder(moveFolderRequest, handleList, ref this.responseHandles);

            #region Verify RopMoveFolder PartialCompletion

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2191");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2191.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                moveFolderResponse.PartialCompletion,
                2191,
                @"[In RopMoveFolder ROP Response Buffer] PartialCompletion (1 byte): If the ROP fails for a subset of targets, the value of this field is nonzero (TRUE). ");
            
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1114");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1114.
            // The server fails to move any folder and MS-OXCFOLD_R2191 captured the PartialCompletion field was set to a nonzero value, MS-OXCFOLD_R1114 can be verified directly.
            Site.CaptureRequirement(
                1114,
                @"[In Processing a RopMoveFolder ROP Request] If the server fails to move any folder, message, or subfolder, it sets the PartialCompletion field of the RopMoveFolder ROP response buffer to nonzero (TRUE), as specified in section 2.2.1.7.2.");
            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteFolder operation under the condition that the user has insufficient rights to delete a folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S05_TC05_RopDeleteFolderPermissionValidation()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();
            uint pidTagMemberRights;
            uint logonHandle;

            #region Step 1. Assign access permission for common user on the inbox and the root folder.
            uint inboxHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxHandle);

            // Add folder visible permission for the inbox.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, inboxHandle);

            // Add folder visible permission for the root folder.
            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible | (uint)PidTagMemberRightsEnum.ReadAny;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, this.RootFolderHandle);
            #endregion

            #region Step 2. Create a subfolder [MSOXCFOLDSubfolder1] in the root folder and assign FolderVisible for common user on the new folder.
            uint adminSubfolderHandle1 = 0;
            ulong adminSubfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref adminSubfolderId1, ref adminSubfolderHandle1);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FolderVisible;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle1);
            #endregion

            #region Step 3. Create a subfolder [MSOXCFOLDSubfolder2] in the root folder and assign full permission for common user on the new folder.
            uint adminSubfolderHandle2 = 0;
            ulong adminSubfolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref adminSubfolderId2, ref adminSubfolderHandle2);

            pidTagMemberRights = (uint)PidTagMemberRightsEnum.FullPermission;
            this.AddPermission(this.commonUserEssdn, pidTagMemberRights, adminSubfolderHandle2);
            #endregion

            #region Step 4. The common user logon the private mailbox and open the root folder and [MSOXCFOLDSubfolder1].

            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(this.sutServer, ConnectionType.PrivateMailboxServer, this.commonUserEssdn, this.domain, this.commonUser, this.commonUserPassword);
            this.Logon(LogonFlags.Private, out logonHandle, (uint)OpenFlags.UsePerMDBReplipMapping);

            // Find and open the root folder.
            ulong commonUserRootFolderId = this.GetSubfolderIDByName(this.DefaultFolderIds[this.inboxIndex], logonHandle, this.RootFolder);
            uint commonUserRootFolderHandle = 0;
            this.OpenFolder(logonHandle, commonUserRootFolderId, ref commonUserRootFolderHandle);

            // Find and open the folder named [MSOXCFOLDSubfolder1].
            ulong commonUserSubfolder1 = this.GetSubfolderIDByName(commonUserRootFolderId, logonHandle, Constants.Subfolder1);
            uint commonUserRootSubfolderHandle1 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolder1, ref commonUserRootSubfolderHandle1);

            // Find and open the folder named [MSOXCFOLDSubfolder2].
            ulong commonUserSubfolder2 = this.GetSubfolderIDByName(commonUserRootFolderId, commonUserRootFolderHandle, Constants.Subfolder2);
            uint commonUserRootSubfolderHandle2 = 0;
            this.OpenFolder(logonHandle, commonUserSubfolder2, ref commonUserRootSubfolderHandle2);
            #endregion

            #region Step 5. Delete the folder [MSOXCFOLDSubfolder1], the expected error is ecAccessDenied.
            RopDeleteFolderRequest ropDeleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = 0x1D,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete,
                FolderId = commonUserSubfolder1
            };
            RopDeleteFolderResponse ropDeleteFolderResponse = this.Adapter.DeleteFolder(ropDeleteFolderRequest, commonUserRootFolderHandle, ref this.responseHandles);

            if (Common.IsRequirementEnabled(503, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R2500");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R2500
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    ropDeleteFolderResponse.ReturnValue,
                    2500,
                    @"[In Processing a RopDeleteFolder ROP Request] The value of error code ecAccessDenied is 0x80070005.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R503");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R503
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    ropDeleteFolderResponse.ReturnValue,
                    503,
                    @"[In Processing a RopDeleteFolder ROP Request] When the error code is ecAccessDenied, it indicates the client does not have permissions to delete this folder.");
            }

            if (Common.IsRequirementEnabled(408, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R408");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R408.
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    ropDeleteFolderResponse.PartialCompletion,
                    408,
                    @"[In Deleting a Folder] The PartialCompletion field of the ROP response, as specified in section 2.2.1.3.2, indicates whether there are any subfolders or messages that could not be deleted and, consequently, that the folder was not deleted.");
            }
            #endregion

            #region The common user delete the folder [MSOXCFOLDSubfolder2].
            ropDeleteFolderRequest = new RopDeleteFolderRequest
            {
                RopId = 0x1D,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                DeleteFolderFlags = (byte)DeleteFolderFlags.DeleteHardDelete,
                FolderId = commonUserSubfolder2
            };
            ropDeleteFolderResponse = this.Adapter.DeleteFolder(ropDeleteFolderRequest, commonUserRootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, ropDeleteFolderResponse.ReturnValue, "RopDeleteFolder ROP operation performs successfully!");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R405");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R405
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                ropDeleteFolderResponse.PartialCompletion,
                405,
                @"[In Deleting a Folder] To be deleted, a folder MUST exist, and the client needs the access rights to delete it.");
            #endregion
        }

        #region Test Case Initialization
        /// <summary>
        /// Test initialize. Overrides the method TestInitialize defined in base class.
        /// </summary>
        protected override void TestInitialize()
        {
            this.Adapter = Site.GetAdapter<IMS_OXCFOLDAdapter>();
            this.inboxIndex = Constants.InboxIndex;
            this.commonUser = Common.GetConfigurationPropertyValue("CommonUser", this.Site);
            this.sutServer = Common.GetConfigurationPropertyValue("SutComputerName", this.Site);
            this.domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.commonUserPassword = Common.GetConfigurationPropertyValue("CommonUserPassword", this.Site);
            this.commonUserEssdn = Common.GetConfigurationPropertyValue("CommonUserEssdn", this.Site);
            this.RootFolder = Common.GenerateResourceName(this.Site, Constants.RootFolder) + Constants.StringNullTerminated;
        }

        /// <summary>
        /// Overrides TestClassBase's TestCleanup()
        /// </summary>
        protected override void TestCleanup()
        {
            if (this.NeedCleanup == false)
            {
                return;
            }

            // Reinitialize to connect to the server use administrator for test cleanup.
            this.Adapter.DoDisconnect();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            uint logonHandle = 0;
            this.Logon(LogonFlags.Private, out logonHandle);
            ulong folderId = this.GetSubfolderIDByName(this.DefaultFolderIds[this.inboxIndex], logonHandle, this.RootFolder);

            if (folderId != 0)
            {
                uint inboxFolderHandle = 0;
                this.OpenFolder(logonHandle, this.DefaultFolderIds[this.inboxIndex], ref inboxFolderHandle);

                RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest();
                RopDeleteFolderResponse deleteFolderResponse;
                deleteFolderRequest.RopId = (byte)RopId.RopDeleteFolder;
                deleteFolderRequest.LogonId = Constants.CommonLogonId;
                deleteFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;

                // Set the DeleteFolderFlags to indicate hard delete the common generic folder,
                // including all folders and messages under the folder.
                deleteFolderRequest.DeleteFolderFlags = (byte)DeleteFolderFlags.DelFolders
                    | (byte)DeleteFolderFlags.DelMessages
                    | (byte)DeleteFolderFlags.DeleteHardDelete;
                deleteFolderRequest.FolderId = this.RootFolderId;

                int count = 0;
                bool rootFolderCleanUpSuccess = false;
                do
                {
                    deleteFolderResponse = this.Adapter.DeleteFolder(deleteFolderRequest, inboxFolderHandle, ref this.responseHandles);
                    if (deleteFolderResponse.ReturnValue == Constants.SuccessCode && deleteFolderResponse.PartialCompletion == 0)
                    {
                        rootFolderCleanUpSuccess = true;
                    }
                    else
                    {
                        Thread.Sleep(this.WaitTime);
                    }

                    if (count > this.RetryCount)
                    {
                        break;
                    }

                    count++;
                }
                while (!rootFolderCleanUpSuccess);
            }

            #region  RopRelease
            RopReleaseRequest releaseRequest = new RopReleaseRequest();
            object ropResponse = null;
            releaseRequest.RopId = (byte)RopId.RopRelease;
            releaseRequest.LogonId = Constants.CommonLogonId;
            releaseRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            this.Adapter.DoRopCall(releaseRequest, this.LogonHandle, ref ropResponse, ref this.responseHandles);
            #endregion

            this.responseHandles = null;
            this.Adapter.DoDisconnect();
            this.Adapter.Reset();
        }

        #endregion

        #region Private methods.

        /// <summary>
        /// Add a permission for a user to the permission list of the specified folder.
        /// </summary>
        /// <param name="commonUserEssdn">UserDN used to connect server.</param>
        /// <param name="memberRights">The permission value.</param>
        /// <param name="folderHandle">The specified folder object handle.</param>
        private void AddPermission(string commonUserEssdn, uint memberRights, uint folderHandle)
        {
            PermissionData[] permissionsDataArray = this.GetPermissionDataArrayForAdd(commonUserEssdn, memberRights);

            ModifyFlags modifyFlags = ModifyFlags.None;
            RopModifyPermissionsRequest modifyPermissionsRequest = this.CreateModifyPermissionsRequestBuffer(permissionsDataArray, modifyFlags);
            object response = new object();
            this.Adapter.DoRopCall(modifyPermissionsRequest, folderHandle, ref response, ref this.responseHandles);
            RopModifyPermissionsResponse modifyPermissionsResponse = (RopModifyPermissionsResponse)response;
            Site.Assert.AreEqual<uint>(0, modifyPermissionsResponse.ReturnValue, "0 indicates the server adds permission successfully.");
        }

        /// <summary>
        /// Create buffer to get ModifyPermissions
        /// </summary>
        /// <param name="permissionsDataArray">Permission data array is used to set permission</param>
        /// <param name="modifyFlags">Set the ModifyFlags, specified in [MS-OXCPERM] section 2.2.2</param>
        /// <returns>A request used to modify permissions</returns>
        private RopModifyPermissionsRequest CreateModifyPermissionsRequestBuffer(PermissionData[] permissionsDataArray, ModifyFlags modifyFlags)
        {
            RopModifyPermissionsRequest modifyPermissionsRequest = new RopModifyPermissionsRequest
            {
                RopId = (byte)RopId.RopModifyPermissions,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                ModifyFlags = (byte)modifyFlags,
                ModifyCount = (ushort)permissionsDataArray.Length,
                PermissionsData = permissionsDataArray
            };

            return modifyPermissionsRequest;
        }

        /// <summary>
        /// Set the permission data array for the specified user.
        /// </summary>
        /// <param name="userEssdn">The ESSDN of the specified user.</param>
        /// <param name="rights">The rights which will be assigned to the specified user.</param>
        /// <returns>The permission data array for the specified user.</returns>
        private PermissionData[] GetPermissionDataArrayForAdd(string userEssdn, uint rights)
        {
            PropertyValue[] propertyValues = new PropertyValue[2];
            propertyValues[0] = this.CreateRightsProperty(rights);
            propertyValues[1] = this.CreateEntryIdProperty(userEssdn);

            PermissionData[] permissionsDataArray = new PermissionData[1];
            permissionsDataArray[0].PermissionDataFlags = (byte)PermissionDataFlags.AddRow;
            permissionsDataArray[0].PropertyValueCount = (ushort)propertyValues.Length;
            permissionsDataArray[0].PropertyValues = propertyValues;

            return permissionsDataArray;
        }
        
        /// <summary>
        /// Create an instance of the property PidTagMemberRights.
        /// </summary>
        /// <param name="rights">The specified value of the property PidTagMemberRights.</param>
        /// <returns>The instance of the property PidTagMemberRights</returns>
        private TaggedPropertyValue CreateRightsProperty(uint rights)
        {
            TaggedPropertyValue rightsProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = (ushort)FolderPropertyId.PidTagMemberRights;
            temp.PropertyType = (ushort)PropertyType.PtypInteger32;
            rightsProperty.PropertyTag = temp;
            rightsProperty.Value = BitConverter.GetBytes(rights);

            return rightsProperty;
        }

        /// <summary>
        /// Generate the EntryId for modify the permissions.
        /// </summary>
        /// <param name="userEssdn">The user ESSDN.</param>
        /// <returns>TaggedPropertyValue indicate EntryId value.</returns>
        private TaggedPropertyValue CreateEntryIdProperty(string userEssdn)
        {
            TaggedPropertyValue entryIdProperty = new TaggedPropertyValue();
            PropertyTag temp;
            temp.PropertyId = (ushort)FolderPropertyId.PidTagEntryId;
            temp.PropertyType = (ushort)PropertyType.PtypBinary;
            entryIdProperty.PropertyTag = temp;

            entryIdProperty.VarLength = true;
            entryIdProperty.Value = this.GetEntryId(userEssdn);
            return entryIdProperty;
        }
        
        /// <summary>
        /// Get EntryId by user name.
        /// </summary>
        /// <param name="userEssdn">The user ESSDN.</param>
        /// <returns>EntryId in bytes which is retrieved by server.</returns>
        private byte[] GetEntryId(string userEssdn)
        {
            // Generate the Entry ID.
            if (string.IsNullOrEmpty(userEssdn))
            {
                return new byte[0];
            }

            string distinguishedName = userEssdn + Constants.StringNullTerminated;
            int pidEntryIdLength = 28 + distinguishedName.Length;
            byte[] pidEntryId = new byte[pidEntryIdLength];

            // Create the PidTagEntryId as PermanentEntryID described in section 2.3.8.3, [MS-OXNSPI]
            int i = 0;
            pidEntryId[i] = 0x00;
            i++;
            pidEntryId[i] = 0x00;
            i++;
            pidEntryId[i] = 0x00;
            i++;
            pidEntryId[i] = 0x00;
            i++;

            byte[] providerUID = new byte[16] { 0xDC, 0xA7, 0x40, 0xC8, 0xC0, 0x42, 0x10, 0x1A, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82 };
            Array.Copy(providerUID, 0, pidEntryId, i, 16);
            i += 16;

            byte[] r4 = BitConverter.GetBytes(0x00000001);
            Array.Copy(r4, 0, pidEntryId, i, 4);
            i += 4;

            byte[] displayTypeString = new byte[4] { 0, 0, 0, 0 };
            Array.Copy(displayTypeString, 0, pidEntryId, i, 4);
            i += 4;

            byte[] distinguishedNameBytes = System.Text.ASCIIEncoding.ASCII.GetBytes(distinguishedName);
            Array.Copy(distinguishedNameBytes, 0, pidEntryId, i, distinguishedName.Length);

            return pidEntryId;
        }

        #endregion
    }
}