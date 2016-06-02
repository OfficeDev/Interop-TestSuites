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
    /// This class is used to verify the ROP operations related to messages or subfolders in a folder object.
    /// </summary>
    [TestClass]
    public class S02_MessageRopOperations : TestSuiteBase
    {
        /// <summary>
        /// Server response handle list.
        /// </summary>
        private List<List<uint>> responseHandles;

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
        /// This test case is designed to validate that the RopMoveCopyMessages operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC01_RopMoveCopyMessagesSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            bool sourceMessageExist;
            bool sourceMessageRemoved;
            bool destinationMessageExist;
            ulong[] messageIds;
            List<uint> handlelist;

            #region Step 1. Create a message in the root folder.
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.
            ulong subfolderId1 = 0;
            uint subfolderHandle1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 3. Call RopMoveCopyMessages to copy the message created in step 1 from the root folder to [MSOXCFOLDSubfolder1] synchronously.
            messageIds = new ulong[1];
            messageIds[0] = messageId;
            handlelist = new List<uint>
            {
                this.RootFolderHandle, subfolderHandle1
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
            RopMoveCopyMessagesResponse moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.PartialCompletion, "The ROP successes for all subsets of targets");
            handlelist.Clear();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1105");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1105.
            // The return value 0 of MoveCopyMessages indicates that the server responds with a RopMoveCopyMessages ROP response buffer.
            Site.CaptureRequirement(
                1105,
                @"[In Processing a RopMoveCopyMessages ROP Request] The server responds with a RopMoveCopyMessages ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R155");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R155.
            // The WantAsynchronous was set to zero and the server responds a RopMoveCopyMessages ROP response indicates the ROP is processed synchronously, MS-OXCFOLD_R155 can be verified directly.
            Site.CaptureRequirement(
                155,
                @"[In RopMoveCopyMessages ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            #endregion

            #region Step 4. Validate the message is copied successfully in step 3.
            uint rootFolderContentsCountExpect = 1;
            uint rootFolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, this.RootFolderHandle);
            sourceMessageExist = rootFolderContentsCountExpect == rootFolderContentsCountActual;

            uint subfolderContentsCountExpect = 1;
            uint subfolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, subfolderHandle1);
            destinationMessageExist = subfolderContentsCountExpect == subfolderContentsCountActual;

            #region Verify MS-OXCFOLD_R159, MS-OXCFOLD_R14902, MS-OXCFOLD_R15102 and MS-OXCFOLD_R14501.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R159:
                Expected contents count of the root folder is {0},
                Actual contents count of the root folder is {1};
                Expected contents count of the [MSOXCFOLDSubfolder1] is {2},
                Actual contents count of the [MSOXCFOLDSubfolder1] is {3};",
                rootFolderContentsCountExpect,
                rootFolderContentsCountActual,
                subfolderContentsCountExpect,
                subfolderContentsCountActual);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R159.
            bool isVerifyR159 = sourceMessageExist && destinationMessageExist;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR159,
                159,
                @"[In RopMoveCopyMessages ROP Request Buffer] WantCopy (1 byte): A Boolean value that is nonzero (TRUE) if this [RopMoveCopyMessages ROP] is a copy operation.");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R14902.
            bool isVerifyR14902 = sourceMessageExist && destinationMessageExist;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R14902: expected contents count of the root folder is {0}, actual contents count of the root folder is {1}; expected contents count of the [MSOXCFOLDSubfolder1] is {2}, actual contents count of the [MSOXCFOLDSubfolder1] is {3};", rootFolderContentsCountExpect, rootFolderContentsCountActual, subfolderContentsCountExpect, subfolderContentsCountActual);

            // The source handle lists contain RootFolderHandle and subfolderHandle1, and the sourceHandleIndex is 0 and the destHandleIndex is 1 in the moveCopyMessage request, so if the message was copied successfully from root folder to the subfolder1, R14902 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR14902,
                14902,
                @"[In RopMoveCopyMessages ROP Request Buffer] SourceHandleIndex (1 byte): The source Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder from which the messages will be copied.");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R15102.
            bool isVerifyR15102 = sourceMessageExist && destinationMessageExist;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R15102: expected contents count of the root folder is {0}, actual contents count of the root folder is {1}; expected contents count of the [MSOXCFOLDSubfolder1] is {2}, actual contents count of the [MSOXCFOLDSubfolder1] is {3};", rootFolderContentsCountExpect, rootFolderContentsCountActual, subfolderContentsCountExpect, subfolderContentsCountActual);

            // The source handle lists contain RootFolderHandle and subfolderHandle1, and the sourceHandleIndex is 0 and the destHandleIndex is 1 in the moveCopyMessage request, so if the message was copied successfully from root folder to the subfolder1, R15102 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR15102,
                15102,
                @"[In RopMoveCopyMessages ROP Request Buffer] DestHandleIndex (1 byte): The destination Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder to which the messages will be copied.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R14501");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R14501.
            // MS-OXCFOLD_R159 is verified and this scenario use private mailbox logon, MS-OXCFOLD_R14501 can be verified directly.
            Site.CaptureRequirement(
                14501,
                @"[In RopMoveCopyMessages ROP] This ROP applies to [both public folders and] private mailboxes.");
            #endregion
            #endregion

            #region Step 5. Call RopMoveCopyMessages to move the message created in step 1 from the root folder to [MSOXCFOLDSubfolder1] synchronously.
            messageIds = new ulong[] { messageId };
            handlelist = new List<uint>
            {
                this.RootFolderHandle, subfolderHandle1
            };

            moveCopyMessagesRequest = new RopMoveCopyMessagesRequest();
            moveCopyMessagesRequest.RopId = (byte)RopId.RopMoveCopyMessages;
            moveCopyMessagesRequest.LogonId = Constants.CommonLogonId;
            moveCopyMessagesRequest.SourceHandleIndex = 0x00;
            moveCopyMessagesRequest.DestHandleIndex = 0x01;
            moveCopyMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            moveCopyMessagesRequest.MessageIds = messageIds;
            moveCopyMessagesRequest.WantAsynchronous = 0x00;

            // WantCopy is zero (FALSE) indicates this is a move operation.
            moveCopyMessagesRequest.WantCopy = 0x00;
            moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs successfully.");

            handlelist.Clear();

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R163");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R163.
            // According to the test case designing and open specification, the RopMoveCopyMessagesResponse operation here does not fail for a subset of targets, MS-OXCFOLD_R163 can be captured.
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                moveCopyMessagesResponse.PartialCompletion,
                163,
                @"[In RopMoveCopyMessages ROP Response Buffer]PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value is zero (FALSE).");
            #endregion

            #region Step 6. Validate the message is moved successfully in step 5.
            rootFolderContentsCountExpect = 0;
            rootFolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, this.RootFolderHandle);
            sourceMessageRemoved = rootFolderContentsCountExpect == rootFolderContentsCountActual;

            subfolderContentsCountExpect = 2;
            subfolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, subfolderHandle1);
            destinationMessageExist = subfolderContentsCountExpect == subfolderContentsCountActual;

            #region Verify MS-OXCFOLD_R160, MS-OXCFOLD_R14901 and MS-OXCFOLD_R15101.
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R160: The validation result of whether the source message was removed is {0}, the validation result of whether the destination message is existed is {1}",
                sourceMessageRemoved,
                destinationMessageExist);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R160.
            bool isVerifyR160 = sourceMessageRemoved && destinationMessageExist;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR160,
                160,
                @"[In RopMoveCopyMessages ROP Request Buffer] WantCopy (1 byte): [A Boolean value that is] zero (FALSE) if this [RopMoveCopyMessages ROP] is a move operation.");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R14901.
            bool isVerifyR14901 = sourceMessageRemoved && destinationMessageExist;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R14901: the validation result of whether the source message was removed is {0}, the validation result of whether the destination message is existed is {1}", sourceMessageRemoved, destinationMessageExist);

            // The source handle lists contain RootFolderHandle and subfolderHandle1, and the sourceHandleIndex is 0 and the destHandleIndex is 1 in the moveCopyMessage request, so if the message was moved successfully from root folder to the subfolder1, R14901 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR14901,
                14901,
                @"[In RopMoveCopyMessages ROP Request Buffer] SourceHandleIndex (1 byte): The source Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder from which the messages will be moved.");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R15101.
            bool isVerifyR15101 = sourceMessageRemoved && destinationMessageExist;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R15101: the validation result of whether the source message was removed is {0}, the validation result of whether the destination message is existed is {1}", sourceMessageRemoved, destinationMessageExist);

            // The source handle lists contain RootFolderHandle and subfolderHandle1, and the sourceHandleIndex is 0 and the destHandleIndex is 1 in the moveCopyMessage request, so if the message was moved successfully from root folder to the subfolder1, R15101 can be verified.
            Site.CaptureRequirement(
                15101,
                @"[In RopMoveCopyMessages ROP Request Buffer] DestHandleIndex (1 byte): The destination Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder to which the messages will be moved.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopEmptyFolder operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC02_RopEmptyFolderSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region  Step 2. Call RopCreateFolder to create [MSOXCFOLDSubfolder3] under [MSOXCFOLDSubfolder1].
            uint subfolderHandle3 = 0;
            ulong subfolderId3 = 0;
            this.CreateFolder(subfolderHandle1, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 3. Creates a message in [MSOXCFOLDSubfolder1].
            ulong messageId = 0;
            uint messageHandle = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);
            #endregion

            #region Step 4. Call RopEmptyFolder to empty [MSOXCFOLDSubfolder1] synchronously.
            RopEmptyFolderRequest emptyFolderRequest = new RopEmptyFolderRequest
            {
                RopId = (byte)RopId.RopEmptyFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0x00
            };

            // Invoke RopEmptyFolder operation to soft delete Subfolder3 from Subfolder1 without deleting Subfolder1.
            RopEmptyFolderResponse emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, subfolderHandle1, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.ReturnValue, "The RopEmptyFolder Rop operation performs successfully on [MSOXCFOLDSubfolder1]");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1126");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1126.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopEmptyFolderResponse),
                emptyFolderResponse.GetType(),
                1126,
                @"[In Processing a RopEmptyFolder ROP Request] The server responds with a RopEmptyFolder ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R235");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R235.
            // The WantAsynchronous was set to zero and the server responds a RopEmptyFolder ROP response indicates the ROP is processed synchronously, MS-OXCFOLD_R235 can be verified directly.
            Site.CaptureRequirement(
                235,
                @"[In RopEmptyFolder ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            #region Verify MS-OXCFOLD_R97502 and MS-OXCFOLD_R243 and MS-OXCFOLD_R225.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R97502");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R97502
            // The RopEmptyFolder ROP operation performs successfully on a private mailbox, MS-OXCFOLD_R97502 can be verified directly.
            Site.CaptureRequirement(
                97502,
                @"[In RopEmptyFolder ROP] This ROP [RopEmptyFolder] applies to [both public folders and] private mailboxes.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R243");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R243
            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                emptyFolderResponse.PartialCompletion,
                243,
                @"[In RopEmptyFolder ROP Response Buffer]PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value [of PartialCompletion field] is zero (FALSE).");

            bool isDelete = this.IsFolderDeleted(subfolderId1);
            
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R225");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R235.
            Site.CaptureRequirementIfIsFalse(
                isDelete,
                225,
                @"[In RopEmptyFolder ROP] The RopEmptyFolder ROP ([MS-OXCROPS] section 2.2.4.9) is used to soft delete messages and subfolders from a folder without deleting the folder itself.");

            #endregion

            #endregion

            #region Step 5. Call RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder.
            uint subfolderHandle2 = 0;
            ulong subfolderId2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);
            #endregion

            #region Step 6. Call RopCreateFolder to create [MSOXCFOLDSubfolder3] under [MSOXCFOLDSubfolder2].
            this.CreateFolder(subfolderHandle2, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 7. Creates a message in [MSOXCFOLDSubfolder2].
            this.CreateSaveMessage(subfolderHandle2, subfolderId2, ref messageId, ref messageHandle);
            #endregion

            #region Step 8. Creates a FAI message in [MSOXCFOLDSubfolder2].
            uint faiMessageHandle = 0;
            ulong faiMessageId = 0;
            this.CreateSaveMessage(subfolderHandle2, subfolderId2, 0x01, ref faiMessageId, ref faiMessageHandle);
            #endregion

            #region Step 9. The client calls RopEmptyFolder on [MSOXCFOLDSubfolder2] with WantDeleteAssociated set to zero (FALSE).
            object ropResponse = null;
            emptyFolderRequest.RopId = (byte)RopId.RopEmptyFolder;
            emptyFolderRequest.LogonId = Constants.CommonLogonId;
            emptyFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            emptyFolderRequest.WantAsynchronous = 0x00;
            emptyFolderRequest.WantDeleteAssociated = 0x00;
            this.Adapter.DoRopCall(emptyFolderRequest, subfolderHandle2, ref ropResponse, ref this.responseHandles);
            emptyFolderResponse = (RopEmptyFolderResponse)ropResponse;

            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.ReturnValue, "The RopEmptyFolder ROP operation performs successfully on [MSOXCFOLDSubfolder2].");
            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.PartialCompletion, "The ROP successes for all subsets of targets, the value of PartialCompletion field is zero (FALSE).");

            bool isSubfolerRemovedWithoutWantDeleteAssociated = this.IsFolderDeleted(subfolderId3);
            bool isNormalMessageRemovedWithoutWantDeleteAssociated = this.IsMessageDeleted(messageId, subfolderId2);

            bool faiMessageDeleted = this.IsMessageDeleted(faiMessageId, subfolderId2);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R239");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R239.
            Site.CaptureRequirementIfIsFalse(
                faiMessageDeleted,
                239,
                @"[In RopEmptyFolder ROP Request Buffer] WantDeleteAssociated (1 byte): The value is zero (FALSE) otherwise [if the folder associated information (FAI) messages are not to be included in the deletion].");
            #endregion

            #region Step 10. Creates a message in [MSOXCFOLDSubfolder2].
            this.CreateSaveMessage(subfolderHandle2, subfolderId2, ref messageId, ref messageHandle);
            #endregion

            #region Step 11. Call RopCreateFolder to create [MSOXCFOLDSubfolder3] under [MSOXCFOLDSubfolder2].
            this.CreateFolder(subfolderHandle2, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 12. The client calls RopEmptyFolder on [MSOXCFOLDSubfolder2] with WantDeleteAssociated set to nonzero (TRUE).
            ropResponse = null;
            emptyFolderRequest.RopId = (byte)RopId.RopEmptyFolder;
            emptyFolderRequest.LogonId = Constants.CommonLogonId;
            emptyFolderRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            emptyFolderRequest.WantAsynchronous = 0x00;
            emptyFolderRequest.WantDeleteAssociated = 0x01;
            this.Adapter.DoRopCall(emptyFolderRequest, subfolderHandle2, ref ropResponse, ref this.responseHandles);
            emptyFolderResponse = (RopEmptyFolderResponse)ropResponse;
            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.ReturnValue, "RopEmptyFolder ROP operation performs Successfully.");
            Site.Assert.AreEqual<uint>(0, emptyFolderResponse.PartialCompletion, "The RopEmptyFolder ROP operation performs Successfully for all subset of targets, the value of PartialCompletion field is zero (FALSE).");

            #region Verify MS-OXCFOLD_R617 and MS-OXCFOLD_238, MS-OXCFOLD_R618 and MS-OXCFOLD_R43201.

            faiMessageDeleted = this.IsMessageDeleted(faiMessageId, subfolderId2);
            bool isNormalMessageRemovedWithWantDeleteAssociated = this.IsMessageDeleted(messageId, subfolderId2);

            // Add the debug information.
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R617: The validation result of whether the fai message is existed is {0}",
                faiMessageDeleted);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R617.
            Site.CaptureRequirementIfIsTrue(
                faiMessageDeleted,
                617,
                @"[In Processing a RopEmptyFolder ROP Request]If the WantDeleteAssociated field of the RopEmptyFolder ROP request buffer is set to nonzero (TRUE), as specified in section 2.2.1.9.1, then the server removes all FAI messages in addition to the normal messages.");

            // Add the debug information.
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R238:
                The validation result of whether the normal message was deleted when the WantDeleteAssociated was not set is: {0};
                The validation result of whether the normal message was deleted when the WantDeleteAssociated was set is: {1};
                The validation result of whether the FAI message was deleted when the WantDeleteAssociated was set is: {2};",
                isNormalMessageRemovedWithoutWantDeleteAssociated,
                isNormalMessageRemovedWithWantDeleteAssociated,
                faiMessageDeleted);

            bool isVerifiedR238 = isNormalMessageRemovedWithoutWantDeleteAssociated && isNormalMessageRemovedWithWantDeleteAssociated && faiMessageDeleted;

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R238.
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR238,
                238,
                @"[In RopEmptyFolder ROP Request Buffer] WantDeleteAssociated (1 byte): A Boolean value that is nonzero (TRUE) if the folder associated information (FAI) messages are to be included in the deletion.");

            bool isSubfolerRemovedWithWantDeleteAssociated = this.IsFolderDeleted(subfolderId3);

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R618: The validation result of whether the subfolder was removed when the faiMessageDeleted was not set is {0}; The validation result of whether the subfolder was removed when the faiMessageDeleted was set is {1};",
                isSubfolerRemovedWithoutWantDeleteAssociated,
                isSubfolerRemovedWithWantDeleteAssociated);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R618.
            Site.CaptureRequirementIfIsTrue(
                isSubfolerRemovedWithoutWantDeleteAssociated && isSubfolerRemovedWithWantDeleteAssociated,
                618,
                @"[In Processing a RopEmptyFolder ROP Request]The server removes all subfolders regardless of the value of the WantDeleteAssociated field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R43201");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R43201.
            // The RopEmptyFolder ROP operation is verified, MS-OXCFOLD_R43201 can be verified directly.
            Site.CaptureRequirement(
                43201,
                @"[In Deleting the Contents of a Folder] To delete all messages and subfolders from a folder without deleting the folder itself, the client sends either a RopEmptyFolder ROP request ([MS-OXCROPS] section 2.2.4.9) [or a RopHardDeleteMessagesAndSubfolders ROP request ([MS-OXCROPS] section 2.2.4.10)].");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessagesAndSubfolders operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC03_RopHardDeleteMessagesAndSubfoldersSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] in the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 2. Call RopCreateFolder to create [MS-OXCFOLDSubfolder3] in the [MSOXCFOLDSubfolder1].
            uint subfolderHandle3 = 0;
            ulong subfolderId3 = 0;
            this.CreateFolder(subfolderHandle1, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 3. Create message in [MSOXCFOLDSubfolder1].
            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);
            #endregion

            #region Step 4. Call RopHardDeleteMessagesAndSubfolders applying to the [MSOXCFOLDSubfolder1] synchronously.
            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };
            RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse = this.Adapter.HardDeleteMessagesAndSubfolders(hardDeleteMessagesAndSubfoldersRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, hardDeleteMessagesAndSubfoldersResponse.ReturnValue, "If ROP succeeds, ReturnValue of its response will be 0 (success)");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R261");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R261.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                hardDeleteMessagesAndSubfoldersResponse.PartialCompletion,
                261,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value [of PartialCompletion field] is zero (FALSE).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1135");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1135
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopHardDeleteMessagesAndSubfoldersResponse),
                hardDeleteMessagesAndSubfoldersResponse.GetType(),
                1135,
                @"[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] The server responds with a RopHardDeleteMessagesAndSubfolders ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R98302");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R98302.
            Site.CaptureRequirementIfAreEqual<uint>(
                Constants.SuccessCode,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                98302,
                @"[In RopHardDeleteMessagesAndSubfolders ROP] This ROP [RopHardDeleteMessagesAndSubfolders] applies to [both public folders and] private mailboxes.");

            #region Verify MS-OXCFOLD_R251,MS-OXCFOLD_113801 and MS-OXCFOLD_R43202

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R251");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R251.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopHardDeleteMessagesAndSubfoldersResponse),
                hardDeleteMessagesAndSubfoldersResponse.GetType(),
                251,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");
 
            bool isTargetFolderSoftDeleted = this.IsFolderSoftDeleted(subfolderId1);
            bool isTargetFolderHardDeleted = this.IsFolderHardDeleted(subfolderId1);
            bool isTargetFolderExist = !isTargetFolderSoftDeleted && !isTargetFolderHardDeleted;

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R113801: The validation result of whether the target folder was removed is {0}.", isTargetFolderExist);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R113801.
            // MS-OXCFOLD_R434 is verified, if the target folder of RopHardDeleteMessagesAndSubfolders is exist, MS-OXCFOLD_R113801 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isTargetFolderExist,
                113801,
                @"[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] In the server behavior, the server hard deletes the folder's messages and subfolders but does not delete the folder itself.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R43202");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R43202.
            // MS-OXCFOLD_R113801 is verified, MS-OXCFOLD_R113801 can be verified directly.
            Site.CaptureRequirement(
                43202,
                @"[In Deleting the Contents of a Folder] To delete all messages and subfolders from a folder without deleting the folder itself, the client sends [either a RopEmptyFolder ROP request ([MS-OXCROPS] section 2.2.4.9)] or a RopHardDeleteMessagesAndSubfolders ROP request ([MS-OXCROPS] section 2.2.4.10).");
            #endregion
            #endregion

            #region Step 5. Create [MSOXCFOLDSubfolder3] in [MSOXCFOLDSubfolder1].
            this.CreateFolder(subfolderHandle1, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 6. Create a FAI message in [MSOXCFOLDSubfolder1].
            uint faiMessageHandle = 0;
            ulong faiMessageId = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, 0x01, ref faiMessageId, ref faiMessageHandle);
            #endregion

            #region Step 7. Create a normal message in [MSOXCFOLDSubfolder1].
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);
            #endregion

            #region Step 8. Call RopHardDeleteMessagesAndSubfolders applying to the [MSOXCFOLDSubfolder1] synchronously, the FAI message are not to be included in the deletion.
            hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0x00
            };

            // The field WantDeleteAssociated is zero (FALSE) indicates the FAI messages are not to be included in the deletion.
            hardDeleteMessagesAndSubfoldersResponse = this.Adapter.HardDeleteMessagesAndSubfolders(hardDeleteMessagesAndSubfoldersRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(
                0,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                "If ROP succeeds, ReturnValue of its response will be 0 (success)");
            Site.Assert.AreEqual<byte>(
                0,
                hardDeleteMessagesAndSubfoldersResponse.PartialCompletion,
                "If delete succeeds, PartialCompletion of its response will be 0 (success)");

            bool isNormalMessageRemovedWithWantDeleteAssociated = this.IsMessageDeleted(messageId, subfolderId1);
            bool isSubfolderRemovedWithWantDeleteAssociated = this.IsFolderHardDeleted(subfolderId3);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R257");

            bool faiMessageDeleted = this.IsMessageDeleted(faiMessageId, subfolderId1);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R257
            // If the FAI message is not deleted, MS-OXCFOLD_R257 can be verified.
            Site.CaptureRequirementIfIsFalse(
                faiMessageDeleted,
                257,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Request Buffer] WantDeleteAssociated (1 byte): The value is zero (FALSE) otherwise [if the FAI messages are not to be included in the deletion].");
            #endregion

            #region Step 9. Create [MSOXCFOLDSubfolder3] in [MSOXCFOLDSubfolder1].
            this.CreateFolder(subfolderHandle1, Constants.Subfolder3, ref subfolderId3, ref subfolderHandle3);
            #endregion

            #region Step 10. Create a normal message in [MSOXCFOLDSubfolder1].
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);
            #endregion

            #region Step 11. Call RopHardDeleteMessagesAndSubfolders applying to the [MSOXCFOLDSubfolder1] synchronously, the FAI message to be included in the deletion.
            hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };
            hardDeleteMessagesAndSubfoldersResponse = this.Adapter.HardDeleteMessagesAndSubfolders(hardDeleteMessagesAndSubfoldersRequest, subfolderHandle1, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, hardDeleteMessagesAndSubfoldersResponse.ReturnValue, "If ROP succeeds, ReturnValue of its response will be 0 (success)");
            Site.Assert.AreEqual<uint>(0, hardDeleteMessagesAndSubfoldersResponse.PartialCompletion, "If delete all subsets of targets succeeds, PartialCompletion of its response will be 0 (success)");

            #region Verify MS-OXCFOLD_R256,MS-OXCFOLD_R113804 and MS-OXCFOLD_R113805.

            // If the FAI message is hard deleted, MS-OXCFOLD_R256 can be verified.
            bool isVerifiedR256 = this.IsMessageDeleted(faiMessageId, subfolderId1);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R256: The validation result of whether the FAI message was deleted when the WantDeleteAssociated was set is {0}", isVerifiedR256);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R256
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR256,
                256,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Request Buffer] WantDeleteAssociated (1 byte): A Boolean value that is nonzero (TRUE) if the FAI messages are to be included in the deletion.");

            bool isNormalMessageRemovedWithoutWantDeleteAssociated = this.IsMessageDeleted(messageId, subfolderId1);

            // Add the debug information.
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCFOLD_R113804:
                The validation result of whether the normal message was deleted when the WantDeleteAssociated was not set is {0};
                The validation result of whether the normal message was deleted when the WantDeleteAssociated was set is {1}",
                isNormalMessageRemovedWithWantDeleteAssociated,
                isNormalMessageRemovedWithoutWantDeleteAssociated);

            // MS-OXCFOLD_R256 is verified and the FAI message is hard deleted, if normal message is hard deleted, MS-OXCFOLD_R113804 can be captured.
            Site.CaptureRequirementIfIsTrue(
                isNormalMessageRemovedWithWantDeleteAssociated && isNormalMessageRemovedWithoutWantDeleteAssociated,
                113804,
                "[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] In the server behavior, if the WantDeleteAssociated field of the RopHardDeleteMessagesAndSubfolders ROP request buffer is set to nonzero (TRUE), as specified in section 2.2.1.10.1, then the server removes all FAI messages in addition to the normal messages.");
  
            bool isSubfolderRemovedWithoutWantDeleteAssociated = this.IsFolderHardDeleted(subfolderId3);

            if (Common.IsRequirementEnabled(46201002, this.Site))
            {
                bool isVerifiedR113805 = isSubfolderRemovedWithoutWantDeleteAssociated && isSubfolderRemovedWithWantDeleteAssociated;

                // Add the debug information.
                Site.Log.Add(
                    LogEntryKind.Debug,
                    @"Verify MS-OXCFOLD_R113805:
                    The validation result of whether the subfolder was removed when the WantDeleteAssociated was not set is {0}
                    The validation result of whether the subfolder was removed when the WantDeleteAssociated was set is {1}",
                    isSubfolderRemovedWithoutWantDeleteAssociated,
                    isSubfolderRemovedWithWantDeleteAssociated);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R113805
                Site.CaptureRequirementIfIsTrue(
                    isVerifiedR113805,
                    113805,
                    @"[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] In the server behavior, the server removes all subfolders regardless of the value of the WantDeleteAssociated field.");
            }
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessagesAndSubfolders operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC04_RopDeleteMessagesSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Create tree message in root folder.
            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);
            uint messageHandle2 = 0;
            ulong messageId2 = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId2, ref messageHandle2);
            uint messageHandle3 = 0;
            ulong messageId3 = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId3, ref messageHandle3);
            #endregion

            #region Step 2. Delete the message created in step 1 in root folder.
            this.OpenMessage(messageId, this.RootFolderHandle, MessageOpenModeFlags.ReadWrite);
            object ropResponse = null;
            ulong[] messageIds = new ulong[] { messageId };

            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest();
            RopDeleteMessagesResponse deleteMessagesResponse;
            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = Constants.CommonLogonId;
            deleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteMessagesRequest.WantAsynchronous = 0x00;

            // The server does not generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = 0x00;
            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;
            this.Adapter.DoRopCall(deleteMessagesRequest, this.RootFolderHandle, ref ropResponse, ref this.responseHandles);
            deleteMessagesResponse = (RopDeleteMessagesResponse)ropResponse;
            Site.Assert.AreEqual<uint>(
                0,
                deleteMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1143");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1143
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(RopDeleteMessagesResponse),
                ropResponse.GetType(),
                1143,
                @"[In Processing a RopDeleteMessages ROP Request] The server responds with a RopDeleteMessages ROP response buffer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R98802");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R98802.
            // The RopDeleteMessages ROP operation performs successfully on a private mailbox, MS-OXCFOLD_R98802 can be captured directly.
            Site.CaptureRequirement(
                98802,
                @"[In RopDeleteMessages ROP] This ROP [RopDeleteMessages] applies to [both public folders and] private mailboxes.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R271");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R271.
            // The WantAsynchronous was set to zero and the server responds a RopDeleteMessages ROP response indicates the ROP is processed synchronously, MS-OXCFOLD_R271 can be verified directly.
            Site.CaptureRequirement(
                271,
                @"[In RopDeleteMessages ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            #region Verify MS-OXCFOLD_R282 and MS-OXCFOLD_R436.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R282");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R282.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                deleteMessagesResponse.PartialCompletion,
                282,
                "[In RopDeleteMessages ROP Response Buffer] PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value [of PartialCompletion field] is zero (FALSE).");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R436");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R436.
            // The RopDeleteMessage performs successfully, and MS-OXCFOLD_R282 is verified, MS-OXCFOLD_R436 can be verified directly.
            Site.CaptureRequirement(
                436,
                "[In Deleting the Contents of a Folder] To remove particular messages from a folder, the client sends [either] a RopDeleteMessages ROP request ([MS-OXCROPS] section 2.2.4.11) [or a RopHardDeleteMessages ROP request ([MS-OXCROPS] section 2.2.4.12)].");
            #endregion

            messageIds = new ulong[] { messageId2, messageId3 };
            deleteMessagesRequest.RopId = (byte)RopId.RopDeleteMessages;
            deleteMessagesRequest.LogonId = Constants.CommonLogonId;
            deleteMessagesRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            deleteMessagesRequest.WantAsynchronous = 0x00;

            // The server does not generate a non-read receipt for the deleted messages.
            deleteMessagesRequest.NotifyNonRead = 0x00;
            deleteMessagesRequest.MessageIdCount = (ushort)messageIds.Length;
            deleteMessagesRequest.MessageIds = messageIds;
            this.Adapter.DoRopCall(deleteMessagesRequest, this.RootFolderHandle, ref ropResponse, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R262 and MS-OXCFOLD_R1017.

            bool allDeleted = this.IsMessageDeleted(messageId2, this.RootFolderId) && this.IsMessageDeleted(messageId3, this.RootFolderId);

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R262");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R262.
            Site.CaptureRequirementIfIsTrue(
                allDeleted,
                262,
                "[In RopDeleteMessages ROP] The RopDeleteMessages ROP ([MS-OXCROPS] section 2.2.4.11) is used to soft delete one or more messages from a folder.");

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest
            {
                RopId = (byte)RopId.RopGetContentsTable,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                TableFlags = (byte)FolderTableFlags.SoftDeletes
            };
            RopGetContentsTableResponse getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, this.RootFolderHandle, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");

            if (Common.IsRequirementEnabled(1017, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1017");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1017
                // Method GetContentsTable succeeds and it's row count is 3, this requirement verified.
                Site.CaptureRequirementIfAreEqual<uint>(
                    3,
                    getContentsTableResponse.RowCount,
                    1017,
                    @"[In RopGetContentsTable ROP Request Buffer] If this bit [SoftDeletes] is set, the contents table lists only the messages that are soft deleted.");
            }

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessages operation performs successfully.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC05_RopHardDeleteMessagesSuccess()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Create message in the root folder.
            uint normalMessageHandle = 0;
            ulong normalMmessageId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref normalMmessageId, ref normalMessageHandle);
            #endregion

            #region Step 2. Call RopHardDeleteMessage ROP operation to hard delete the message in the root folder.

            ulong[] messageIds = new ulong[] { normalMmessageId };
            RopHardDeleteMessagesRequest hardDeleteMessagesRequest = new RopHardDeleteMessagesRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessages,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                NotifyNonRead = 0x00,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds
            };
            RopHardDeleteMessagesResponse hardDeleteMessagesResponse = this.Adapter.HardDeleteMessages(hardDeleteMessagesRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(
                0,
                hardDeleteMessagesResponse.ReturnValue,
                "If ROP succeeds, the ReturnValue of its response is 0(success)");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1153");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1153.
            // The return value in RopHardDeleteMessages is zero indicates the client get the RopHardDeleteMessages ROP response successfully, MS-OXCFOLD_R1153 can be verified directly.
            Site.CaptureRequirement(
                1153,
                @"[In Processing a RopHardDeleteMessages ROP Request] The server responds with a RopHardDeleteMessages ROP response buffer.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R99402");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R99402.
            // The RopHardDeleteMessages ROP operation performs successfully on a private mailbox, MS-OXCFOLD_R99402 can be verified directly.
            Site.CaptureRequirement(
                99402,
                @"[In RopHardDeleteMessages ROP] This ROP [RopHardDeleteMessages] applies to [both public folders and] private mailboxes.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R293");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R293.
            // The WantAsynchronous was set to zero and the server responds a RopHardDeleteMessages ROP response indicates the ROP is processed synchronously, MS-OXCFOLD_R293 can be verified directly.
            Site.CaptureRequirement(
                293,
                @"[In RopHardDeleteMessages ROP Request Buffer] WantAsynchronous (1 byte): [A Boolean value that is] zero (FALSE) if the ROP is to be processed synchronously.");

            #region Verify MS-OXCFOLD_R304 and MS-OXCFOLD_R437.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R304");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R304
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                hardDeleteMessagesResponse.PartialCompletion,
                304,
                @"[In RopHardDeleteMessages ROP Response Buffer] PartialCompletion (1 byte): Otherwise [if the ROP successes for a subset of targets], the value [of PartialCompletion field] is zero (FALSE).");
  
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R437");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R437
            // The RopHardDeleteMessages ROP operation performs successfully on a private mailbox, MS-OXCFOLD_R437 can be verified directly.
            Site.CaptureRequirement(
                437,
                @"[In Deleting the Contents of a Folder] To remove particular messages from a folder, the client sends [either a RopDeleteMessages ROP request ([MS-OXCROPS] section 2.2.4.11) or] a RopHardDeleteMessages ROP request ([MS-OXCROPS] section 2.2.4.12).");
            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveCopyMessages operation responds with the related error codes.  
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC06_RopMoveCopyMessagesFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            List<uint> handlelist;
            this.GenericFolderInitialization();

            #region Step 1. Create a message in the root folder.
            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.
            uint subfolderHandle1 = 0;
            ulong subfolderId1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 3. Call RopMoveCopyMessages using a logon object handle as a source folder handle to move the message created in step 2 to [MSOXCFOLDSubfolder1].
            ulong[] messageIds = new ulong[] { subfolderId1 };
            List<uint> handleList = new List<uint>
            {
                // Use logon object handle as a source handle for RopMoveCopyMessages in which case is purposed to get an error code ecNotSupported [0x80040102].  
                this.LogonHandle, subfolderHandle1
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
                WantCopy = 0
            };
            RopMoveCopyMessagesResponse moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handleList, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R589 and MS-OXCFOLD_R590.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R589");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R589
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                moveCopyMessagesResponse.ReturnValue,
                589,
                @"[In Processing a RopMoveCopyMessages ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R590");

            // The MS-OXCFOLD_R589 captured error code [ecNotSupported], using logon handle as a source folder handle, capture this requirement directly.
            Site.CaptureRequirement(
                590,
                @"[In Processing a RopMoveCopyMessages ROP Request] When the error code is ecNotSupported, it indicates that either the source object or the destination object is not a Folder object.");

            #endregion

            handleList.Clear();
            #endregion

            #region Step 4. Call RopMoveCopyMessages to move the message created in step 1 from the root folder to [MSOXCFOLDSubfolder1] synchronously.
            messageIds = new ulong[] { messageId };
            handlelist = new List<uint>
            {
                this.RootFolderHandle, subfolderHandle1
            };

            moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
            {
                RopId = (byte)RopId.RopMoveCopyMessages,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds,
                WantAsynchronous = 0x00,
                WantCopy = 0x00
            };

            // WantCopy is zero (FALSE) indicates this is a move operation.
            moveCopyMessagesResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(0, moveCopyMessagesResponse.PartialCompletion, "The ROP successes for all subsets of targets");
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate the RopEmptyFolder operation responds with error codes.  
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC07_RopEmptyFolderFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Calls RopEmptyfolder using a logon handle rather than a folder handle.
            RopEmptyFolderRequest emptyFolderRequest = new RopEmptyFolderRequest
            {
                RopId = (byte)RopId.RopEmptyFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };

            // Using a logon object handle to refer a logon object in which case is purposed to test error code ecNotSupported [0x80040102].  
            RopEmptyFolderResponse emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R624 and MS-OXCFOLD_R625.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R624");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R624.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                emptyFolderResponse.ReturnValue,
                624,
                @"[In Processing a RopEmptyFolder ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R625");

            // The MS-OXCFOLD_R624 captured error code [ecNotSupported], capture this requirement directly.
            Site.CaptureRequirement(
                625,
                @"[In Processing a RopEmptyFolder ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopEmptyFolder ROP] was called on is not a Folder object.");
            #endregion

            #endregion

            #region Step 2. The client calls RopCreateFolder to create the search folder [MSOXCFOLDSearchFolder1] under the root folder.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Searchfolder,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle1 = this.responseHandles[0][createFolderResponse.OutputHandleIndex];
            #endregion

            #region Step 3. The client calls RopEmptyFolder on the search folder created in step 2.
            emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, searchFolderHandle1, ref this.responseHandles);

            if (Common.IsRequirementEnabled(124901, this.Site))
            {
                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R124901");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R124901.
                // If the client attempt to empty a search folder, the server will return ecNotSupported in the ReturnValue.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    emptyFolderResponse.ReturnValue,
                    124901,
                    @"[In Appendix A: Product Behavior] Implementation does return ecNotSupported (0x80040102) in the ReturnValue field of the RopEmptyFolder ROP response buffer, if the client attempts to empty a search folder. (Exchange 2007 and above follow this behavior.)");

                // Add the debug information.
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R3012");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R3012.
                // If the client attempt to empty a search folder, the server will return ecNotSupported in the ReturnValue.
                // Because the error code 0x80040102 has been captured by R124901, so R3012 can be captured directly. 
                Site.CaptureRequirement(
                    3012,
                    @"[In Processing a RopEmptyFolder ROP Request] When the error code is ecNotSupported, it indicates that a folder that this ROP [RopEmptyFolder ROP] was called on is not allowed to be emptied.");
            }
            #endregion

            #region Step 4. The client calls RopOpenFolder to open the root folder.
            uint rootFolderHandle = 0;
            this.OpenFolder(this.LogonHandle, this.DefaultFolderIds[0], ref rootFolderHandle); 
            #endregion

            #region Step 5. The client calls RopEmptyFolder on the root folder.
            emptyFolderResponse = this.Adapter.EmptyFolder(emptyFolderRequest, rootFolderHandle, ref this.responseHandles);

            if (Common.IsRequirementEnabled(124801, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R124801");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R124801
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x80040102,
                    emptyFolderResponse.ReturnValue,
                    124801,
                    @"[In Appendix A: Product Behavior] Implementation does not return ecNotSupported when the RopEmptyFolder ROP is called on the Root folder. <18> Section 3.2.5.9: Exchange 2007 do not return ecNotSupported (0x80040102) when the RopEmptyFolder ROP ([MS-OXCROPS] section 2.2.4.9) is called on the Root folder.");
            }

            if (Common.IsRequirementEnabled(124802, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R124802");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R124802
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    emptyFolderResponse.ReturnValue,
                    124802,
                    @"[In Appendix A: Product Behavior] Implementation does return ecNotSupported when the RopEmptyFolder ROP is called on the Root folder. <18> Section 3.2.5.9: Update Rollup 4 for Exchange Server 2010 Service Pack 2 (SP2), Exchange 2013 and Exchange 2016 return ecNotSupported when the RopEmptyFolder ROP is called on the Root folder.");
            }
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopDeleteMessages operation responds with error code. 
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC08_RopDeleteMessagesFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Create a message in the root folder.
            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. Delete the message created in step 1 with logon object handle.

            RopDeleteMessagesRequest deleteMessagesRequest = new RopDeleteMessagesRequest
            {
                RopId = (byte)RopId.RopDeleteMessages,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                NotifyNonRead = 0x00,
                MessageIds = new ulong[]
                {
                    messageId
                }
            };
            deleteMessagesRequest.MessageIdCount = (ushort)deleteMessagesRequest.MessageIds.Length;

            // Use logon object handle to delete message is purposed to get an error code ecNotSupported [0x80040102].  
            RopDeleteMessagesResponse deleteMessagesResponse = this.Adapter.DeleteMessages(deleteMessagesRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R637 and MS-OXCFOLD_R638.
            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R637");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R637
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                deleteMessagesResponse.ReturnValue,
                637,
                @"[In Processing a RopDeleteMessages ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R638");

            // MS-OXCFOLD_R637 captured error code [ecNotSupported], capture this requirement directly.
            Site.CaptureRequirement(
                638,
                @"[In Processing a RopDeleteMessages ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopDeleteMessages ROP] was called on is not a Folder object.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessages operation responds with error code.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC09_RopHardDeleteMessagesFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Create a message in the root folder.
            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(this.RootFolderHandle, this.RootFolderId, ref messageId, ref messageHandle);
            #endregion

            #region Step 2. Hard delete the message created in step 1 with a logon object handle.
            ulong[] messageIds = new ulong[] { messageId };

            RopHardDeleteMessagesRequest hardDeleteMessagesRequest = new RopHardDeleteMessagesRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessages,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                NotifyNonRead = 0x00,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds
            };

            // Use logon object handle to hard-delete message is purposed to get error code ecNotSupported [0x80040102].  
            RopHardDeleteMessagesResponse hardDeleteMessagesResponse = this.Adapter.HardDeleteMessages(hardDeleteMessagesRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R644 and MS-OXCFOLD_R645.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R644");

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R644.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                hardDeleteMessagesResponse.ReturnValue,
                644,
                @"[In Processing a RopHardDeleteMessages ROP Request]The value of error code ecNotSupported is 0x80040102.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R645");

            // The MS-OXCFOLD_R644 captured error code [ecNotSupported], capture this requirement directly.
            Site.CaptureRequirement(
                645,
                @"[In Processing a RopHardDeleteMessages ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopHardDeleteMessages ROP] was called on is not a Folder object.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopHardDeleteMessagesAndSubFolders operation responds with error code.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC10_RopHardDeleteMessagesAndSubfoldersFailure()
        {
            this.CheckWhetherSupportTransport();
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1.	Call RopHardDeleteMessagesAndSubfolders using a logon handle rather than a folder handle.
            RopHardDeleteMessagesAndSubfoldersRequest hardDeleteMessagesAndSubfoldersRequest = new RopHardDeleteMessagesAndSubfoldersRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                WantAsynchronous = 0x00,
                WantDeleteAssociated = 0xFF
            };

            // Using logon object handle to hard-delete messages and subfolders is purposed to get error code ecNotSupported [0x80040102].  
            RopHardDeleteMessagesAndSubfoldersResponse hardDeleteMessagesAndSubfoldersResponse = this.Adapter.HardDeleteMessagesAndSubfolders(hardDeleteMessagesAndSubfoldersRequest, this.LogonHandle, ref this.responseHandles);

            #region Verify MS-OXCFOLD_R630 and MS-OXCFOLD_R629.

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R630");

            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                hardDeleteMessagesAndSubfoldersResponse.ReturnValue,
                630,
                "[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request] When the error code is ecNotSupported, it indicates that the object that this ROP [RopHardDeleteMessagesAndSubfolders ROP] was called on is not a Folder object.");

            // The MS-OXCFOLD_R630 captured error code [ecNotSupported] using a logon object handle, capture this requirement directly.
            Site.CaptureRequirement(
                629,
                @"[In Processing a RopHardDeleteMessagesAndSubfolders ROP Request]The value of error code ecNotSupported is 0x80040102.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case is designed to validate that the RopMoveCopyMessages operation copy or move a message from search folder.
        /// </summary>
        [TestCategory("MSOXCFOLD"), TestMethod()]
        public void MSOXCFOLD_S02_TC11_RopMoveCopyMessagesUseSearchFolderSuccess()
        {
            this.Adapter.DoConnect(ConnectionType.PrivateMailboxServer);
            this.GenericFolderInitialization();

            #region Step 1. Call RopCreateFolder to create [MSOXCFOLDSubfolder1] under the root folder.
            ulong subfolderId1 = 0;
            uint subfolderHandle1 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder1, ref subfolderId1, ref subfolderHandle1);
            #endregion

            #region Step 2. Create a message in the [MSOXCFOLDSubfolder1] folder created in step 1.

            uint messageHandle = 0;
            ulong messageId = 0;
            this.CreateSaveMessage(subfolderHandle1, subfolderId1, ref messageId, ref messageHandle);

            #endregion

            #region Step 3. Call RopCreateFolder to create [MSOXCFOLDSubfolder2] under the root folder.
            ulong subfolderId2 = 0;
            uint subfolderHandle2 = 0;
            this.CreateFolder(this.RootFolderHandle, Constants.Subfolder2, ref subfolderId2, ref subfolderHandle2);
            #endregion

            #region Step 4. Create a message in the [MSOXCFOLDSubfolder2] folder created in step 3.

            uint messageHandleInSubFolder2 = 0;
            ulong messageIdInSubFolder2 = 0;
            this.CreateSaveMessage(subfolderHandle2, subfolderId2, ref messageIdInSubFolder2, ref messageHandleInSubFolder2);
            #endregion

            #region Step 4. Call RopCreateFolder to create a search folder [MSOXCFOLDSearchFolder] under the root folder.
            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                OutputHandleIndex = Constants.CommonOutputHandleIndex,
                FolderType = (byte)FolderType.Searchfolder,
                UseUnicodeStrings = 0x0,
                OpenExisting = 0x00,
                Reserved = 0x0,
                DisplayName = Encoding.ASCII.GetBytes(Constants.SearchFolder),
                Comment = Encoding.ASCII.GetBytes(Constants.SearchFolder)
            };
            RopCreateFolderResponse createFolderResponse = this.Adapter.CreateFolder(createFolderRequest, this.RootFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, createFolderResponse.ReturnValue, "RopCreateFolder ROP operation performs successfully!");
            uint searchFolderHandle = this.responseHandles[0][createFolderResponse.OutputHandleIndex];

            #endregion

            #region Step 5. The client calls RopSetSearchCriteria to establish search criteria for [MSOXCFOLDSubFolder1].

            RopSetSearchCriteriaRequest setSearchCriteriaRequest = new RopSetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopSetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex
            };
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = (ushort)MessagePropertyId.PidTagMessageClass,
                PropertyType = (ushort)PropertyType.PtypString
            };
            ExistRestriction existRestriction = new ExistRestriction
            {
                PropTag = propertyTag
            };
            setSearchCriteriaRequest.RestrictionDataSize = (ushort)existRestriction.Size();
            setSearchCriteriaRequest.RestrictionData = existRestriction.Serialize();
            setSearchCriteriaRequest.FolderIds = new ulong[] { subfolderId1 };
            setSearchCriteriaRequest.FolderIdCount = (ushort)setSearchCriteriaRequest.FolderIds.Length;
            setSearchCriteriaRequest.SearchFlags = (uint)SetSearchFlags.ContentIndexedSearch | (uint)SetSearchFlags.StaticSearch | (uint)SetSearchFlags.RestartSearch;
            RopSetSearchCriteriaResponse setSearchCriteriaResponse = this.Adapter.SetSearchCriteria(setSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, setSearchCriteriaResponse.ReturnValue, "RopSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 6. The client calls RopGetSearchCriteria to obtain the search criteria and the status of the search folder [MSOXCFOLDSubFolder1].
            RopGetSearchCriteriaRequest getSearchCriteriaRequest = new RopGetSearchCriteriaRequest
            {
                RopId = (byte)RopId.RopGetSearchCriteria,
                LogonId = Constants.CommonLogonId,
                InputHandleIndex = Constants.CommonInputHandleIndex,
                UseUnicode = 0x00,
                IncludeRestriction = 0x01,
                IncludeFolders = 0x01
            };
            RopGetSearchCriteriaResponse getSearchCriteriaResponse = this.Adapter.GetSearchCriteria(getSearchCriteriaRequest, searchFolderHandle, ref this.responseHandles);
            Site.Assert.AreEqual<uint>(Constants.SuccessCode, getSearchCriteriaResponse.ReturnValue, "RopGetSearchCriteria ROP operation performs successfully!");

            #endregion

            #region Step 7. The client calls RopGetContentsTable to get handle of the contents table in the search folder [MSOXCFOLDSearchFolder].

            RopGetContentsTableRequest getContentsTableRequest = new RopGetContentsTableRequest();
            RopGetContentsTableResponse getContentsTableResponse;
            getContentsTableRequest.RopId = (byte)RopId.RopGetContentsTable;
            getContentsTableRequest.LogonId = Constants.CommonLogonId;
            getContentsTableRequest.InputHandleIndex = Constants.CommonInputHandleIndex;
            getContentsTableRequest.OutputHandleIndex = Constants.CommonOutputHandleIndex;
            getContentsTableRequest.TableFlags = (byte)FolderTableFlags.None;

            int count = 0;
            uint tableHandle;
            do
            {
                getContentsTableResponse = this.Adapter.GetContentsTable(getContentsTableRequest, searchFolderHandle, ref this.responseHandles);
                Site.Assert.AreEqual<uint>(Constants.SuccessCode, getContentsTableResponse.ReturnValue, "RopGetContentsTable ROP operation performs successfully!");
                tableHandle = this.responseHandles[0][getContentsTableResponse.OutputHandleIndex];
                if (getContentsTableResponse.RowCount != 1)
                {
                    Thread.Sleep(this.WaitTime);
                }
                else
                {
                    break;
                }

                count++;
            }
            while (count < this.RetryCount);

            #endregion

            #region Step 8. Sets the properties PidTagMid visible on the content table.

            RopSetColumnsRequest setColumnsRequest;
            PropertyTag[] propertyTags = new PropertyTag[1];
            propertyTags[0].PropertyId = (ushort)MessagePropertyId.PidTagMid;
            propertyTags[0].PropertyType = (ushort)PropertyType.PtypInteger64;

            setColumnsRequest.RopId = 0x12;
            setColumnsRequest.LogonId = 0x00;
            setColumnsRequest.InputHandleIndex = 0x00;
            setColumnsRequest.PropertyTagCount = (ushort)propertyTags.Length;
            setColumnsRequest.PropertyTags = propertyTags;
            setColumnsRequest.SetColumnsFlags = 0x00; // Sync

            object ropResponse = new object();
            this.Adapter.DoRopCall(setColumnsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            #endregion

            #region Step 9. Gets the message ID in the search folder [MSOXCFOLDSearchFolder].

            RopQueryRowsRequest queryRowsRequest;
            RopQueryRowsResponse queryRowsResponse;

            queryRowsRequest.RopId = 0x15;
            queryRowsRequest.LogonId = 0x00;
            queryRowsRequest.InputHandleIndex = 0x00;
            queryRowsRequest.QueryRowsFlags = 0x00;
            queryRowsRequest.ForwardRead = 0x01;
            queryRowsRequest.RowCount = 1;

            this.Adapter.DoRopCall(queryRowsRequest, tableHandle, ref ropResponse, ref this.responseHandles);
            queryRowsResponse = (RopQueryRowsResponse)ropResponse;
            ulong messageID = BitConverter.ToUInt64(queryRowsResponse.RowData.PropertyRows[0].PropertyValues[0].Value, 0);
            #endregion

            #region Step 10. Call RopMoveCopyMessages to copy the message in the [MSOXCFOLDSearchFolder] to root folder synchronously.
            ulong[] messageIds = new ulong[1];
            messageIds[0] = messageID;
            List<uint> handlelist = new List<uint>
            {
                searchFolderHandle, this.RootFolderHandle
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
            RopMoveCopyMessagesResponse copyMessagesFromSearchFolderResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, copyMessagesFromSearchFolderResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(0, copyMessagesFromSearchFolderResponse.PartialCompletion, "The ROP successes for all subsets of targets");
            handlelist.Clear();
            #endregion

            #region Step 11. Validate the message is copied successfully in above step 9.
            uint rootFolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, this.RootFolderHandle);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R121702.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R121702: expected contents count of the root folder is {0}, actual contents count of the root folder is {1};", 1, rootFolderContentsCountActual);

            // The source folder is a search folder, so if the message was copied successfully from MSOXCFOLDSearchFolder to root folder, R121702 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                1,
                rootFolderContentsCountActual,
                121702,
                @"[In RopMoveCopyMessages ROP Request Buffer] SourceHandleIndex (1 byte): [The source Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder from which the messages will be copied] This folder can be a search folder.");
            #endregion

            #region Step 12. Call RopMoveCopyMessages to move the message in the [MSOXCFOLDSearchFolder] to root folder synchronously.
            messageIds = new ulong[1];
            messageIds[0] = messageID;
            handlelist = new List<uint>
            {
                searchFolderHandle, this.RootFolderHandle
            };

            moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
            {
                RopId = (byte)RopId.RopMoveCopyMessages,
                LogonId = Constants.CommonLogonId,
                SourceHandleIndex = 0x00,
                DestHandleIndex = 0x01,
                MessageIdCount = (ushort)messageIds.Length,
                MessageIds = messageIds,
                WantAsynchronous = 0x00,
                WantCopy = 0x00
            };

            // WantCopy is zero (FALSE) indicates this is a move operation.
            RopMoveCopyMessagesResponse moveMessagesFromSearchFolderResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);

            Site.Assert.AreEqual<uint>(0, moveMessagesFromSearchFolderResponse.ReturnValue, "The RopMoveCopyMessages ROP operation performs successfully.");
            Site.Assert.AreEqual<uint>(0, moveMessagesFromSearchFolderResponse.PartialCompletion, "The ROP successes for all subsets of targets");
            handlelist.Clear();
            #endregion

            #region Step 13. Validate the message is moved successfully in above step 12.
            rootFolderContentsCountActual = this.GetContentsTable(FolderTableFlags.None, this.RootFolderHandle);

            // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R121701.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R121701: expected contents count of the root folder is {0}, actual contents count of the root folder is {1};", 2, rootFolderContentsCountActual);

            // The source folder is a search folder, so if the message was moved successfully from MSOXCFOLDSearchFolder to root folder, R121701 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                2,
                rootFolderContentsCountActual,
                121701,
                @"[In RopMoveCopyMessages ROP Request Buffer] SourceHandleIndex (1 byte): [The source Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder from which the messages will be moved] This folder can be a search folder.");
            #endregion

            #region Step 14. Call RopMoveCopyMessages to copy the message in the [MSOXCFOLDSubfolder2] to the [MSOXCFOLDSearchFolder] synchronously.
            RopMoveCopyMessagesResponse copyMessagesToSearchFolderResponse = new RopMoveCopyMessagesResponse();
            if (Common.IsRequirementEnabled(1246, this.Site))
            {
                messageIds[0] = messageIdInSubFolder2;
                handlelist = new List<uint>
                {
                subfolderHandle2, searchFolderHandle
                };
                moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
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
                copyMessagesToSearchFolderResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R121802: the PartialCompletion in the MoveCopyMessages response is {0}", copyMessagesToSearchFolderResponse.PartialCompletion);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R121802
                // If the ROP RopMoveCopyMessages fails to copy message, the value of PartialCompletion field is nonzero.
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x00000000,
                    copyMessagesToSearchFolderResponse.ReturnValue,
                    121802,
                    @"[In RopMoveCopyMessages ROP Request Buffer] DestHandleIndex (1 byte): [The destination Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder to which the messages will be copied.] This folder cannot be a search folder.");
            }
            #endregion

            #region Step 15. Call RopMoveCopyMessages to move the message in the [MSOXCFOLDSubfolder2] to the [MSOXCFOLDSearchFolder] synchronously.
            RopMoveCopyMessagesResponse moveMessagesToSearchFolderResponse;
            if (Common.IsRequirementEnabled(1246, this.Site))
            {
                messageIds[0] = messageIdInSubFolder2;
                handlelist = new List<uint>
                {
                subfolderHandle2, searchFolderHandle
                };
                moveCopyMessagesRequest = new RopMoveCopyMessagesRequest
                {
                    RopId = (byte)RopId.RopMoveCopyMessages,
                    LogonId = Constants.CommonLogonId,
                    SourceHandleIndex = 0x00,
                    DestHandleIndex = 0x01,
                    MessageIdCount = (ushort)messageIds.Length,
                    MessageIds = messageIds,
                    WantAsynchronous = 0x00,
                    WantCopy = 0x00
                };

                // WantCopy is zero (FLASE) indicates this is a move operation.
                 moveMessagesToSearchFolderResponse = this.Adapter.MoveCopyMessages(moveCopyMessagesRequest, handlelist, ref this.responseHandles);

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R121801: the PartialCompletion in the MoveCopyMessages response is {0}", moveMessagesToSearchFolderResponse.PartialCompletion);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R121801
                // If the ROP RopMoveCopyMessages fails to move message, the value of PartialCompletion field is nonzero.
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0x00000000,
                    moveMessagesToSearchFolderResponse.ReturnValue,
                    121801,
                    @"[In RopMoveCopyMessages ROP Request Buffer] DestHandleIndex (1 byte): [The destination Server object for this operation [RopMoveCopyMessages ROP] is a Folder object that represents the folder to which the messages will be moved.] This folder cannot be a search folder.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1216");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1216
                bool isVerifiedR1216 = copyMessagesFromSearchFolderResponse.ReturnValue == 0x00000000 &&
                                    moveMessagesFromSearchFolderResponse.ReturnValue == 0x00000000 &&
                                    copyMessagesToSearchFolderResponse.ReturnValue != 0x00000000 &&
                                    moveMessagesToSearchFolderResponse.ReturnValue != 0x00000000;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR1216,
                    1216,
                    @"[In RopMoveCopyMessages ROP] The source folder can be a search folder, but the destination folder cannot.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1246: the return value is {0} when copy a message to a search folder and the return value is {1} when move a message to a search folder.", copyMessagesToSearchFolderResponse.ReturnValue, moveMessagesToSearchFolderResponse.ReturnValue);

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1246
                bool isVerifiedR1246 = copyMessagesToSearchFolderResponse.ReturnValue == 0x00000460 && moveMessagesToSearchFolderResponse.ReturnValue == 0x00000460;

                this.Site.CaptureRequirementIfIsTrue(
                    isVerifiedR1246,
                    1246,
                    @"[In Processing a RopMoveCopyMessages ROP Request] When the error code is ecSearchFolder, it indicates the destination object is a search folder. ");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCFOLD_R1245");

                // Verify MS-OXCFOLD requirement: MS-OXCFOLD_R1245
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000460,
                    copyMessagesToSearchFolderResponse.ReturnValue,
                    1245,
                    @"[In Processing a RopMoveCopyMessages ROP Request] The value of error code ecSearchFolder is 0x00000460.");
            }
            #endregion
        }
    }
}