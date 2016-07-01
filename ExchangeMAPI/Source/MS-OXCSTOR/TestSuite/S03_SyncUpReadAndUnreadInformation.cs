namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is to test the ROP related to sync up the per-user information
    /// </summary>
    [TestClass]
    public class S03_SyncUpReadAndUnreadInformation : TestSuiteBase
    {
        #region Test Class Initialization
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
        /// </summary>
        /// <param name="context">Test context value</param>
        [ClassInitialize]
        public static void ClassInitialize(TestContext context)
        {
            TestClassBase.Initialize(context);
        }

        /// <summary>
        ///  Use ClassCleanup to run code after all tests in a class have run.
        /// </summary>
        [ClassCleanup]
        public static void ClassCleanup()
        {
            TestClassBase.Cleanup();
        }
        #endregion

        /// <summary>
        /// This test case verifies the requirements related to RopReadPerUserInformation ROP in public folder logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S03_TC01_ReadPerUserInformationForPublicLogon()
        {
            this.CheckTransportIsSupported();

            #region Step 1: Connect to the server.
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 2: Logon to public folder.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region step 3: Open a public folder
            RopOpenFolderRequest openFolderRequest;
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = ConstValues.LoginId;
            openFolderRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            openFolderRequest.OutputHandleIndex = ConstValues.OutputHandleIndex;
            openFolderRequest.FolderId = this.logonResponse.FolderIds[1];
            openFolderRequest.OpenModeFlags = 0x0;

            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

            RopOpenFolderResponse openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x0,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            uint openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];
            #endregion

            #region step 4: Create a public folder
            string publicFolderName = Common.GenerateResourceName(Site, "PublicFolder") + "\0";
            RopCreateFolderRequest createFolderRequest;
            createFolderRequest.RopId = 0x1C;
            createFolderRequest.LogonId = ConstValues.LoginId;
            createFolderRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            createFolderRequest.OutputHandleIndex = ConstValues.OutputHandleIndex;
            createFolderRequest.FolderType = 0x01;
            createFolderRequest.UseUnicodeStrings = 0x0;
            createFolderRequest.OpenExisting = 0xFF;
            createFolderRequest.Reserved = 0x0;
            createFolderRequest.DisplayName = System.Text.Encoding.ASCII.GetBytes(publicFolderName);
            createFolderRequest.Comment = System.Text.Encoding.ASCII.GetBytes(publicFolderName);
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            RopCreateFolderResponse createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x0,
                createFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            ulong folderId = createFolderResponse.FolderId;
            uint folderHandle = this.outputBuffer.ServerObjectHandleTable[createFolderRequest.OutputHandleIndex];

            #endregion

            #region Step 5: Get LongTermID of the folder created in step 4.
            LongTermId longTermId = this.GetLongTermIdFromId(folderId);
            #endregion

            #region Step 6: Call RopReadPerUserInformation ROP request to check if user information exists.
            this.readPerUserInformationRequest.FolderId = longTermId;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1257");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1257.
            Site.CaptureRequirementIfIsNull(
                this.readPerUserInformationResponse.Data,
                1257,
                @"[In Public Folders Specific Behavior] If the row does not exist, then the server returns an empty array in the Data field of the response.");
            #endregion

            #region Step 7: Create a message in the public folder created in step 4.
            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest();
            createMessageRequest.RopId = (byte)RopId.RopCreateMessage;
            createMessageRequest.LogonId = ConstValues.LoginId;
            createMessageRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            createMessageRequest.OutputHandleIndex = ConstValues.OutputHandleIndex;
            createMessageRequest.CodePageId = 0x0FFF;
            createMessageRequest.FolderId = folderId;
            createMessageRequest.AssociatedFlag = 0x0;
            this.oxcstorAdapter.DoRopCall(createMessageRequest, folderHandle, ROPCommandType.Others, out this.outputBuffer);

            RopCreateMessageResponse createMessageResponse = (RopCreateMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "Creating Message should succeed");
            uint messageHandle = this.outputBuffer.ServerObjectHandleTable[createMessageRequest.OutputHandleIndex];

            // Save a Message.
            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest();
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = ConstValues.LoginId;
            saveChangesMessageRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            saveChangesMessageRequest.ResponseHandleIndex = ConstValues.OutputHandleIndex;
            saveChangesMessageRequest.SaveFlags = 0x0C;
            this.oxcstorAdapter.DoRopCall(saveChangesMessageRequest, messageHandle, ROPCommandType.Others, out this.outputBuffer);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                saveChangesMessageResponse.ReturnValue,
                "Save Messages Success.");
            ulong messageId = saveChangesMessageResponse.MessageId;
            #endregion

            #region Step 8: Call RopReadPerUserInformation ROP request to get data.
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            Site.Assert.IsNotNull(this.readPerUserInformationResponse.Data, "Data should be exist if user reads mail in public folder.");
            byte[] data = this.readPerUserInformationResponse.Data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R965");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R965.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                965,
                @"[In Receiving a RopReadPerUserInformation ROP Request] This operation [RopReadPerUserInformation] can be issued against either a private mailbox logon or a public folders logon.");
            #endregion

            #region Step 9: Call RopReadPerUserInformation MaxDataSize set to zero.
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1005");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1005
            // MaxDataSize be set to 0 server will use the default value 4096
            int adjMaxDataSize = 4096;
            int blobMinDataOffset = data.Length - (int)this.readPerUserInformationRequest.DataOffset;
            bool isVerify_R1005 = (blobMinDataOffset > adjMaxDataSize)
                ? (this.readPerUserInformationResponse.DataSize == adjMaxDataSize)
                : (this.readPerUserInformationResponse.DataSize == blobMinDataOffset);
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1005,
                1005,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server MUST set DataSize to the lesser of the following two values [the adjusted value of MaxDataSize, the entire BLOB minus the value of DataOffset.].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1000");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1000.
            // The DataSize field in the RopReadPerUserInformation is the lesser of the following two values [the adjusted value of MaxDataSize, the entire BLOB minus the value of DataOffset.], it indicates the server compares the adjusted value of MaxDataSize to the size of the remaining BLOB segment.
            // MS-OXCSTOR_R1000 can be verified directly.
            Site.CaptureRequirement(
                1000,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] 	The server compares the adjusted value of MaxDataSize to the size of the remaining BLOB segment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R517");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R517.
            Site.CaptureRequirementIfIsNotNull(
                this.readPerUserInformationResponse.Data,
                517,
                @"[In RopReadPerUserInformation ROP] When this ROP is issued against a public folders logon, the current per-user read/unread data for the public folder is retrieved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R971");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R971.
            // MS-OXCSTOR_R517 has verified the server returned current per-user read/unread data for the public folder successfully, MS-OXCSTOR_R971 can be verified directly.
            Site.CaptureRequirement(
                971,
                @"[In Public Folders Specific Behavior] The server searches the per-user data table for the only row with an FID equal to the value of the FolderId field and the user ID equal to the logged on user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R975");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R975
            // Verify the data is a BLOB which is formatted as a serialized IDSET with REPLGUID
            bool isVerify_R975 = this.VerifyDataIsIDSETStructure(this.readPerUserInformationResponse.Data);

            Site.CaptureRequirementIfIsTrue(
                isVerify_R975,
                975,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] [The change number set is serialized into a binary large object (BLOB) that is formatted as a serialized IDSET with REPLGUID structure, as specified in [MS-OXCFXICS] section 2.2.2.4.2.] The server then returns the BLOB in the Data field of the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R974");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R974.
            // MS-OXCSTOR_R975 has verified the change number set structure, MS-OXCSTOR_R545 can be verified directly.
            Site.CaptureRequirement(
                974,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The change number set MUST be serialized into a binary large object (BLOB) that is formatted as a serialized IDSET with REPLGUID structure, as specified in [MS-OXCFXICS] section 2.2.2.4.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1003");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1003
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)this.readPerUserInformationResponse.DataSize,
                (ushort)this.readPerUserInformationResponse.Data.Length,
                1003,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The DataSize field specifies the actual number of bytes that are returned in the response.");
            #endregion

            #region Step 10: RopReadPerUserInformation with DataOffset less than zero.
            this.readPerUserInformationRequest.DataOffset = 0xffffffff;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            #region Capture

            if (Common.IsRequirementEnabled(1042, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1042");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1042.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    this.readPerUserInformationResponse.ReturnValue,
                    1042,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The implementation does fail the operation with 0x80004005 (ecError) in the ReturnValue field, if the value of the DataOffset field is less than zero. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1007, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R554");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R554.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.readPerUserInformationResponse.ReturnValue,
                    554,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecRpcFormat: Its value is 0x000004B6.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R555");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R555.
                // MS-OXCSTOR_R554 was captured by calling RopReadPerUserInformation with DataOffset value less than zero, MS-OXCSTOR_R555 can be verified directly.
                Site.CaptureRequirement(
                    555,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecRpcFormat: The DataOffset value was less than zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1007");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1007.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.readPerUserInformationResponse.ReturnValue,
                    1007,
                    @"[In Appendix A: Product Behavior] The implementation does fail the operation with 0x000004B6 (ecRpcFormat) in the ReturnValue field, if the value of the DataOffset field is less than zero. (<46> Section 3.2.5.12.1: Exchange 2003, Exchange 2007, and Exchange 2010 fail the operation with 0x000004B6 (ecRpcFormat).)");
            }
            #endregion
            #endregion

            #region Step 11: The first call of RopReadPerUserInformation with MaxDataSize field set to 1.

            // In the response hasFinished should be false, only the first byte should be returned.
            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = 1;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            RopReadPerUserInformationResponse perUserInformationPart1 = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R552");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R552.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                perUserInformationPart1.ReturnValue,
                552,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R553");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R553.
            // The RopReadPerUserInformation ROP performs successfully and the ecNone error code was captured by MS-OXCSTOR_R552, MS-OXCSTOR_R553 can be verified directly.
            Site.CaptureRequirement(
                553,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1284");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1284.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                perUserInformationPart1.HasFinished,
                1284,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: The value of this field is FALSE if the last block of data is not being returned.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCSTOR_R543,
                the value of the MaxDataSize field of the RopReadPerUserInformation ROP request is {0},
                the value of the DataSize field of the RopReadPerUserInformation ROP response is {1}.",
                this.readPerUserInformationRequest.MaxDataSize,
                perUserInformationPart1.DataSize);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R543.
            bool isVerifiedR543 = perUserInformationPart1.DataSize <= this.readPerUserInformationRequest.MaxDataSize;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR543,
                543,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] DataSize: The value of this field MUST be less than or equal to the value of the MaxDataSize field of the request.");
            #endregion

            #region Step 12: The second call of RopReadPerUserInformation with DataOffset field set to "1" and MaxDataSize field set to 0.

            // The remained bytes should be returned, and HasFinished should be true.
            this.readPerUserInformationRequest.DataOffset = 1;
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            RopReadPerUserInformationResponse perUserInformationPart2 = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                perUserInformationPart2.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1002");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1002
            // At step8, read again for the remaining Blob segment
            // verify if the size of returned data is equal to dataSizeWritten(entire blob size) minus the value of DataOffset
            bool isVerify_R1002 = perUserInformationPart2.Data.Length == data.Length - this.readPerUserInformationRequest.DataOffset;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1002,
                1002,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The size of the remaining BLOB segment is equal to the size of the entire BLOB minus the value of DataOffset.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1009");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1009.
            // At step8, read again for the remaining Blob segment, 
            // verify if the size of returned data is equal to dataSizeWritten(entire blob size) minus the value of DataOffset.
            bool isVerify_R1009 = perUserInformationPart2.Data.Length == data.Length - this.readPerUserInformationRequest.DataOffset;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1009,
                1009,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] This [The size of the remaining BLOB segment] is equal to the size of the entire BLOB minus the value of DataOffset.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1010");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1010
            // Verify "HasFinshed" field is true "0x01" when the dataSize plus data offset equals the size of entire BloB
            Site.CaptureRequirementIfAreEqual<byte>(
                0x01,
                perUserInformationPart2.HasFinished,
                1010,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server MUST set HasFinished to TRUE if DataOffset plus DataSize equals the size of the entire BLOB.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1011");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1011
            // "In other words" means a different statement, but for same verification logic with R1010
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0x00,
                perUserInformationPart2.HasFinished,
                1011,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In other words, when the server sends the last segment of the BLOB, HasFinished MUST be set to TRUE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1008");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1008
            // Step8 is to read the remaining data after the first read at step7.
            // If the data size read at step8 is the remaining BLOB segment size, then could verify R1008
            bool isVerifiedR1008 = perUserInformationPart2.DataSize == perUserInformationPart1.DataSize + perUserInformationPart2.DataSize - this.readPerUserInformationRequest.DataOffset;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1008,
                1008,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] This [The size of the remaining BLOB segment] is the size of the portion of the BLOB that remains to be sent to the client.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R540");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R540.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                perUserInformationPart2.HasFinished,
                540,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: The value of this field is TRUE if the last block of data is being returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R539");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R539.
            // MS-OXCSTOR_R540 and MS-OXCSTOR_R1284 was verified, MS-OXCSTOR_R539 can be verified directly.
            Site.CaptureRequirement(
                539,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: Indicates whether the last block of data is being returned.");
            #endregion

            #region Step 13: RopReadPerUserInformation, read once with the DataOffset field is greater than the size of the next BLOB segment to be returned.

            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = (ushort)(data.Length - 1);
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Read once with MaxDataSize = entireBlobSize -1, the returned value of hasFinished in the response will be false,
            // all the data except the last byte will be returned, remaining data in the server will be 1
            this.readPerUserInformationRequest.DataOffset = (ushort)(data.Length + 1);
            this.readPerUserInformationRequest.MaxDataSize = (ushort)(data.Length - 1);
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1260");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1260
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.readPerUserInformationResponse.ReturnValue,
                1260,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If the value of the DataOffset field is greater than the size of the next BLOB segment to be returned, the server MUST fail the operation with 0x80004005 (ecError) in the ReturnValue field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R556");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R556
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.readPerUserInformationResponse.ReturnValue,
                556,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: Its value is 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R557");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R557.
            // MS-OXCSTOR_R556 has captured error code 0x80004005, MS-OXCSTOR_R557 can be verified directly.
            Site.CaptureRequirement(
                557,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The DataOffset value was greater than the data size.");
            #endregion

            #region Step 14: Disconnect and re-connect.
            this.oxcstorAdapter.DisconnectEx();
            this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);

            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 15: Open the folder and get folder handle.
            // Root public folder
            openFolderRequest.FolderId = this.logonResponse.FolderIds[1];
            openFolderRequest.OpenModeFlags = 0x0;
            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x0,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];

            // Created public folder
            openFolderRequest.FolderId = folderId;
            openFolderRequest.OpenModeFlags = 0x0;
            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x0,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            folderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];
            #endregion

            #region Step 16: Create a message in the folder again.
            this.oxcstorAdapter.DoRopCall(createMessageRequest, folderHandle, ROPCommandType.Others, out this.outputBuffer);

            createMessageResponse = (RopCreateMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, createMessageResponse.ReturnValue, "Creating Message should succeed");
            uint messageHandle2 = this.outputBuffer.ServerObjectHandleTable[createMessageRequest.OutputHandleIndex];

            // Save a Message.
            saveChangesMessageRequest = new RopSaveChangesMessageRequest();
            saveChangesMessageRequest.RopId = (byte)RopId.RopSaveChangesMessage;
            saveChangesMessageRequest.LogonId = ConstValues.LoginId;
            saveChangesMessageRequest.InputHandleIndex = ConstValues.InputHandleIndex;
            saveChangesMessageRequest.ResponseHandleIndex = ConstValues.OutputHandleIndex;
            saveChangesMessageRequest.SaveFlags = 0x0C;
            this.oxcstorAdapter.DoRopCall(saveChangesMessageRequest, messageHandle2, ROPCommandType.Others, out this.outputBuffer);
            saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                saveChangesMessageResponse.ReturnValue,
                "Save Messages Success.");
            ulong messageId2 = saveChangesMessageResponse.MessageId;
            #endregion

            #region Step 17: Call RopReadPerUserInformation ROP request again.
            this.readPerUserInformationRequest.FolderId = longTermId;
            this.readPerUserInformationRequest.MaxDataSize = 0x100;
            this.readPerUserInformationRequest.DataOffset = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            Site.Assert.IsNotNull(this.readPerUserInformationResponse.Data, "Data should be exist if user marks mail as un-read in public folder.");
            bool isChanged = !this.ByteArrayEquals(this.readPerUserInformationResponse.Data, data);

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1027");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1027.
            Site.CaptureRequirementIfIsTrue(
                isChanged,
                1027,
                @"[In Public Folders Specific Behavior] If the row exists, the accumulated change number information MUST replace any existing values in the table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1026");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1026.
            Site.CaptureRequirementIfIsTrue(
                isChanged,
                1026,
                @"[In Public Folders Specific Behavior] The server searches the per-user data table for the only row with a user ID equal to user ID associated with the session logon and an FID equal to a value of the FolderId field.");
            #endregion

            #region Step 17: Delete the folder created in step 4.
            RopDeleteFolderRequest deleteFolderRequest;
            deleteFolderRequest.RopId = 0x1D;
            deleteFolderRequest.LogonId = 0x00;
            deleteFolderRequest.InputHandleIndex = 0x00;

            // The folder and all of the Message objects in the folder are deleted.
            deleteFolderRequest.DeleteFolderFlags = 0x01;

            // Folder to be deleted
            deleteFolderRequest.FolderId = folderId;
            this.oxcstorAdapter.DoRopCall(deleteFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                deleteFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion
        }

        /// <summary>
        /// This test case verifies the RopWritePerUserInformation ROP related requirements in public folder logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S03_TC02_WritePerUserInformationForPublicLogon()
        {
            this.CheckTransportIsSupported();

            #region Variable
            IDSETWithReplGuid validIdset = this.GenerateRandomValidIdset();
            LongTermId longTermIdForValidData;
            LongTermId longTermIdForInValidData;
            #endregion

            #region Step 1: Connect to the server.
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 2: Logon to public folder.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion this.logonRequestForPrivateMailBox

            #region Step 3: Get the LongTermID of the public folders for the following RopReadPerUserInformation and RopWritePerUserInformation request.

            // Get LongTermID of the folder[1], used in below RopRead/WritePerUserInformation request with valid data
            longTermIdForValidData = this.GetLongTermIdFromId(this.logonResponse.FolderIds[1]);

            validIdset.ReplGuid = longTermIdForValidData.DatabaseGuid;

            // Get LongTermID of the folder[2], used in below RopRead/WritePerUserInformation request with invalid data
            longTermIdForInValidData = this.GetLongTermIdFromId(this.logonResponse.FolderIds[2]);

            #endregion

            #region Step 4: Call RopReadPerUserInformation ROP request to check current user information existed in server.
            this.readPerUserInformationRequest.FolderId = longTermIdForValidData;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #endregion

            #region Step 5: Call RopWritePerUserInformation ROP request to write subpart user information (all the data of user information except the last byte).
            byte[] dataForWrite = validIdset.Serialize();
            byte[] subpartData = new byte[dataForWrite.Length - 1];
            Array.Copy(dataForWrite, subpartData, dataForWrite.Length - 1);
            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = subpartData;
            this.writePerUserInformationRequest.DataSize = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.HasFinished = 0x00;
            byte[] dataHasBeenWritten = new byte[dataForWrite.Length - 1];
            Array.Copy(dataForWrite, dataHasBeenWritten, dataHasBeenWritten.Length);

            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R597");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R597.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                597,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R598");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R598.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                598,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Success.");


            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1012");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1012.
            // If the RopWritePerUserInformation performs successfully on a public folder logon, MS-OXCSTOR_R1012 can be verified partially.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                1012,
                @"[In Receiving a RopWritePerUserInformation ROP Request] This operation [RopWritePerUserInformation] can be issued against either a private mailbox logon or a public folders logon.");
            #endregion

            #region Step 6: Call RopWritePerUserInformation ROP request to write the last byte of the user information with wrong folderId.

            byte[] lastByteForWrite = new byte[1];
            this.writePerUserInformationRequest.FolderId = longTermIdForInValidData;
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R601");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R601.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                601,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The FolderId didn't match the value on the previous call, AND THEN DataOffset wasn't zero.");
            #endregion

            #region Step 7: Call RopWritePerUserInformation ROP request to write the last byte of the user information with wrong DataOffset.

            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)dataForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = (uint)(lastByteForWrite.Length + 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1019");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1019
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                1019,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In addition, if the current invocation's DataOffset isn't zero, the ROP [RopWritePerUserInformation] MUST fail with a ReturnValue of 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R599");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R599
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                599,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: Its value is 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R600");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R600.
            // MS-OXCSTOR_R599 has captured error code 0x80004005 and the DataOffset didn't match the size of the data written so far, MS-OXCSTOR_R600 can be verified directly.
            Site.CaptureRequirement(
                600,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The DataOffset didn't match the size of the data written so far, AND THEN DataOffset wasn't zero.");
            #endregion
            #endregion

            #region Step 8: Write the last byte of the user information with the proper offset, the ROP expect to be failed because the server assume the previous operation was aborted.
            lastByteForWrite = new byte[1];
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = lastByteForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1033");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1033.
            // The previous operation was aborted
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                1033,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If the DataOffset value does not equal the amount of data already written, the server MUST assume the previous operation was aborted.");
            #endregion

            #region Step 9: Rewrite sub-part user information again (all the data of user information except the last byte).
            dataForWrite = validIdset.Serialize();
            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.HasFinished = 0x00;
            dataHasBeenWritten = new byte[dataForWrite.Length - 1];
            Array.Copy(dataForWrite, dataHasBeenWritten, dataHasBeenWritten.Length);

            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.writePerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 10: Write the last byte of the date which was missing in last step.
            lastByteForWrite = new byte[1];
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = lastByteForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = null;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1016.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                1016,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server determines whether the current invocation is a continuation of a previous invocation by examining the FolderId and DataOffset fields.");
            #endregion
        }

        /// <summary>
        /// This test case tests sync-up read/unread information against a private mail box logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S03_TC03_ReadPerUserInformationForPrivateLogon()
        {
            this.CheckTransportIsSupported();

            #region Variable
            ushort defaultDataSize = 4096;

            LongTermId longTermIdForValidData;
            int entireBlobSize = 0;
            #endregion

            #region Step 1: Connect to the server.
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 2: Logon to a public folder to get the LongTermID of the public folder [Interpersonal messages subtree].
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            // Get LongTermID of the folder "Interpersonal messages subtree", used in below RopRead/WritePerUserInformation request with valid data
            longTermIdForValidData = this.GetLongTermIdFromId(this.logonResponse.FolderIds[1]);
            IDSETWithReplGuid validIdset1 = this.GenerateRandomValidIdset();
            IDSETWithReplGuid validIdset2 = this.GenerateRandomValidIdset();
            validIdset1.ReplGuid = longTermIdForValidData.DatabaseGuid;
            byte[] validDataForWrite1 = validIdset1.Serialize();
            byte[] validDataForWrite2 = validIdset2.Serialize();
            #endregion

            #region Step 2: Logon to a PrivateMailBox.
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            this.Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 4: Write a null data to the public folder.

            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = null;
            this.writePerUserInformationRequest.DataSize = 0;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1012");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1012.
            // If the RopWritePerUserInformation performs successfully on a private mailbox logon, MS-OXCSTOR_R1012 can be verified partially.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                1012,
                @"[In Receiving a RopWritePerUserInformation ROP Request] This operation [RopWritePerUserInformation] can be issued against either a private mailbox logon or a public folders logon.");
            #endregion

            #region Step 5: Call RopReadPerUserInformation to check no per user information exist in the specified public folder.
            this.readPerUserInformationRequest.FolderId = longTermIdForValidData;
            this.readPerUserInformationRequest.MaxDataSize = defaultDataSize;
            this.readPerUserInformationRequest.DataOffset = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1255");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1255.
            Site.CaptureRequirementIfIsNull(
                this.readPerUserInformationResponse.Data,
                1255,
                @"[In Private Mailbox Specific Behavior] If the row does not exist, then the server returns an empty array in the Data field of the response.");
            #endregion

            #region Step 6: Call RopWritePerUserInformation to write a valid user information.
            entireBlobSize = validDataForWrite1.Length;
            this.writePerUserInformationRequest.Data = validDataForWrite1;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.DataSize = (ushort)validDataForWrite1.Length;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            Site.Assert.AreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 7: Call GetPerUserGuid using the longTermID of the specified public folder.
            getPerUserGuidRequest.LongTermId = longTermIdForValidData;

            // Send a GetPerUserGuid request to the server and verify the success response.
            this.oxcstorAdapter.DoRopCall(this.getPerUserGuidRequest, this.outObjHandle, ROPCommandType.RopGetPerUserGuid, out this.outputBuffer);
            this.getPerUserGuidResponse = (RopGetPerUserGuidResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R488");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R488.
            // If the RopGetPerUserGuid performs successfully on a private mailbox logon, MS-OXCSTOR_R488 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getPerUserGuidResponse.ReturnValue,
                488,
                @"[In RopGetPerUserGuid ROP Request Buffer] This ROP [RopGetPerUserGuid] MUST be issued against a logon that was made to a private mailbox.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R490");

            bool isR490Verified = this.ByteArrayEquals(this.getPerUserGuidResponse.DatabaseGuid, this.logonResponse.ReplGuid);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R490.
            // If the RopGetPerUserGuid performs successfully on a private mailbox logon, MS-OXCSTOR_R490 can be verified.
            Site.CaptureRequirementIfIsTrue(
                isR490Verified,
                490,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] DatabaseGuid: Contains the REPLGUID of the last public folder database for which relevant read/unread information was cached.");

            #endregion

            #region Step 8: Call RopReadPerUserInformation with MaxDataSize set to zero.
            this.readPerUserInformationRequest.FolderId = longTermIdForValidData;
            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R965");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R965.
            // The RopReadPerUserInformation performs successfully on a private mailbox logon, if the RopReadPerUserInformation ROP performs successfully, MS-OXCSTOR_R965 can be partially verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                965,
                @"[In Receiving a RopReadPerUserInformation ROP Request] This operation [RopReadPerUserInformation] can be issued against either a private mailbox logon or a public folders logon.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1005");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1005
            // MaxDataSize be set to 0 server will use the default value 4096
            int adjMaxDataSize = 4096;
            int blobMinDataOffset = entireBlobSize - (int)this.readPerUserInformationRequest.DataOffset;
            bool isVerify_R1005 = (blobMinDataOffset > adjMaxDataSize)
                                 ? (this.readPerUserInformationResponse.DataSize == adjMaxDataSize)
                                 : (this.readPerUserInformationResponse.DataSize == blobMinDataOffset);
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1005,
                1005,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server MUST set DataSize to the lesser of the following two values [the adjusted value of MaxDataSize, the entire BLOB minus the value of DataOffset.].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1000");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1000.
            // The DataSize field in the RopReadPerUserInformation is the lesser of the following two values [the adjusted value of MaxDataSize, the entire BLOB minus the value of DataOffset.], it indicates the server compares the adjusted value of MaxDataSize to the size of the remaining BLOB segment.
            // MS-OXCSTOR_R1000 can be verified directly.
            Site.CaptureRequirement(
                1000,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] 	The server compares the adjusted value of MaxDataSize to the size of the remaining BLOB segment.");
            #endregion

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R516");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R569
            // At step5, call RopWritePerUserInformation ROP to write valid data to a folder and
            // at step6, call RopReadPerUserInformation Rop to read the information written at step5,
            // if could get the data, it indicates the data is saved. So, this requirement can be verified.
            bool isR516Verified = this.ByteArrayEquals(validDataForWrite1, this.readPerUserInformationResponse.Data);
            Site.CaptureRequirementIfIsTrue(
                isR516Verified,
                516,
                @"[In RopReadPerUserInformation ROP] When this ROP is issued against a private mailbox logon, cached data for the public folder is retrieved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R966");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R966.
            // MS-OXCSTOR_R516 has verified the server returned the cached date for the specified public folder successfully, MS-OXCSTOR_R966 can be verified directly.
            Site.CaptureRequirement(
                966,
                @"[In Private Mailbox Specific Behavior] The server searches the per-user data table for the mailbox for the only row with an FID equal to the value of the FolderId field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R967");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R967.
            // MS-OXCSTOR_R516 has verified the server returned the cached date for the specified public folder successfully, MS-OXCSTOR_R967 can be verified directly.
            Site.CaptureRequirement(
                967,
                @"[In Private Mailbox Specific Behavior] If the row exists, then the server retrieves from that row the stored change number set of read items.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R975");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R975
            // Verify the data is a BLOB which is formatted as a serialized IDSET with REPLGUID
            bool isVerify_R975 = this.VerifyDataIsIDSETStructure(this.readPerUserInformationResponse.Data);

            Site.CaptureRequirementIfIsTrue(
                isVerify_R975,
                975,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] [The change number set is serialized into a binary large object (BLOB) that is formatted as a serialized IDSET with REPLGUID structure, as specified in [MS-OXCFXICS] section 2.2.2.4.2.] The server then returns the BLOB in the Data field of the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R974");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R974.
            // MS-OXCSTOR_R975 has verified the change number set structure, MS-OXCSTOR_R545 can be verified directly.
            Site.CaptureRequirement(
                974,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The change number set MUST be serialized into a binary large object (BLOB) that is formatted as a serialized IDSET with REPLGUID structure, as specified in [MS-OXCFXICS] section 2.2.2.4.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R545");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R545.
            // MS-OXCSTOR_R975 has verified the change number set structure, MS-OXCSTOR_R545 can be verified directly.
            Site.CaptureRequirement(
                545,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] Data: Contains the change number set (CNSET), which is serialized into a binary large object (BLOB).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59729");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59729.
            // MS-OXCSTOR_R975 has verified the change number set structure, MS-OXCSTOR_R545 can be verified directly.
            Site.CaptureRequirement(
                59729,
                @"[in RopReadPerUserInformation ROP Success Response Buffer] Data: The format of the BLOB is the same as that of a serialized identifier set, which is specified in [MS-OXCFXICS] section 2.2.2.4.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1003");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1003
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)this.readPerUserInformationResponse.DataSize,
                (ushort)this.readPerUserInformationResponse.Data.Length,
                1003,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The DataSize field specifies the actual number of bytes that are returned in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1025");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1025
            // In the step 5, write the data into the server in 2 write ROPs, the server will accumulated the data (change number information). 
            // If succeed to get the data that write, it indicates a new row is added, and the row contains the accumulated change number information
            bool isVerify_R1025 = this.ByteArrayEquals(validDataForWrite1, this.readPerUserInformationResponse.Data);
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1025,
                1025,
                @"[In Private Mailbox Specific Behavior] If the row does not presently exist, a new row for the mailbox and folder MUST be added, setting the REPLGUID field and accumulated change number information onto that row.");
            #endregion
            #endregion

            #region Step 9: Call RopReadPerUserInformation with DataOffset less than zero.
            this.readPerUserInformationRequest.DataOffset = 0xffffffff;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            #region Capture

            if (Common.IsRequirementEnabled(1042, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1042");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1042.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80004005,
                    this.readPerUserInformationResponse.ReturnValue,
                    1042,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The implementation does fail the operation with 0x80004005 (ecError) in the ReturnValue field, if the value of the DataOffset field is less than zero. (Exchange 2013 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(1007, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R554");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R554.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.readPerUserInformationResponse.ReturnValue,
                    554,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecRpcFormat: Its value is 0x000004B6.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R555");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R555.
                // MS-OXCSTOR_R554 was captured by calls RopReadPerUserInformation with DataOffset value less than zero, MS-OXCSTOR_R555 can be verified directly.
                Site.CaptureRequirement(
                    555,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecRpcFormat: The DataOffset value was less than zero.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1007");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1007.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x000004B6,
                    this.readPerUserInformationResponse.ReturnValue,
                    1007,
                    @"[In Appendix A: Product Behavior] The implementation does fail the operation with 0x000004B6 (ecRpcFormat) in the ReturnValue field, if the value of the DataOffset field is less than zero. (<46> Section 3.2.5.12.1: Exchange 2003, Exchange 2007, and Exchange 2010 fail the operation with 0x000004B6 (ecRpcFormat).)");
            }
            #endregion
            #endregion

            #region Step 10: Restart the call of RopReadPerUserInformation with MaxDataSize field set to "entireBlobSize - 1".

            // Read once with MaxDataSize = entireBlobSize -1, the returned value of hasFinished in the response will be false,
            // all the data except the last byte will be returned, remaining data in the server will be 1
            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = (ushort)(entireBlobSize - 1);
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            RopReadPerUserInformationResponse responseForOnceRead = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R552");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R552.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseForOnceRead.ReturnValue,
                552,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R553");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R553.
            // The RopReadPerUserInformation ROP performs successfully and the ecNone error code was captured by MS-OXCSTOR_R552, MS-OXCSTOR_R553 can be verified directly.
            Site.CaptureRequirement(
                553,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1284");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1284.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                responseForOnceRead.HasFinished,
                1284,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: The value of this field is FALSE if the last block of data is not being returned.");

            #region Capture
            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCSTOR_R543,
                the value of the MaxDataSize field of the RopReadPerUserInformation ROP request is {0},
                the value of the DataSize field of the RopReadPerUserInformation ROP response is {1}.",
                this.readPerUserInformationRequest.MaxDataSize,
                responseForOnceRead.DataSize);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R543.
            bool isVerifiedR543 = responseForOnceRead.DataSize <= this.readPerUserInformationRequest.MaxDataSize;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR543,
                543,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] DataSize: The value of this field MUST be less than or equal to the value of the MaxDataSize field of the request.");
            #endregion

            #endregion

            #region Step 11: Continue the call of RopReadPerUserInformation with DataOffset field set to "dataSizeWritten - 1" and MaxDataSize field set to "1".

            // The remained 1 byte should be returned, and HasFinished will be true
            this.readPerUserInformationRequest.DataOffset = (ushort)(entireBlobSize - 1);
            this.readPerUserInformationRequest.MaxDataSize = 1;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            RopReadPerUserInformationResponse responseForAgainRead = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                responseForAgainRead.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1002");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1002
            // At step8, read again for the remaining Blob segment
            // verify if the size of returned data is equal to dataSizeWritten(entire blob size) minus the value of DataOffset
            bool isVerify_R1002 = responseForAgainRead.Data.Length == entireBlobSize - this.readPerUserInformationRequest.DataOffset;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1002,
                1002,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The size of the remaining BLOB segment is equal to the size of the entire BLOB minus the value of DataOffset.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1009");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1009.
            // At step8, read again for the remaining Blob segment, 
            // verify if the size of returned data is equal to dataSizeWritten(entire blob size) minus the value of DataOffset.
            bool isVerify_R1009 = responseForAgainRead.Data.Length == entireBlobSize - this.readPerUserInformationRequest.DataOffset;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R1009,
                1009,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] This [The size of the remaining BLOB segment] is equal to the size of the entire BLOB minus the value of DataOffset.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1010");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1010
            // Verify "HasFinshed" field is true "0x01" when the dataSize plus data offset equals the size of entire BloB
            Site.CaptureRequirementIfAreEqual<byte>(
                0x01,
                responseForAgainRead.HasFinished,
                1010,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server MUST set HasFinished to TRUE if DataOffset plus DataSize equals the size of the entire BLOB.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1011");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1011
            // "In other words" means a different statement, but for same verification logic with R1010
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0x00,
                responseForAgainRead.HasFinished,
                1011,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In other words, when the server sends the last segment of the BLOB, HasFinished MUST be set to TRUE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1008");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1008
            // Step8 is to read the remaining data after the first read at step7.
            // If the data size read at step8 is the remaining BLOB segment size, then could verify R1008
            bool isVerifiedR1008 = responseForAgainRead.DataSize == entireBlobSize - this.readPerUserInformationRequest.DataOffset;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1008,
                1008,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] This [The size of the remaining BLOB segment] is the size of the portion of the BLOB that remains to be sent to the client.");

            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R540");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R540.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                0,
                responseForAgainRead.HasFinished,
                540,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: The value of this field is TRUE if the last block of data is being returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R539");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R539.
            // MS-OXCSTOR_R540 and MS-OXCSTOR_R1284 was verified, MS-OXCSTOR_R539 can be verified directly.
            Site.CaptureRequirement(
                539,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished: Indicates whether the last block of data is being returned.");
            #endregion

            #region Step 12: Call RopReadPerUserInformation with incorrect value set in DataOffset field.

            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = (ushort)(entireBlobSize - 1);
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Read once with MaxDataSize = entireBlobSize -1, the returned value of hasFinished in the response will be false, 
            // all the data except the last byte will be returned, remaining data in the server will be 1
            this.readPerUserInformationRequest.DataOffset = (ushort)(entireBlobSize + 1);
            this.readPerUserInformationRequest.MaxDataSize = (ushort)(entireBlobSize - 1);
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1260");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1260
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.readPerUserInformationResponse.ReturnValue,
                1260,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If the value of the DataOffset field is greater than the size of the next BLOB segment to be returned, the server MUST fail the operation with 0x80004005 (ecError) in the ReturnValue field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R556");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R556
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.readPerUserInformationResponse.ReturnValue,
                556,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: Its value is 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R557");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R557.
            // MS-OXCSTOR_R556 has captured error code 0x80004005, MS-OXCSTOR_R557 can be verified directly.
            Site.CaptureRequirement(
                557,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The DataOffset value was greater than the data size.");
            #endregion

            #endregion

            #region Step 13: Call RopWritePerUserInformationRequest with new valid user information which expect to replace the information written previously.

            validIdset2.ReplGuid = longTermIdForValidData.DatabaseGuid;
            this.writePerUserInformationRequest.Data = validDataForWrite2;
            this.writePerUserInformationRequest.DataSize = (ushort)validIdset2.Size();
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            entireBlobSize = this.writePerUserInformationRequest.DataSize;

            Site.Assert.AreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 14: Call RopReadPerUserInformation to verify the per user information data was updated.
            this.readPerUserInformationRequest.Reserved = 0x00;
            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = defaultDataSize;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1024");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1024
            // If this.readPerUserInformationResponse.Data is not null, it means the row exists.
            // In the step 5, write the data into the server. 
            // In the step 9, write another valid data into the server with same FolderId
            // So verify the data got after step9 is same as the data written in step9, not same to the data written in step5
            bool isVerifiedR1024 =
                this.ByteArrayEquals(this.readPerUserInformationResponse.Data, validDataForWrite2) &&
                !this.ByteArrayEquals(this.readPerUserInformationResponse.Data, validDataForWrite1);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1024,
                1024,
                @"[In Private Mailbox Specific Behavior] If the row exists, the REPLGUID field and accumulated change number information MUST replace any existing values in the table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1023");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1023.
            // MS-OXCSTOR_R1024 has verified the server replace existing values successfully, it indicates the server get the per-user data table successfully by the user ID associated with the session logon and an FID. MS-OXCSTOR_R1023 can be verified.
            Site.CaptureRequirement(
                1023,
                @"[In Private Mailbox Specific Behavior] The server searches the per-user data table of the mailbox for the only row with an FID equal to the value of the FolderId field.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case verifies the RopWritePerUserInformation ROP related requirements in private mailbox logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S03_TC04_WritePerUserInformationForPrivateLogon()
        {
            this.CheckTransportIsSupported();

            #region Variable
            ushort maxDataSize = 4096;
            ushort defaultDataSize = 4096;

            IDSETWithReplGuid validIdset = this.GenerateRandomValidIdset();
            LongTermId longTermIdForValidData;
            LongTermId longTermIdForInValidData;
            #endregion

            #region Step 1: Connect and log on to a public folder
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");

            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            // Get LongTermID of the folder[1], used in below RopRead/WritePerUserInformation request with valid data
            longTermIdForValidData = this.GetLongTermIdFromId(this.logonResponse.FolderIds[1]);

            validIdset.ReplGuid = longTermIdForValidData.DatabaseGuid;

            // Get LongTermID of the folder[2], used in below RopRead/WritePerUserInformation request with invalid data
            longTermIdForInValidData = this.GetLongTermIdFromId(this.logonResponse.FolderIds[2]);
            #endregion

            #region Step 2: Disconnect.
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            this.Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully.");
            #endregion

            #region Step 3: Connect and log on to a private mailbox.
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 4: Call WritePerUserInformationRequest to write invalid data with HasFinished field set to false, so the server will cache these data.

            byte[] data = new byte[5000];
            this.writePerUserInformationRequest.FolderId = longTermIdForInValidData;
            this.writePerUserInformationRequest.DataSize = (ushort)data.Length;
            this.writePerUserInformationRequest.Data = data;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            if (this.writePerUserInformationResponse.ReturnValue == 0x8004011B)
            {
                IDSETWithReplGuid validIdset1 = this.GenerateRandomValidIdset(true);
                byte[] validDataForWrite1 = validIdset1.Serialize();
                this.writePerUserInformationRequest.DataSize = (ushort)(validDataForWrite1.Length);
                this.writePerUserInformationRequest.Data = validDataForWrite1;
                this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
                this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            }

            this.writePerUserInformationRequest.HasFinished = 0x00;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 5: Call RopReadPerUserInformation with MaxDataSize set to 4097.
            ushort bigerMaxDateSize = (ushort)(maxDataSize + 1);
            this.readPerUserInformationRequest.Reserved = 0x00;
            this.readPerUserInformationRequest.FolderId = longTermIdForInValidData;
            this.readPerUserInformationRequest.MaxDataSize = bigerMaxDateSize;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            #region Capture
            if (Common.IsRequirementEnabled(999, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R999");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R999.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    maxDataSize,
                    this.readPerUserInformationResponse.DataSize,
                    999,
                    @"[In Appendix A: Product Behavior] Implementation's maximum value of MaxDataSize is 4096. (If MaxDataSize > [server's suitable maximum (4096)], then Implementation does adjust the value of MaxDataSize to the suitable maximum value (4096) in Microsoft Exchanges. <48> Section 3.2.5.12.1: Exchange 2003, Exchange 2007, Exchange 2010, Exchange 2013 and Exchange 2016 use 4096 for the maximum value.)");
            }
            #endregion

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R979");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R979.
            Site.CaptureRequirementIfAreEqual<ushort>(
                maxDataSize,
                (ushort)this.readPerUserInformationResponse.Data.Length,
                979,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] On each invocation of RopReadPerUserInformation, the server inspects the value of the MaxDataSize field of the RopReadPerUserInformation request because the value can be different in each request.");

            if (Common.IsRequirementEnabled(997, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R997");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R997.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    maxDataSize,
                    (ushort)this.readPerUserInformationResponse.Data.Length,
                    997,
                    @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] Implementation does compare the value of MaxDataSize to some suitable maximum value (4096), as determined by the implementation in Microsoft Exchanges. (Exchange 2003 and above follow this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R981");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R981
            // Write the data with 5000 bytes to the server, and pass 4097 to the MaxDataSize field of RopReadPerUserInformation ROP request,
            // because 4097<5000, so the server will adjust the value of MaxDataSize to 4096. 
            // Thus, if check the length of the obtained data is not equal to 4097 that set, it indicates that the server adjusted the value of MaxDataSize.
            Site.CaptureRequirementIfAreNotEqual<ushort>(
                bigerMaxDateSize,
                (ushort)this.readPerUserInformationResponse.Data.Length,
                981,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] After the server has inspected and, if necessary, adjusted the value of MaxDataSize, the server compares the value to the size of the remaining BLOB segment.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R988");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R988.
            // MS-OXCSTOR_R981 has verified the server adjusted the value of MaxDataSize, MS-OXCSTOR_R988 can be verified directly.
            Site.CaptureRequirement(
                988,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server MUST adjust the MaxDataSize value in certain cases, as specified in item 2[When the client retrieves a BLOB in segments, the client can set MaxDataSize to a different value in each RopReadPerUserInformation request that is used to retrieve the BLOB.] of this summary.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R982");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R982.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                this.readPerUserInformationResponse.HasFinished,
                982,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If the adjusted MaxDataSize value is less than the size of the remaining BLOB segment, then the server MUST set HasFinished field to FALSE to indicate to the client that some data remains to be retrieved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R990");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R990
            // If the MaxDataSize is changed by the server, it indicates server inspect the value of MaxDataSize.
            Site.CaptureRequirementIfAreNotEqual<ushort>(
                bigerMaxDateSize,
                (ushort)this.readPerUserInformationResponse.Data.Length,
                990,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] [When the client retrieves a BLOB in segments, the client can set MaxDataSize to a different value in each RopReadPerUserInformation request that is used to retrieve the BLOB]Therefore, the server examines the value of MaxDataSize on each invocation of RopReadPerUserInformation as follows.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1001");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1001
            // First write the data with 5000 bytes, and the Exchange (all versions) uses 4096 for the maximum value to make
            // [size of remaining BLOB segment] > [adjusted MaxDataSize](4096). 
            // Then if the HasFinished is set to FALSE, this Requirement is verified.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                this.readPerUserInformationResponse.HasFinished,
                1001,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If [size of remaining BLOB segment] > [adjusted MaxDataSize], then the server MUST set HasFinished to FALSE to indicate to the client that additional requests are necessary to retrieve all of the remaining portions of the BLOB.");

            #endregion
            #endregion

            #region Step 6: Call RopReadPerUserInformation with MaxDataSize set to 0.
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            byte[] userInformation1 = this.readPerUserInformationResponse.Data;

            #region Capture
            if (Common.IsRequirementEnabled(1356, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1356");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1356.
                Site.CaptureRequirementIfAreEqual<ushort?>(
                    defaultDataSize,
                    this.readPerUserInformationResponse.DataSize,
                    1356,
                    @"[In Appendix A: Product Behavior] Implementation's default value of MaxDataSize is 4096. (If MaxDataSize equals 0, then the server MUST adjust the value of MaxDataSize to a suitable default value (4096). <47> Section 3.2.5.12.1: Exchange 2003, Exchange 2007, Exchange 2010, Exchange 2013, and Exchange 2016 use 4096 for the default value.)");
            }
            #endregion

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R980");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R980
            // At step14, request with MaxDataSize=0 when Blob segment in server is greater than MaxDataSize
            // So if the response data size is equal to MaxDataSize, then could verify this requirement.
            Site.CaptureRequirementIfAreNotEqual<ushort>(
                0,
                (ushort)this.readPerUserInformationResponse.DataSize,
                980,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In certain cases [if MaxDataSize equals 0], the server MUST adjust the value of MaxDataSize.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                @"Verify MS-OXCSTOR_R1004,
                the value of DataSize is {0},
                the adjusted value of the MaxDataSize field is {1}.",
                this.readPerUserInformationResponse.DataSize,
                defaultDataSize);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1004
            // request with MaxDataSize=0, server will adjust the MaxDataSize to default suitable MaxDataSize (4096)
            // Verify the returned data size does not exceed the 4096
            bool isVerifiedR1004 = this.readPerUserInformationResponse.DataSize <= maxDataSize;

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR1004,
                1004,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The value of DataSize MUST NOT exceed the adjusted value of the MaxDataSize field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R991");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R991
            // At step14, MaxDataSize=0 in the RopReadPerUserInformation ROP request,
            // but the data returned is not null. It indicates the MaxDataSize is changed by the server. 
            // So, the server must compare the value of MaxDataSize to zero.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                this.readPerUserInformationResponse.Data.Length,
                991,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server compares the value of MaxDataSize to zero.");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R976");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R976.
            // The size of the BLOB is 5000 which exceed the maximum amount of data (4096) that can be communicated in a single RopReadPerUserInformation response.
            // The ROP performs successfully, MS-OXCSTOR_R976 can be verified.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                976,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The size of the BLOB can potentially exceed the maximum amount of data that can be communicated in a single RopReadPerUserInformation response (section 2.2.1.12.2).");

            #endregion

            #region Step 7: Call RopReadPerUserInformation with DataOffset field based on the number of bytes received in the previous response and MaxDataSize field set to zero.
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.readPerUserInformationRequest.DataOffset = (uint)userInformation1.Length;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.readPerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            byte[] userInformation2 = this.readPerUserInformationResponse.Data;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R977");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R977.
            // The RopReadPerUserInformation was called twice to get the data which exceed the maximum amount of data that can be communicated in a single RopReadPerUserInformation response, MS-OXCSTOR_R977 can be verified.
            Site.CaptureRequirementIfIsTrue(
                userInformation1.Length + userInformation2.Length > maxDataSize,
                977,
                "[In Behavior Common to Both Private Mailbox and Public Folder Logon] For this reason [the size of the BLOB can potentially exceed the maximum amount of data that can be communicated in a single RopReadPerUserInformation response], the RopReadPerUserInformation ROP ([MS-OXCROPS] section 2.2.3.12) is designed to stream the data to the client by having the client invoke the ROP multiple times.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R986");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R986.
            // Set the DataOffset field of the second RopReadPerUserInformation based on the Data length of the first RopReadPerUserInformation, and the entire BLOB data was got successfully, MS-OXCSTOR_R986 can be verified directly.
            Site.CaptureRequirement(
                986,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The client updates DataOffset based on the number of bytes received in the previous response so that DataOffset always points to the first byte of the next BLOB segment to be returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R984");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R984.
            // Set the DataOffset field of the second RopReadPerUserInformation based on the Data length of the first RopReadPerUserInformation, and the entire BLOB data was get successfully, MS-OXCSTOR_R984 can be verified directly.
            Site.CaptureRequirement(
                984,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In other words, the value of DataOffset specifies the position within the BLOB of the first byte of data to be returned to the client.");
            #endregion

            #region Step 8: Call RopWritePerUserInformation to write subpart user information (all the data of user information except the last byte).
            byte[] dataForWrite = validIdset.Serialize();
            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.HasFinished = 0x00;
            byte[] dataHasBeenWritten = new byte[dataForWrite.Length - 1];
            Array.Copy(dataForWrite, dataHasBeenWritten, dataHasBeenWritten.Length);

            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R597");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R597.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                597,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R598");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R598.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                598,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecNone: Success.");
            #endregion

            #region Step 9: Call RopWritePerUserInformation to write the last byte of the user information to another public folder.
            // Get the last byte for write. This is the next write-ROP which will write the last byte of data
            // (the former part of  data has been written at former write-ROP at step16
            byte[] lastByteForWrite = new byte[1];
            this.writePerUserInformationRequest.FolderId = longTermIdForInValidData;
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R601");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R601.
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                601,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The FolderId didn't match the value on the previous call, AND THEN DataOffset wasn't zero.");
            #endregion

            #region Step 10: Call RopWritePerUserInformation with the wrong DataOffset to write the last byte of the user information.

            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)dataForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = (uint)(lastByteForWrite.Length + 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1019");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1019
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                1019,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] In addition, if the current invocation's DataOffset isn't zero, the ROP [RopWritePerUserInformation] MUST fail with a ReturnValue of 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R599");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R599
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.writePerUserInformationResponse.ReturnValue,
                599,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: Its value is 0x80004005.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R600");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R600.
            // MS-OXCSTOR_R599 has captured error code 0x80004005 and the DataOffset didn't match the size of the data written so far, MS-OXCSTOR_R600 can be verified directly.
            Site.CaptureRequirement(
                600,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The error code ecError: The DataOffset didn't match the size of the data written so far, AND THEN DataOffset wasn't zero.");
            #endregion
            #endregion

            #region Step 11: Write the last byte of the user information with the proper DataOffset and folder ID, the ROP expect to be failed because the server assume the previous operation was aborted.
            lastByteForWrite = new byte[1];
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = lastByteForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1033");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1033.
            // The previous operation was aborted
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                1033,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] If the DataOffset value does not equal the amount of data already written, the server MUST assume the previous operation was aborted.");
            #endregion

            #region Step 12: Rewrite sub-part user information again (all the data of user information except the last byte).
            dataForWrite = validIdset.Serialize();
            this.writePerUserInformationRequest.FolderId = longTermIdForValidData;
            this.writePerUserInformationRequest.Data = dataForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = 0;
            this.writePerUserInformationRequest.HasFinished = 0x00;
            dataHasBeenWritten = new byte[dataForWrite.Length - 1];
            Array.Copy(dataForWrite, dataHasBeenWritten, dataHasBeenWritten.Length);

            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.writePerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 13: Continue to write the last byte.
            lastByteForWrite = new byte[1];
            lastByteForWrite[0] = dataForWrite[dataForWrite.Length - 1];
            this.writePerUserInformationRequest.Data = lastByteForWrite;
            this.writePerUserInformationRequest.DataSize = (ushort)lastByteForWrite.Length;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.writePerUserInformationRequest.DataOffset = (ushort)(dataForWrite.Length - 1);
            this.writePerUserInformationRequest.HasFinished = 0x01;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.insideObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.writePerUserInformationResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 14: Call ReadPerUserInformationRequest to verify the sub-part data saved successfully.
            this.readPerUserInformationRequest.FolderId = longTermIdForValidData;
            this.readPerUserInformationRequest.DataOffset = 0;
            this.readPerUserInformationRequest.MaxDataSize = 0;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1016");

            bool isR1016Verified = this.ByteArrayEquals(this.readPerUserInformationResponse.Data, dataForWrite);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1016.
            Site.CaptureRequirementIfIsTrue(
                isR1016Verified,
                1016,
                @"[In Behavior Common to Both Private Mailbox and Public Folder Logon] The server determines whether the current invocation is a continuation of a previous invocation by examining the FolderId and DataOffset fields.");
            #endregion
        }
    }
}