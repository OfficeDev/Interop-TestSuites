namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// This scenario is to test the ROP against the private mailbox logon.
    /// </summary>
    [TestClass]
    public class S01_PrivateMailboxLogon : TestSuiteBase
    {
        #region Test Class Initialization
        /// <summary>
        /// Use ClassInitialize to run code before running the first test in the class
        /// </summary>
        /// <param name="context">Context value</param>
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
        /// This test case verifies the requirements of logging on to the private mailbox.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC01_TestLogonToPrivateMailBox()
        {
            this.CheckTransportIsSupported();

            #region Step1: connect server via user configured by "AdminUserName"
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon ROP with set USE_PER_MDB_REPLID_MAPPING to log on to the private mailbox
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // The Private flag is set
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping; // Set USE_PER_MDB_REPLID_MAPPING flag.
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            RopLogonResponse logonResponseAdmin = this.logonResponse;

            #region Capture requirements
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R165");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R165
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                165,
                @"[In RopLogon ROP Common Return Codes] The value of return code ecNone: 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R166");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R166
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                166,
                @"[In RopLogon ROP Common Return Codes] The meaning of return code ecNone: Success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R56");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R56
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                56,
                @"[In RopLogon ROP Request Buffer] [LogonFlags] The description of the flag Private: This flag is set for logon to a private mailbox and is not set for logon to public folders.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R785");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R785
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                785,
                @"[In Receiving a RopLogon ROP Request] If the LogonFlags field has the Private bit set, the logon is going to a private mailbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R124");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R124
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.OwnerRight,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.OwnerRight),
                124,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The value of flag OwnerRight: 0x02.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R126");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R126
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.SendAsRight,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.SendAsRight),
                126,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The value of flag SendAsRight: 0x04.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R127");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R127
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.SendAsRight,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.SendAsRight),
                127,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The description of flag SendAsRight: The user has the right to send mail from the mailbox.");

            PropertyTag propertyPidTagOutOfOfficeState;
            uint resultSetPidTagOutOfOfficeState;
            TaggedPropertyValue pidTagOutOfOfficeState = new TaggedPropertyValue();

            // According to the Open Specification MS-OXPROPS, PidTagOutOfOfficeState's id is 0x661D
            const ushort PropertyIdForPidTagOutOfOfficeState = 0x661D;

            propertyPidTagOutOfOfficeState.PropertyId = PropertyIdForPidTagOutOfOfficeState;

            // According to [MS-OXPROPS], PidTagOutOfOfficeState's data type is 0x000B
            propertyPidTagOutOfOfficeState.PropertyType = 0x000B;
            pidTagOutOfOfficeState.PropertyTag = propertyPidTagOutOfOfficeState;
            pidTagOutOfOfficeState.Value = new byte[1];
            pidTagOutOfOfficeState.Value[0] = 0x01;

            resultSetPidTagOutOfOfficeState = this.TrySetLogonProperty(pidTagOutOfOfficeState);

            // Wait 5 seconds to enable the PidTagOutOfOfficeState
            Thread.Sleep(5000);

            bool valueOfInitial = BitConverter.ToBoolean(this.TryGetLogonPropertyValue(propertyPidTagOutOfOfficeState), 0);
            Site.Assert.IsTrue(valueOfInitial, "True indicates getting PidTagOutOfOfficeState successfully.");

            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R128");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R128
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.OOF,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.OOF),
                128,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The value of flag OOF: 0x10.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R129");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R129
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.OOF,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.OOF),
                129,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] The description of flag OOF: The Out of Office (OOF) state is set on the mailbox.");

            ushort replId = BitConverter.ToUInt16(this.logonResponse.ReplId, 0);
            this.longTermIdFromIdRequest.ObjectId = (ulong)replId;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.longTermIdFromIdResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            bool isReplIdConsistent = Common.CompareByteArray(this.logonResponse.ReplGuid, this.longTermIdFromIdResponse.LongTermId.DatabaseGuid);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R131
            bool isVerifiedR131 = isReplIdConsistent;
            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR131,
                131,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] ReplId: This field contains the short form of the value specified in the REPLGUID field, which is the replica ID for the logon.");

            // Confirm the FIDs are correct and all in the server.
            bool isCorrectFolderId = false;
            foreach (ulong folderId in this.logonResponse.FolderIds)
            {
                if (folderId != 0)
                {
                    isCorrectFolderId = true;
                }
                else
                {
                    isCorrectFolderId = false;
                    break;
                }
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R767");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R767
            // if server return FIDs, the FolderIds field must not be null, then this requirement will be verified.
            Site.CaptureRequirementIfIsTrue(
                isCorrectFolderId,
                767,
                @"[In Abstract Data Model] mailbox table: If the logon is allowed, then the Folder IDs of various special folders are obtained from the table and returned to the client. For more details about special folders, see [MS-OXOSFLD].");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R821");

            // Verify MS-OXCSTOR requirement: 821
            // For each FolderId returned in the log on response, it will be set in the Request IdFromLongTermId to check its validity.
            // The requirement will be verified just when all request of IdFromLongTermId succeeds.
            Site.CaptureRequirementIfIsTrue(
                isCorrectFolderId,
                821,
                @"[In Private Mailbox Logon] The server then determines the appropriate FIDs to return to the client. For details, see section 2.2.1.1.3.");
            #endregion
            #endregion Step2

            #region Step3: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
            #endregion Step3

            #region Step4: Connect server via user configured by "AdminUserName"
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step4

            #region Step5: RopLogon to private mailbox with NO_MAIL flag
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // The Private flag is set
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.NoMail | (uint)OpenFlags.UsePerMDBReplipMapping | (uint)OpenFlags.UseAdminPrivilege; // Set NO_MAIL, USE_ADMIN_PRIVILEGE and USE_PER_MDB_REPLID_MAPPING flags.
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R72");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R72
            this.Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.OwnerRight,
                (byte)(this.logonResponse.ResponseFlags & (byte)ResponseFlags.OwnerRight),
                72,
                @"[In RopLogon ROP Request Buffer] [OpenFlags] The description of flag USE_ADMIN_PRIVILEGE: A request for administrative access to the mailbox.");
            #endregion Step5

            #region Step6: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion Step6

            #region Step7: Connect server via the user that has no owner right on the mailbox of user configured by "AdminUserName".
            this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PrivateMailboxServer, this.essdnOfUser1, this.domain, this.userNameOfMailboxOnServer1, this.passwordOfUserOfMailboxOnServer1);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step7

            #region Step8: RopLogon ROP with set USE_PER_MDB_REPLID_MAPPING
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // Set the Private flag
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping; // Set USE_PER_MDB_REPLID_MAPPING flag.
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            RopLogonResponse logonResponseWithoutOwnerRight = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                logonResponseWithoutOwnerRight.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            bool ownerRightIsConsistent = (byte)(logonResponseAdmin.ResponseFlags & (byte)ResponseFlags.OwnerRight) == (byte)ResponseFlags.OwnerRight
                                            && (byte)(logonResponseWithoutOwnerRight.ResponseFlags & (byte)ResponseFlags.OwnerRight) != (byte)ResponseFlags.OwnerRight;

            this.Site.Assert.AreNotEqual<byte>(
                (byte)ResponseFlags.OwnerRight,
                (byte)(logonResponseWithoutOwnerRight.ResponseFlags & (byte)ResponseFlags.OwnerRight),
                "If the user has not owner right, the OwnerRight(0x02) is not set.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R125");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R125
            bool isVerifiedR125 = ownerRightIsConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR125,
                125,
                @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The description of flag OwnerRight: The user has owner permission on the mailbox.");
            #endregion

            #region Step9: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion Step9

            #region Step10: connect server via user configured by "AdminUserName"
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step10

            #region Step11: RopLogon ROP with set HOME_LOGON
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // Set the Private flag
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.HomeLogon | (uint)OpenFlags.UsePerMDBReplipMapping; // Set the HOME_LOGON flag
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            RopLogonResponse homeLogonResponse = this.logonResponse;
            #endregion Step11

            #region Step12: RopLogon ROP without set HOME_LOGON
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // Set the Private flag
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            RopLogonResponse nonHomeLogonResponse = this.logonResponse;

            #region Capture
            if (Common.IsRequirementEnabled(7900, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R7900, when set HOME_LOGON the return value is: {0}, when not set HOME_LOGON the return value is: {1}", homeLogonResponse.ReturnValue, nonHomeLogonResponse.ReturnValue);

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R7900            
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    homeLogonResponse.ReturnValue,
                    nonHomeLogonResponse.ReturnValue,
                    7900,
                    @"[In Appendix A: Product Behavior] The implementation returns the same response, no matter whether this flag [HOME_LOGON] is set or not. (Exchange 2010 and above follow this behavior.)");
            }
            #endregion capture

            #endregion

            #region Step13: RopLogon ROP log on to the private mailbox with set OpenFlags to TAKE_OWNERSHIP and USE_PER_MDB_REPLID_MAPPING
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.TakeOwnerShip | (uint)OpenFlags.UsePerMDBReplipMapping;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step13
        }

        /// <summary>
        /// This test case verifies the requirements of logging on to a mailbox that is different from the first logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC02_TestLogonToPrivateMailBox_AdditionalSession()
        {
            this.CheckTransportIsSupported();

            Site.Assume.IsFalse(this.transportByMAPIOrNot, "This case doesn't support the mapi_http transport.");

            #region Initialization: user configured by "UserNameOfMailboxOnServer1" logs on to the mailbox
            this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PrivateMailboxServer, this.essdnOfUser1, this.domain, this.userNameOfMailboxOnServer1, this.passwordOfUserOfMailboxOnServer1);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");

            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.essdnOfUser1 + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
            #endregion

            #region Step1: Disable the auto-redirect
            this.oxcstorAdapter.SetAutoRedirect(false);
            #endregion Step1

            if (Common.IsRequirementEnabled(3145, this.Site))
            {
                #region Step2: Connect server1 via user configured by "AdminUserName"
                this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step2

                #region Step3: RopLogon ROP without set OpenFlags
                this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = (uint)0;
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

                #region Capture
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3145");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3145
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.logonResponse.ReturnValue,
                    3145,
                    @"[In Appendix A: Product Behavior] For the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field, the implementation has the following behavior: If the logon is the first on the RPC session, or if the logon is additional on the RPC session and it is to the same mailbox that is associated with the first logon, then the server ignores the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior.)");
                #endregion capture
                #endregion Step3

                #region Step4: Disconnect
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
                #endregion Step4

                #region Step5: Connect server1 via user configured by "AdminUserName"
                this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step5

                #region Step6:  RopLogon ROP with set USE_PER_MDB_REPLID_MAPPING
                this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping; // Set USE_PER_MDB_REPLID_MAPPING flags.
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
                RopLogonResponse logonResponse = this.logonResponse;
                #endregion Step6

                #region Step7: RopLogon ROP without set OpenFlags
                this.logonRequestForPrivateMailBox.OpenFlags = (uint)0; // Does not set USE_PER_MDB_REPLID_MAPPING flags.
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                RopLogonResponse logonResponseAdditional = this.logonResponse;

                #region Capture
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3145");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3145
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    logonResponse.ReturnValue,
                    logonResponseAdditional.ReturnValue,
                    3145,
                    @"[In Appendix A: Product Behavior] For the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field, the implementation has the following behavior: If the logon is the first on the RPC session, or if the logon is additional on the RPC session and it is to the same mailbox that is associated with the first logon, then the server ignores the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior.)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R133");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R133
                bool isVerifyR133 = Common.CompareByteArray(logonResponse.ReplGuid, logonResponseAdditional.ReplGuid);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR133,
                    133,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] ReplGuid: If the client did not set the USE_PER_MDB_REPLID_MAPPING bit in the OpenFlags field, this value MUST be identical for all private mailbox logons on the same remote procedure call (RPC) session.");
                #endregion capture
                #endregion Step7

                #region Step8: Disconnect
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
                #endregion Step8
            }

            #region Step9: Connect server1 via user configured by "AdminUserName"
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step9

            #region Step10:  RopLogon ROP with set USE_PER_MDB_REPLID_MAPPING to logon the mailbox of user configured by "AdminUserName"
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private; // Set the Private flag
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping; // Set USE_PER_MDB_REPLID_MAPPING flags.
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            RopLogonResponse logonToUser1 = this.logonResponse;
            Site.Assert.AreEqual<uint>(0, logonToUser1.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step10

            #region Step11:  RopLogon ROP without set USE_PER_MDB_REPLID_MAPPING to logon another mailbox
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)0; // Does not set USE_PER_MDB_REPLID_MAPPING flags.
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.essdnOfUser1 + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            RopLogonResponse logonToUser2 = this.logonResponse;

            #region Capture
            if (Common.IsRequirementEnabled(1204, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1204");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1204
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    logonToUser2.ReturnValue,
                    1204,
                    @"[In Private Mailbox Logon] For the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field, the server has the following behavior: If the USE_PER_MDB_REPLID_MAPPING flag is not set, then Implementation does fail the ROP with a ReturnValue of ecInvalidParameter 0x80070057.<29> (Exchange 2013 and above follow this behavior)");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1800001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1800001
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.logonResponse.ReturnValue,
                    1800001,
                    @"[In RopLogon ROP Common Return Codes] The value of return code ecInvalidParameter: 0x80070057.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1820001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1820001
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.logonResponse.ReturnValue,
                    1820001,
                    @"[In RopLogon ROP Common Return Codes] The meaning of return code ecInvalidParameter: The client has not set the USE_PER_MDB_REPLID_MAPPING flag.");
            }

            if (Common.IsRequirementEnabled(800, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R800");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R800
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000478,
                    logonToUser2.ReturnValue,
                    800,
                    @"[In Appendix A: Product Behavior] The implementation does fail the RopLogon with ecWrongServer (0x00000478) [if the USE_PER_MDB_REPLID_MAPPING flag is not set]. (<29> Section 3.2.5.1.1: Exchange 2010 return ecWrongServer (0x00000478).)");
            }

            if (Common.IsRequirementEnabled(3133, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3133");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3133
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonToUser2.ReturnValue,
                    3133,
                    @"[In Appendix A: Product Behavior] The implementation does return the RopLogon with Success. <28> Section 3.2.5.1.1: If the USE_PER_MDB_REPLID_MAPPING flag is not set, Exchange 2007 does not fail the ROP and instead has the following behavior: Exchange 2007 maintains one REPLID-to-REPLGUID mapping and one named property-to-property ID mapping per RPC session, and these mappings are shared by all logons on the TRPC session.)");
            }

            if (Common.IsRequirementEnabled(1204, this.Site) || Common.IsRequirementEnabled(800, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R802");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R802
                // If the logon is additional on the RPC session and it is to a mailbox that is different from the mailbox that is associated with the first logon and the USE_PER_MDB_REPLID_MAPPING flag is not set, the server will return an error.
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    this.logonResponse.ReturnValue,
                    802,
                    @"[In Private Mailbox Logon] For the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field, the server has the following behavior: If the logon is additional on the RPC session and it is to a mailbox that is different from the mailbox that is associated with the first logon, the server inspects the USE_PER_MDB_REPLID_MAPPING flag of the OpenFlags field to see if it is set.");
            }
            #endregion capture
            #endregion Step11

            #region Step12: Revert to enable the auto redirect
            this.oxcstorAdapter.SetAutoRedirect(true);
            #endregion Step12
        }

        /// <summary>
        /// This test case verifies the notifications using RopGetReceiveFolder, RopSetReceiveFolder and RopGetReceiveFolderTable ROPs.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC03_TestRopSetAndGetReceiveFolderAndFolderTable()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful.");
            #endregion Step1

            // This operation is only valid when the Logon object refers to a private mailbox logon
            #region Step2: RopLogon ROP log on to the private mailbox
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private;
            this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(0x00000000, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step2

            #region Step3: RopSetReceiveFolder ROP with MessageClass="My.Class" and FolderId = 0 to remove the table row for the specified message class
            this.setReceiveFolderRequest.FolderId = 0;
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, this.setReceiveFolderResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step3

            #region Step4: RopGetReceiveFolder ROP to confirm whether the MessageClass "My.Class" exists in the server
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, this.getReceiveFolderResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R852");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R852
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                this.getReceiveFolderResponse.MessageClass[this.getReceiveFolderResponse.MessageClass.Length - 1],
                852,
                @"[In Receiving a RopGetReceiveFolder ROP Request] If no entry in the table [Receive folder table] can be matched, the server returns an empty string in the ExplicitMessageClass field. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1223");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1223
            // inbox folder id is this.logonResponse.FolderIds[4]
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.logonResponse.FolderIds[4],
                (ulong)this.getReceiveFolderResponse.FolderId,
                1223,
                @"[In Receiving a RopGetReceiveFolder ROP Request] If no entry in the table [Receive folder table] can be matched, the server returns [an empty string in the ExplicitMessageClass field and] the Folder ID for the user's Inbox folder in the FolderId field.");
            #endregion Capture
            #endregion Step4

            #region Step5: RopSetReceiveFolder ROP with MessageClass="My.Class"
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[3];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            DateTime modificationTime = DateTime.UtcNow;
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59756");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59756
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.setReceiveFolderResponse.ReturnValue,
                59756,
                @"[in Receiving a RopSetReceiveFolder ROP Request] The value of error code ecNone: 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59757");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59757
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.setReceiveFolderResponse.ReturnValue,
                59757,
                @"[in Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecNone: Success.");
            #endregion Step5

            #region Step6: RopGetReceiveFolder ROP with MessageClass "My.Class"
            this.getReceiveFolderRequest.MessageClass = this.setReceiveFolderRequest.MessageClass;
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            RopGetReceiveFolderResponse getReceiveFolderResponseWithSetClass = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R232");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R232
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getReceiveFolderResponseWithSetClass.ReturnValue,
                232,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The value of error code ecNone: 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R233");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R233
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                getReceiveFolderResponseWithSetClass.ReturnValue,
                233,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The meaning of error code ecNone: Success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1224");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1224
            this.Site.CaptureRequirementIfAreNotEqual<ulong>(
                0,
                getReceiveFolderResponseWithSetClass.FolderId,
                1224,
                @"[In Receiving a RopGetReceiveFolder ROP Request] If a match is found, the server returns a string specifying the actual configured message class and the Folder ID of the associated Receive folder.");

            // The receive folder for MessageClass "My.Class" has been removed in step 2. So if the receive folder for MessageClass "My.Class" can be got after set, it indicates that new row is added.
            bool newRowAdded = getReceiveFolderResponseWithSetClass.FolderId != 0 && getReceiveFolderResponseWithSetClass.MessageClass != null && getReceiveFolderResponseWithSetClass.MessageClass[0] != 0x00;

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59754. The added MessageClass: {0}, the folder id: {1}.", Encoding.ASCII.GetString(getReceiveFolderResponseWithSetClass.MessageClass).Replace("\0", string.Empty), getReceiveFolderResponseWithSetClass.FolderId);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59754
            bool isVerifiedR59754 = newRowAdded;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR59754,
                59754,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If a match is not found, a new row is added with the MessageClass and FolderId field values.");

            ulong folderIdForMyClass = getReceiveFolderResponseWithSetClass.FolderId;

            Site.Log.Add(LogEntryKind.Comment, "The RopGetReceiveFolder ROP with the MessageClass: {0}", Encoding.ASCII.GetString(getReceiveFolderResponseWithSetClass.MessageClass).Replace("\0", string.Empty));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R200");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R200
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0,
                getReceiveFolderResponseWithSetClass.MessageClass[getReceiveFolderResponseWithSetClass.MessageClass.Length - 1],
                200,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] includes the terminating NULL character.");

            bool isByteValue32To126 = true;
            for (int i = 0; i < getReceiveFolderResponseWithSetClass.MessageClass.Length - 1; i++)
            {
                byte byteValue = getReceiveFolderResponseWithSetClass.MessageClass[i];
                if (byteValue < 32 || byteValue > 126)
                {
                    isByteValue32To126 = false;
                    break;
                }
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R204");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R204
            bool isVerifiedR204 = isByteValue32To126;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR204,
                204,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: Each character value in the string is in the numeric range of 32 to 126, inclusive.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R202");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R202
            bool isVerifiedR202 = isByteValue32To126;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR202,
                202,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: The string uses ASCII encoding.");

            Site.Log.Add(LogEntryKind.Comment, "The length of the MessageClass: {0}", getReceiveFolderResponseWithSetClass.MessageClass.Length);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R203");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R203
            bool isVerifiedR203 = getReceiveFolderResponseWithSetClass.MessageClass.Length > 0 && getReceiveFolderResponseWithSetClass.MessageClass.Length <= 255;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR203,
                203,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: The length (including the terminating NULL character) is greater than zero and less than or equal to 255.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R210");

            string setMessageClass = Encoding.ASCII.GetString(this.setReceiveFolderRequest.MessageClass);
            string gotMessageClass = Encoding.ASCII.GetString(getReceiveFolderResponseWithSetClass.MessageClass);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R210
            this.Site.CaptureRequirementIfAreEqual<string>(
                setMessageClass.ToUpper(),
                gotMessageClass.ToUpper(),
                210,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: This field [ExplicitMessageClass] contains a string specifying the message class that is actually configured for the Receive folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R211");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R211
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0,
                getReceiveFolderResponseWithSetClass.MessageClass[getReceiveFolderResponseWithSetClass.MessageClass.Length - 1],
                211,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] includes the terminating NULL character.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1235");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1235
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.setReceiveFolderRequest.FolderId,
                (ulong)getReceiveFolderResponseWithSetClass.FolderId,
                1235,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If a match is found, the value of the FolderId field replaces the Folder ID, as specified in [MS-OXCDATA] section 2.2.1.1, stored in the table on that row.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R208");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R208
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                this.setReceiveFolderRequest.FolderId,
                (ulong)getReceiveFolderResponseWithSetClass.FolderId,
                208,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] FolderId: This field [FolderId] contains the Folder Id, as specified in [MS-OXCDATA] section 2.2.1.1, of the folder to which messages are being delivered.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R195");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R195
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                this.setReceiveFolderRequest.FolderId,
                (ulong)getReceiveFolderResponseWithSetClass.FolderId,
                195,
                @"[In RopGetReceiveFolder ROP] The RopGetReceiveFolder ROP ([MS-OXCROPS] section 2.2.3.2) is used to determine the Receive folder for messages of a specific message class.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R196");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R196
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                this.setReceiveFolderRequest.FolderId,
                (ulong)getReceiveFolderResponseWithSetClass.FolderId,
                196,
                @"[In RopGetReceiveFolder ROP] This ROP [RopGetReceiveFolder] examines the message class string and returns the folder ID (FID) ([MS-OXCDATA] section 2.2.1.1) of the Receive folder to which messages of that class and all subclasses are delivered.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R851");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R851
            bool isVerifiedR851 = Encoding.ASCII.GetString(getReceiveFolderResponseWithSetClass.MessageClass).ToUpper() == (this.messageMyClass.ToUpper() + '\0') &&
                getReceiveFolderResponseWithSetClass.FolderId == this.logonResponse.FolderIds[3];

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R851
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR851,
                851,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server then retrieves the actual message class string from the Receive folder table, and the associated Folder ID, as specified in [MS-OXCDATA] section 2.2.1.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R847");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R847
            // Because the MessageClass is "My.Class", not "MY.CLASS", so this requirement can be verified.
            Site.CaptureRequirementIfAreEqual<string>(
                this.messageMyClass.ToUpper() + "\0",
                Encoding.ASCII.GetString(getReceiveFolderResponseWithSetClass.MessageClass).ToUpper(),
                847,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server validates the value of the MessageClass field.] The server then searches the Receive folder table to find the entry with the longest case-insensitive prefix string that matches the value of the MessageClass field.");
            #endregion capture
            #endregion Step6

            #region Step7: RopGetReceiveFolder ROP with MessageClass=IPM
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes("IPM\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            RopGetReceiveFolderResponse getReceiveFolderResponseWithIPM;
            getReceiveFolderResponseWithIPM = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, getReceiveFolderResponseWithIPM.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step7

            #region Step8: RopGetReceiveFolder ROP with MessageClass=ipm
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes("ipm\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            RopGetReceiveFolderResponse getReceiveFolderResponseWithipm;
            getReceiveFolderResponseWithipm = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, getReceiveFolderResponseWithipm.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R201");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R201
            bool isVerifiedR201 = getReceiveFolderResponseWithIPM.ReturnValue == 0 && getReceiveFolderResponseWithipm.ReturnValue == 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR201,
                201,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: Examination of the string [Contained by MessageClass] is case-insensitive.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R862");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R862
            Site.CaptureRequirementIfAreEqual<ulong>(
                getReceiveFolderResponseWithIPM.FolderId,
                getReceiveFolderResponseWithipm.FolderId,
                862,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The server searches the Receive folder table using a case-insensitive string comparison for an exact match to the value of the MessageClass field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1225");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1225
            Site.CaptureRequirementIfAreEqual<uint>(
                getReceiveFolderResponseWithipm.ReturnValue,
                getReceiveFolderResponseWithIPM.ReturnValue,
                1225,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server can case-fold the string to all uppercase or all lowercase, or leave the string as stored.");

            #endregion capture
            #endregion Step8

            #region Step9: RopGetReceiveFolderTable ROP to get the receive folder table
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderTableRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolderTable, out this.outputBuffer);
            this.getReceiveFolderTableResponse = (RopGetReceiveFolderTableResponse)outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R320");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R320
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                this.getReceiveFolderTableResponse.ReturnValue,
                320,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] the error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R321");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R321
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                this.getReceiveFolderTableResponse.ReturnValue,
                321,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] the error code ecNone: Success.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R287");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R287
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                this.getReceiveFolderTableResponse.ReturnValue,
                287,
                @"[In RopGetReceiveFolderTable ROP Request Buffer] This operation [RopGetReceiveFolderTable] MUST be issued against a private mailbox logon.");

            // Get the FID of the MessageClass "My.Class" and its Last-modification Time
            int createdClassIndex = 0;
            List<ulong> folderIds = new List<ulong>();
            List<string> messageClasses = new List<string>();
            List<ulong> lastModificationTime = new List<ulong>();
            for (int i = 0; i < this.getReceiveFolderTableResponse.RowCount; i++)
            {
                PropertyRow receiveFolderRow = this.getReceiveFolderTableResponse.Rows.PropertyRows[i];
                Site.Assert.IsTrue(
                    receiveFolderRow.PropertyValues.Count >= 3,
                    "Each row of the Receive folder table must have at least 3 properties: PidTagFolderId, PidTagMessageClass and PidTagLastModificationTime. Actually {0} properties.",
                    receiveFolderRow.PropertyValues.Count);

                if (Encoding.ASCII.GetString(receiveFolderRow.PropertyValues[1].Value).ToUpper() == (this.messageMyClassToUpper + '\0'))
                {
                    createdClassIndex = i;
                }

                if (receiveFolderRow.Flag == 0x00)
                {
                    // PidTagFolderId property 
                    ulong folderId = BitConverter.ToUInt64(receiveFolderRow.PropertyValues[0].Value, 0);
                    folderIds.Add(folderId);

                    // PidTagMessageClass property 
                    byte[] messageClassBytes = receiveFolderRow.PropertyValues[1].Value;
                    string messageClass = Encoding.ASCII.GetString(messageClassBytes);
                    messageClasses.Add(messageClass.ToUpper());

                    // PidTagLastModificationTime property 
                    ulong modificationTicks = BitConverter.ToUInt64(receiveFolderRow.PropertyValues[2].Value, 0);
                    lastModificationTime.Add(modificationTicks);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R868");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R868
                    // If the above 3 properties can be parsed successfully, indicates the behavior is consistent with the requirement.
                    Site.CaptureRequirement(
                        868,
                        @"[In Receiving a RopSetReceiveFolder ROP Request] Each row of the table contains at least the following three columns [Folder ID, Message Class, and Last-modification Time], with each column corresponding to a property.");

                    if (messageClass.Equals("My.Class\0", StringComparison.InvariantCultureIgnoreCase))
                    {
                        DateTime actualModificationTime = DateTime.FromFileTimeUtc(BitConverter.ToInt64(receiveFolderRow.PropertyValues[2].Value, 0));

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R314");

                        TimeSpan timeSpan = modificationTime - actualModificationTime;
                        // Test case has recorded time after execute RopSetReceiveFolder ROP with MessageClass="My.Class", 
                        // If the time span between it and actual modification time is less than 1 second, this requirement can be captured. 
                        bool isEqual = timeSpan.TotalSeconds < 1;

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R314
                        Site.CaptureRequirementIfIsTrue(
                            isEqual,
                            314,
                            @"[In RopGetReceiveFolderTable ROP Success Response Buffer]  [Rows] PidTagLastModificationTime property: PidTagLastModificationTime property ([MS-OXPROPS] section 2.755) -- A PtypTime value that specifies the time, in Coordinated Universal Time (UTC), when the server created or last modified the row in the Receive folder table.");
                    }
                }
                else
                {
                    // The receive folder row encounters error.
                    folderIds.Add(0);
                    messageClasses.Add("\0");
                    lastModificationTime.Add(0);
                }
            }

            List<string> initializedReceiveFolder = new List<string>(new string[] { "\0", "IPM\0", "REPORT.IPM\0", "IPC\0" });

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R782");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R782
            bool isMessageClassConsistent = false;
            foreach (string messageClass in initializedReceiveFolder)
            {
                if (messageClasses.Contains(messageClass.ToUpper()))
                {
                    isMessageClassConsistent = true;
                }
                else
                {
                    isMessageClassConsistent = false;
                    break;
                }
            }

            bool isVerifyR782 = isMessageClassConsistent;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR782,
                782,
                @"[In Initialization] When a mailbox is created, the following entries MUST be added to the Receive folder table for the new mailbox: """" (empty string) - Inbox in the new mailbox; ""IPM"" - Inbox in the new mailbox; ""Report.IPM"" - Inbox in the new mailbox; ""IPC"" - Root folder of the new mailbox.");

            string createdClassWithNullTerminatingChar = messageMyClass + '\0';
            string returnedMyClass = Encoding.ASCII.GetString(this.getReceiveFolderTableResponse.Rows.PropertyRows[createdClassIndex].PropertyValues[1].Value);
            bool isMessageClassCaseConsistent = returnedMyClass == createdClassWithNullTerminatingChar || returnedMyClass == createdClassWithNullTerminatingChar.ToLower() || returnedMyClass == createdClassWithNullTerminatingChar.ToUpper();

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R306, The MessageClass calling RopGetReceiveFolderTable ROP: {0}", returnedMyClass.Replace("\0", string.Empty));

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R306
            bool isVerifiedR306 = isMessageClassCaseConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR306,
                306,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property:  The string [PidTagMessageClass property] can be all upper case, all lower case, or as originally stored by the client.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R872");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R872
            bool isVerifiedR872 = isMessageClassCaseConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR872,
                872,
                @"[In Receiving a RopSetReceiveFolder ROP Request] ""Message Class"" column (PidTagMessageClass property ([MS-OXPROPS] section 2.778)) - Contains a string that specifies the message class that is configured for the Receive folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59801, The MessageClass calling RopGetReceiveFolderTable ROP: {0}", returnedMyClass.Replace("\0", string.Empty));

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59801
            bool isVerifiedR59801 = isMessageClassCaseConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR59801,
                59801,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The case of the characters in the string [Contained by ExplicitMessageClass] is insignificant.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R870");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R870
            this.Site.CaptureRequirementIfAreEqual<ulong>(
                folderIdForMyClass,
                BitConverter.ToUInt64(this.getReceiveFolderTableResponse.Rows.PropertyRows[createdClassIndex].PropertyValues[0].Value, 0),
                870,
                @"[In Receiving a RopSetReceiveFolder ROP Request] ""Folder ID"" column (PidTagFolderId property ([MS-OXPROPS] section 2.691)) - Contains the Folder ID, as specified in [MS-OXCDATA] section 2.2.1.1, of the Receive folder, which is the folder to which messages of the specified message class will be delivered.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R866, The MessageClass calling RopGetReceiveFolderTable ROP: {0}", returnedMyClass.Replace("\0", string.Empty));

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R866
            bool isVerifiedR866 = isMessageClassCaseConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR866,
                866,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The server can case-fold the value of the MessageClass field to upper case or lower case, or leave the value unchanged before storage.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R880");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R880
            bool isVerifiedR880 = isMessageClassCaseConsistent;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR880,
                880,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] The server can convert message class values to all upper case or all lower case or return the value as stored.");
            #endregion capture
            #endregion Step9

            #region Step10: RopSetReceiveFolder ROP with FolderId = 0x0000000000000000 and MessageClass="My.Class"
            this.setReceiveFolderRequest.FolderId = 0x0000000000000000;
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.setReceiveFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step10

            #region Step11: RopGetReceiveFolder ROP with the specified MessageClass "My.Class"
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.getReceiveFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step11

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R247");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R247
            this.Site.CaptureRequirementIfAreNotEqual<ulong>(
                this.logonResponse.FolderIds[3],
                this.getReceiveFolderResponse.FolderId,
                247,
                @"[In RopSetReceiveFolder ROP] A client can change an existing Receive folder configuration for a message class by simply issuing this ROP with a different value in the FolderId field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R864");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R864
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                this.getReceiveFolderResponse.MessageClass[0],
                864,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If the FolderId field is set to zero, then the table row for the specified message class is deleted from the Receive folder table. (The details about the content of a table row are provided following this paragraph.)");
            #endregion capture
        }

        /// <summary>
        /// This test case verifies the requirements of RopGetStoreState ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC04_TestRopGetStoreState()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            // This operation is only valid when the Logon object refers to a private mailbox logon
            #region Step2: RopLogon ROP to logon to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            #region Step3: Create the active search folder
            uint searchFolderHandle = 0;
            ulong searchFolderId = 0;
            string searchFolderName = Common.GenerateResourceName(Site, "SearchFolder");
            this.CreateSearchFolder(searchFolderName, out searchFolderHandle, out searchFolderId);
            #endregion

            // This operation is only valid when the Logon object refers to a private mailbox logon
            #region Step4: RopLogon ROP to logon to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            if (Common.IsRequirementEnabled(1410001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1410001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1410001
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000000,
                    this.logonResponse.StoreState,
                    1410001,
                    @"Implementation set this field [StoreState in RopLogon ROP Success Response Buffer for Private Mailbox] to 0x00000000 [if the mailbox currently has any active search folders]. (Exchange 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(59781, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59781");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59781
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x01000000,
                    this.logonResponse.StoreState,
                    59781,
                    @"[In Appendix A: Product Behavior] If the mailbox currently has any active search folders, then the implementation sets this field [StoreState] to 0x01000000. (<7> Section 2.2.1.1.3: If the mailbox currently has any active search folders, then Exchange 2003 and Exchange 2007 set this field [StoreState] to 0x01000000.)");
            }
            #endregion

            #region Step5: RopGetStoreState ROP to get the store state
            this.oxcstorAdapter.DoRopCall(this.getStoreStateRequest, this.outObjHandle, ROPCommandType.RopGetStoreState, out this.outputBuffer);
            this.getStoreStateResponse = (RopGetStoreStateResponse)this.outputBuffer.RopsList[0];
            #endregion

            // Verify the RopGetStoreState is supported
            if (Common.IsRequirementEnabled(3126, this.Site))
            {
                #region Capture
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3126");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3126
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    3126,
                    @"[In Appendix A: Product Behavior] The implementation does implement the RopGetStoreState remote operation (ROP). (<10> Section 2.2.1.5: Exchange 2003 and Exchange 2007 does implement the RopGetStoreState remote operation (ROP) ([MS-OXCROPS] section 2.2.3.5))");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R339");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R339
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    339,
                    @"[In Receiving a RopGetStoreState ROP Request] The error code Success: Its value is 0x00000000.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R340");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R340
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    340,
                    @"[In Receiving a RopGetStoreState ROP Request] The error code Success: The operation succeeded.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R331");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R331
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    331,
                    @"[In RopGetStoreState ROP Request Buffer] The data to be retrieved is limited to the mailbox that is linked to the LogonID that is passed as part of the ROP request, as specified in [MS-OXCROPS] section 2.2.3.5.1.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R332");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R332
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    332,
                    @"[In RopGetStoreState ROP Request Buffer] This operation [RopGetStoreState] MUST be issued against a private mailbox logon.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R334");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R334
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x01000000,
                    this.logonResponse.StoreState,
                    334,
                    "[In RopGetStoreState ROP Success Response Buffer] StoreState: All other bits [all bits except STORE_HAS_SEARCHES ] MUST NOT be set.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3228");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3228
                Site.CaptureRequirementIfAreEqual<uint>(
                   0x01000000,
                   this.logonResponse.StoreState,
                   3228,
                   "[In RopGetStoreState ROP Success Response Buffer] StoreState: If the mailbox currently has any active search folders, this field [StoreState] MUST have the STORE_HAS_SEARCHES flag (0x01000000) set. ");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R4001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R4001
                this.Site.CaptureRequirementIfAreNotEqual<ulong>(
                    0,
                    searchFolderId,
                    4001,
                    @"[In RopGetStoreState ROP Success Response Buffer] StoreState: If the STORE_HAS_SEARCHES flag is set, the user has created one or more active searches in the message store.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R891");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R891
                Site.CaptureRequirementIfAreEqual<uint>(
                   0x01000000,
                   this.logonResponse.StoreState,
                   891,
                   "[In Receiving a RopGetStoreState ROP Request] The server MUST NOT set any other flags in the response [except STORE_HAS_SEARCHES flag].");


                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R889");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R889
                Site.CaptureRequirementIfAreEqual<uint>(
                   0x01000000,
                   this.logonResponse.StoreState,
                   889,
                   "[In Receiving a RopGetStoreState ROP Request] If the mailbox has any persisted search folders, then the server MUST set the STORE_HAS_SEARCHES flag in the response, as specified in section 2.2.1.5.2.");
                #endregion capture

                #region Step6: Disconnect
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
                #endregion

                #region Step7: Connect
                this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion

                #region Step8: RopLogon ROP to log on to the public folder
                // Log on to public folder
                Site.Log.Add(LogEntryKind.Comment, "Log on to public folder.");
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
                this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
                #endregion

                #region Step9: Call RopGetStoreState ROP against public folder logon
                this.oxcstorAdapter.DoRopCall(this.getStoreStateRequest, this.outObjHandle, ROPCommandType.RopGetStoreState, out this.outputBuffer);
                RopGetStoreStateResponse getStoreStateResponseForPublicFolder = (RopGetStoreStateResponse)this.outputBuffer.RopsList[0];

                #region Capture
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R341");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R341
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    getStoreStateResponseForPublicFolder.ReturnValue,
                    341,
                    @"[In Receiving a RopGetStoreState ROP Request] The error code NotSupported: Its value is 0x80040102.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R342");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R342
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    getStoreStateResponseForPublicFolder.ReturnValue,
                    342,
                    @"[In Receiving a RopGetStoreState ROP Request] The error code ecNotSupported: The ROP was not performed against a private mailbox logon.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R888");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R888
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    getStoreStateResponseForPublicFolder.ReturnValue,
                    888,
                    @"[In Receiving a RopGetStoreState ROP Request] If the operation is performed against a public folders logon, then the server MUST fail the operation with a ReturnValue of 0x80040102 (NotSupported).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R887");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R887
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x80040102,
                    getStoreStateResponseForPublicFolder.ReturnValue,
                    887,
                    @"[In Receiving a RopGetStoreState ROP Request] If the server implements this ROP [RopGetStoreState], it has the following behavior: The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
                #endregion capture
                #endregion
            }

            if (Common.IsRequirementEnabled(3125, this.Site))
            {
                #region Capture
                if (Common.IsRequirementEnabled(884, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R884");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R884
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80040FFF,
                        this.getStoreStateResponse.ReturnValue,
                        884,
                        @"[In Receiving a RopGetStoreState ROP Request] Implementation does return a value of 0x80040FFF (NotImplemented) in the ReturnValue field of the response. (Exchange 2010 and above follow this behavior.)");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1091");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1091
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80040FFF,
                        this.getStoreStateResponse.ReturnValue,
                        1091,
                        @"[In Receiving a RopGetStoreState ROP Request] The error code NotImplemented: Its value is 0x80040FFF.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1092");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1092
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80040FFF,
                        this.getStoreStateResponse.ReturnValue,
                        1092,
                        @"[In Receiving a RopGetStoreState ROP Request] The error code NotImplemented: The server does not implement this ROP.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3125");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3125
                this.Site.CaptureRequirementIfAreNotEqual<uint>(
                    0,
                    this.getStoreStateResponse.ReturnValue,
                    3125,
                    @"[In Appendix A: Product Behavior] The implementation does not implement the RopGetStoreState remote operation (ROP). (<10> Section 2.2.1.5: Exchange 2010  Exchange 2013 and Exchange 2016 does not implement the RopGetStoreState remote operation (ROP) ([MS-OXCROPS] section 2.2.3.5))");
                #endregion
            }

            #region Step10: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion

            #region Step11: Connect
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            // This operation is only valid when the Logon object refers to a private mailbox logon
            #region Step12: RopLogon ROP to logon to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements of RopGetPerUserLongTermIds ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC05_TestRopGetPerUserLongTermIds()
        {
            this.CheckTransportIsSupported();

            #region Variables
            string newPublicFolder = Common.GenerateResourceName(Site, "PublicFolderForGetPerUserLongTermId");
            LongTermId longTermIdForPublicFolder1ForServer;
            LongTermId longTermIdForPublicFolder2ForServer;
            #endregion

            #region Step1-3: Connect server1 with user configured by "AdminUserName" and create a new folder, then disconnect
            #region Step1: Connect to server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step2: RopLogon ROP to log on to the public folder and create a public folder
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            uint folderHandle = 0;
            ulong folderId = this.OpenFolder(newPublicFolder, FolderType.Genericfolder, this.logonResponse.FolderIds[1], out folderHandle);
            this.RopRelease(folderHandle);
            #endregion

            #region Step3: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion
            #endregion

            #region Step4: Connect server1 with the user configured by "UserNameOfMailboxOnServer1"
            this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PublicFolderServer, this.essdnOfUser1, this.domain, this.userNameOfMailboxOnServer1, this.passwordOfUserOfMailboxOnServer1);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step5: logon to a public folder;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Get the longTermId with REPLGUID of server1
            // Folder[7] is a default public folder in server1
            longTermIdForPublicFolder2ForServer = this.GetLongTermIdFromId(this.logonResponse.FolderIds[7]);

            // Send mail to the public folder, and set the mail as read
            folderId = this.OpenFolder(newPublicFolder, FolderType.Genericfolder, this.logonResponse.FolderIds[1], out folderHandle);
            uint targetMessageHandle;
            ulong messageId = this.CreateMessage(folderId, out targetMessageHandle);
            this.SetMessageAsRead(folderId, messageId);
            longTermIdForPublicFolder1ForServer = this.GetLongTermIdFromId(folderId);
            #endregion

            #region Step6: RopReadPerUserInformation ROP to get read/unread information
            this.readPerUserInformationRequest.FolderId = longTermIdForPublicFolder1ForServer;
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step7: RopLogon ROP to log on to a private mailbox with the user configured by "UserNameOfMailboxOnServer1"
            RopLogonRequest logonRequest = this.logonRequestForPrivateMailBox;
            logonRequest.LogonFlags = 0x01;
            logonRequest.OpenFlags = 0x01000000;
            logonRequest.Essdn = Encoding.ASCII.GetBytes(this.essdnOfUser1 + "\0");
            logonRequest.EssdnSize = (ushort)logonRequest.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(logonRequest, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer, this.userNameOfMailboxOnServer1);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step7

            #region step8: Call the RopWritePerUserInformation ROP to write read/unread information to private mailbox
            this.writePerUserInformationRequest.FolderId = longTermIdForPublicFolder1ForServer;
            this.writePerUserInformationRequest.HasFinished = 0x1;
            this.writePerUserInformationRequest.DataSize = this.readPerUserInformationResponse.DataSize;
            this.writePerUserInformationRequest.Data = this.readPerUserInformationResponse.Data;
            this.writePerUserInformationRequest.ReplGuid = longTermIdForPublicFolder1ForServer.DatabaseGuid;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step9: RopGetPerUserLongTermIds ROP with two valid DatabaseGuids of server1
            this.getPerUserLongTermIdsRequest.DatabaseGuid = longTermIdForPublicFolder2ForServer.DatabaseGuid;
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponse2 = (RopGetPerUserLongTermIdsResponse)this.outputBuffer.RopsList[0];

            #region Capture
            Site.Assert.AreEqual<uint>(
                0x00000000,
                getPerUserLongTermIdsResponse2.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.getPerUserLongTermIdsRequest.DatabaseGuid = longTermIdForPublicFolder1ForServer.DatabaseGuid;
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            this.getPerUserLongTermIdsResponse = (RopGetPerUserLongTermIdsResponse)this.outputBuffer.RopsList[0];
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponse1 = (RopGetPerUserLongTermIdsResponse)this.getPerUserLongTermIdsResponse;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.getPerUserLongTermIdsResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R480");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R480
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                this.getPerUserLongTermIdsResponse.ReturnValue,
                480,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R481");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R481
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x00000000,
                this.getPerUserLongTermIdsResponse.ReturnValue,
                481,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R957");

            bool isVerify_R957 = false;

            // If LongTermIds in getPerUserLongTermIdsResponse2 at step9 is null, and the DatabaseGuid in LongTermId in getPerUserLongTermIdsResponse1 is same to the DatabaseGuid in longTermIdForPublicFolder1ForServer (same to the REPLGUID in read/unread information) 
            // then indirectly indicates the server collect the associated public folder long-term ID
            for (int i = 0; i < this.getPerUserLongTermIdsResponse.LongTermIds.Length; i++)
            {
                // Verify the folder is a public folder
                if (this.ByteArrayEquals(getPerUserLongTermIdsResponse1.LongTermIds[i].DatabaseGuid, longTermIdForPublicFolder1ForServer.DatabaseGuid) &&
                    getPerUserLongTermIdsResponse2.LongTermIdCount == 0)
                {
                    isVerify_R957 = true;
                }
                else
                {
                    isVerify_R957 = false;
                    break;
                }
            }

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R957
            Site.CaptureRequirementIfIsTrue(
                isVerify_R957,
                957,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] For each entry in the table, the server collects the associated public folder LongTermID, as specified in [MS-OXCDATA] section 2.2.1.3.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R958");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R958
            Site.CaptureRequirementIfAreEqual<ushort>(
                this.getPerUserLongTermIdsResponse.LongTermIdCount,
                (ushort)this.getPerUserLongTermIdsResponse.LongTermIds.Length,
                958,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The total number of LongTermIDs collected is specified in the LongTermIdCount field and the aggregated list of LongTermIDs constitutes the value of the LongTermIds field.");
            #endregion capture
            #endregion

            #region Step10: RopGetPerUserLongTermIds with an invalid DatabaseGuid
            this.getPerUserLongTermIdsRequest.DatabaseGuid = new byte[16];
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponseForInvaidGuid;
            getPerUserLongTermIdsResponseForInvaidGuid = (RopGetPerUserLongTermIdsResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, getPerUserLongTermIdsResponseForInvaidGuid.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R956");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R956
            // Check the return values of the server by passing an invalid DatabaseGuid.
            bool isVerifiedR956 = this.getPerUserLongTermIdsResponse.LongTermIds != null && getPerUserLongTermIdsResponseForInvaidGuid.LongTermIdCount == 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR956,
                956,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The server searches the per-user data table of the mailbox for entries identified by the DatabaseGuid field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R960");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R960
            Site.CaptureRequirementIfAreEqual<ushort>(
                0,
                getPerUserLongTermIdsResponseForInvaidGuid.LongTermIdCount,
                960,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The server can return an empty list.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R473");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R473
            this.Site.CaptureRequirementIfAreEqual<ushort>(
                0,
                getPerUserLongTermIdsResponseForInvaidGuid.LongTermIdCount,
                473,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIdCount: This field [LongTermIdCount] can be set to zero.");
            #endregion capture
            #endregion Step10

            #region Step11: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion Step11

            #region Step12: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step12

            #region Step13: RopLogon ROP to log on to public folder
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step13

            #region Step14: Delete Folder
            // Open folder to be deleted
            ulong openedFolderId = this.OpenFolder(newPublicFolder, FolderType.Genericfolder, this.logonResponse.FolderIds[1], out folderHandle);
            this.DeleteFolder(openedFolderId, folderHandle);
            #endregion Step14
        }

        /// <summary>
        /// This test case verifies converting between LongTermIDs and ShortTermIDs.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC06_TestRopLongTermIdFromIdAndRopIdFromLongTermId()
        {
            this.CheckTransportIsSupported();

            #region Step 1: Connect the server via user configured by "AdminUserName".
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 2: Logon to a private mailbox.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 3: Call RopLongTermIdFromId ROP to get the LongTermId of the inbox.

            this.longTermIdFromIdRequest.ObjectId = this.logonResponse.FolderIds[4];
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            LongTermId longTermIdForInbox = this.longTermIdFromIdResponse.LongTermId;

            #region Capture MS-OXCSTOR_R449, MS-OXCSTOR_R1243, MS-OXCSTOR_R450 and MS-OXCSTOR_R944.

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R449");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R449
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                449,
                @"[In Receiving a RopLongTermIdFromId ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1243");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1243.
            // The ROP RopLongTermIdFromId performs successfully, this requirement can be verified directly.
            Site.CaptureRequirement(
                1243,
                @"[In Receiving a RopLongTermIdFromId ROP Request] After obtaining the REPLGUID from the REPLID and REPLGUID to-and-from mapping table, the server uses the REPLGUID to construct the 192-bit LongTermID, as specified in [MS-OXCDATA] section 2.2.1.3.1, which is returned in the LongTermId field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R450");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R450
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                450,
                @"[In Receiving a RopLongTermIdFromId ROP Request] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R944");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R944.
            // The ROP RopLongTermIdFromId performs successfully, MS-OXCSTOR_R944 can be verified directly.
            Site.CaptureRequirement(
                944,
                @"[In Receiving a RopLongTermIdFromId ROP Request] The server searches the REPLID and REPLGUID to-and-from mapping table for the replica ID (REPLID) portion of the given Folder ID or Message ID, as specified in [MS-OXCDATA] sections 2.2.1.1 or 2.2.1.2.");
            #endregion
            #endregion

            #region Step 4: Call RopIdFromLongTermId by the LongTermID get in step 3.

            this.getIdFromLongTermIdRequest.LongTermId.DatabaseGuid = this.longTermIdFromIdResponse.LongTermId.DatabaseGuid;
            this.getIdFromLongTermIdRequest.LongTermId.GlobalCounter = this.longTermIdFromIdResponse.LongTermId.GlobalCounter;
            this.oxcstorAdapter.DoRopCall(this.getIdFromLongTermIdRequest, this.outObjHandle, ROPCommandType.RopIdFromLongTermId, out this.outputBuffer);
            this.getIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1252");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1252.
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.logonResponse.FolderIds[4],
                this.getIdFromLongTermIdResponse.ObjectId,
                1252,
                @"[In Receiving a RopIdFromLongTermId ROP Request] After obtaining the REPLID from the REPLID and REPLGUID to-and-from mapping table, the server uses the REPLID to construct the 64-bit Folder ID or Message ID, as specified in [MS-OXCDATA] sections 2.2.1.1 or 2.2.1.2, which is returned in the ObjectId field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59718");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59718.
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.logonResponse.FolderIds[4],
                this.getIdFromLongTermIdResponse.ObjectId,
                59718,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] That [LongTermId] is mapped from the given Folder ID as specified in[MS-OXCDATA] sections 2.2.4.1 or 2.2.4.2, that is contained in the ObjectID field (section 2.2.1.8.1).");

            #region Capture MS-OXCSTOR_R466, MS-OXCSTOR_R467, MS-OXCSTOR_R454 and MS-OXCSTOR_R952.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R466");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R466.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getIdFromLongTermIdResponse.ReturnValue,
                466,
                @"[In Receiving a RopIdFromLongTermId ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R467");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R467.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getIdFromLongTermIdResponse.ReturnValue,
                467,
                @"[In Receiving a RopIdFromLongTermId ROP Request] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R454");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R454.
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.logonResponse.FolderIds[4],
                this.getIdFromLongTermIdResponse.ObjectId,
                454,
                @"[In RopIdFromLongTermId ROP Request Buffer] LongTermId: Contains the LongTermID, as specified in [MS-OXCDATA] section 2.2.1.3.1, to be mapped to the Folder ID, as specified in [MS-OXCDATA] sections 2.2.1.1.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R952");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R952.
            Site.CaptureRequirementIfAreEqual<ulong>(
                this.logonResponse.FolderIds[4],
                this.getIdFromLongTermIdResponse.ObjectId,
                952,
                @"[In Receiving a RopIdFromLongTermId ROP Request] The server MUST map the same REPLGUID to the same REPLID every time it is queried.");
            #endregion
            #endregion

            #region Step 5: Call RopLongTermIdFromId with the folder id get in step 4.

            this.longTermIdFromIdRequest.ObjectId = this.getIdFromLongTermIdResponse.ObjectId;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            RopLongTermIdFromIdResponse longTermIdFromIdResponse1 = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information.
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R947");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R947.
            Site.CaptureRequirementIfIsTrue(
                longTermIdFromIdResponse1.LongTermId.DatabaseGuid.SequenceEqual(longTermIdForInbox.DatabaseGuid),
                947,
                @"[In Receiving a RopLongTermIdFromId ROP Request] The server MUST map the same REPLID to the same replica GUID (REPLGUID) every time it is queried.");
            #endregion

            #region Step 6: Call RopIdFromLongTermId with a non-existent LongTermId.

            // Create a long term id that does not exist.
            LongTermId notExistlongTermId;
            notExistlongTermId.DatabaseGuid = new byte[longTermIdFromIdResponse.LongTermId.DatabaseGuid.Length];
            notExistlongTermId.GlobalCounter = new byte[longTermIdFromIdResponse.LongTermId.GlobalCounter.Length];
            Array.Copy(this.longTermIdFromIdResponse.LongTermId.DatabaseGuid, notExistlongTermId.DatabaseGuid, notExistlongTermId.DatabaseGuid.Length);
            Array.Copy(this.longTermIdFromIdResponse.LongTermId.GlobalCounter, notExistlongTermId.GlobalCounter, notExistlongTermId.GlobalCounter.Length);
            notExistlongTermId.DatabaseGuid[0] = 0xFF;
            notExistlongTermId.GlobalCounter[0] = 0xFF;
            this.getIdFromLongTermIdRequest.LongTermId.DatabaseGuid = notExistlongTermId.DatabaseGuid;
            this.getIdFromLongTermIdRequest.LongTermId.GlobalCounter = notExistlongTermId.GlobalCounter;
            this.oxcstorAdapter.DoRopCall(this.getIdFromLongTermIdRequest, this.outObjHandle, ROPCommandType.RopIdFromLongTermId, out this.outputBuffer);
            RopIdFromLongTermIdResponse fromLongTermIdResponse1 = (RopIdFromLongTermIdResponse)this.outputBuffer.RopsList[0];
            #endregion

            #region Step 7: Call RopLongTermIdFromId with the invalid object ID get from step 6.

            this.longTermIdFromIdRequest.ObjectId = fromLongTermIdResponse1.ObjectId;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.longTermIdFromIdResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture MS-OXCSTOR_R950 and MS-OXCSTOR_R1246.
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R950");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R950.
            // Create a nonexistent LongTermId, so the row must be found. 
            // If succeed to get the LongTermId by using Id, 
            // it indicates the row that contains the REPLGUID and a newly assigned REPLID value is added into the table.
            Site.CaptureRequirementIfIsTrue(
                notExistlongTermId.DatabaseGuid.SequenceEqual(this.longTermIdFromIdResponse.LongTermId.DatabaseGuid) && notExistlongTermId.GlobalCounter.SequenceEqual(this.longTermIdFromIdResponse.LongTermId.GlobalCounter),
                950,
                @"[In Receiving a RopIdFromLongTermId ROP Request] If the REPLGUID is not found, the server adds a new entry consisting of the REPLGUID portion of the LongTermID.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1246");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1246
            // Create a nonexistent LongTermId, so the row must be found. 
            // If succeed to get the LongTermId by using Id, 
            // it indicates the row that contains the REPLGUID and a newly assigned REPLID value is added into the table.
            Site.CaptureRequirementIfIsTrue(
                notExistlongTermId.DatabaseGuid.SequenceEqual(this.longTermIdFromIdResponse.LongTermId.DatabaseGuid) && notExistlongTermId.GlobalCounter.SequenceEqual(this.longTermIdFromIdResponse.LongTermId.GlobalCounter),
                1246,
                @"[In Receiving a RopIdFromLongTermId ROP Request] If the REPLGUID is not found, the server adds a newly assigned replica ID (REPLID).");

            #endregion
            #endregion

            #region Step 8: Call RopLongTermIdFromId with setting object ID to zero.

            this.longTermIdFromIdRequest.ObjectId = 0;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];

            if (Common.IsRequirementEnabled(1337, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1337");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1337.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.longTermIdFromIdResponse.ReturnValue,
                    1337,
                    @"[In Appendix A: Product Behavior] Implementation does not fail the operation [RopLongTermIdFromId] with 0x8004010F, but ecNone. <43> Section 3.2.5.8: If the ObjectId field is set to zero, Exchange 2013 and Exchange 2016 returns ecNone.");
            }

            if (Common.IsRequirementEnabled(1355, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1355");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1355.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x8004010F,
                    this.longTermIdFromIdResponse.ReturnValue,
                    1355,
                    @"[In Appendix A: Product Behavior] Implementation does fail the operation with 0x8004010F if the REPLID is not in the REPLID and REPLGUID to-and-from mapping table [when the server receiving a RopLongTermIdFromId ROP Request]. (Exchange 2010 and below follow this behavior.)");
            }
            #endregion

            #region Step 9: Create a message in the inbox.
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;
            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;
            openFolderRequest.FolderId = this.logonResponse.FolderIds[4];
            openFolderRequest.OpenModeFlags = 0x0;
            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[openFolderResponse.OutputHandleIndex];

            uint messageHandle = 0;
            ulong messageId = 0;
            messageId = this.CreateMessage(this.logonResponse.FolderIds[4], out messageHandle);

            this.longTermIdFromIdRequest.ObjectId = messageId;
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.getIdFromLongTermIdRequest.LongTermId.DatabaseGuid = this.longTermIdFromIdResponse.LongTermId.DatabaseGuid;
            this.getIdFromLongTermIdRequest.LongTermId.GlobalCounter = this.longTermIdFromIdResponse.LongTermId.GlobalCounter;
            this.oxcstorAdapter.DoRopCall(this.getIdFromLongTermIdRequest, this.outObjHandle, ROPCommandType.RopIdFromLongTermId, out this.outputBuffer);
            this.getIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59799");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59799.
            Site.CaptureRequirementIfAreEqual<ulong>(
                messageId,
                this.getIdFromLongTermIdResponse.ObjectId,
                59799,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] That  [LongTermId] is mapped from the MessageID as specified in[MS-OXCDATA] sections 2.2.4.1 or 2.2.4.2, that is contained in the ObjectID field (section 2.2.1.8.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59724");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59724.
            Site.CaptureRequirementIfAreEqual<ulong>(
                messageId,
                this.getIdFromLongTermIdResponse.ObjectId,
                59724,
                @"[In RopIdFromLongTermId ROP Request Buffer] LongTermId: Contains the LongTermID, as specified in [MS-OXCDATA] section 2.2.1.3.1, to be mapped to the Message ID, as specified in [MS-OXCDATA] sections 2.2.1.2.");

            // Delete the created message
            RopHardDeleteMessagesRequest deleteMessageRequest = new RopHardDeleteMessagesRequest
            {
                RopId = (byte)RopId.RopHardDeleteMessages,
                LogonId = 0x00,
                InputHandleIndex = 0x00,
                NotifyNonRead = 0x00,
                MessageIds = new ulong[1]
            };
            deleteMessageRequest.MessageIds[0] = messageId;
            deleteMessageRequest.MessageIdCount = 1;
            this.oxcstorAdapter.DoRopCall(deleteMessageRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);
            #endregion

            #region Step 10: Disconnect.
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion

            #region Step 11: Connect the server via user configured by "AdminUserName".
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 12: Logon to the private mailbox.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            RopLogonResponse logonResponse2 = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                logonResponse2.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 13: Call RopLongTermIdFromId with successful response.

            this.longTermIdFromIdRequest.ObjectId = logonResponse2.FolderIds[4];
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            LongTermId longTermIdForInbox2 = this.longTermIdFromIdResponse.LongTermId;

            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1242");

            bool isVerifyR1242 = longTermIdForInbox2.DatabaseGuid.SequenceEqual(longTermIdForInbox.DatabaseGuid);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1242.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1242,
                1242,
                @"[In Receiving a RopLongTermIdFromId ROP Request] But each server MUST map any particular REPLID to the same REPLGUID every time it is queried");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1251");

            bool isVerifyR1251 = longTermIdForInbox2.DatabaseGuid.SequenceEqual(longTermIdForInbox.DatabaseGuid);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1251.
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1251,
                1251,
                @"[In Receiving a RopIdFromLongTermId ROP Request] [Other servers can map a particular REPLGUID to a different REPLID,] But each server MUST map any particular REPLGUID to the same REPLID every time it is queried.");

            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements of RopGetPerUserGuid.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC07_TestRopGetPerUserGuid()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon ROP to log on to public folder.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step2

            #region Step3: RopLongTermIdFromId ROP
            this.longTermIdFromIdRequest.ObjectId = this.logonResponse.FolderIds[3];
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            LongTermId longTermId = this.longTermIdFromIdResponse.LongTermId;
            #endregion Step3

            #region Step4: RopGetPerUserGuid ROP against a public folder logon
            this.getPerUserGuidRequest.LongTermId = longTermId;
            this.oxcstorAdapter.DoRopCall(this.getPerUserGuidRequest, this.insideObjHandle, ROPCommandType.RopGetPerUserGuid, out this.outputBuffer);
            RopGetPerUserGuidResponse getPerUserGuidResponse1 = (RopGetPerUserGuidResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R498");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R498
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserGuidResponse1.ReturnValue,
                498,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNotSupported: Its value is 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R499");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R499
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserGuidResponse1.ReturnValue,
                499,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNotSupported: ROP was attempted against a public folders logon.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R962");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R962
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserGuidResponse1.ReturnValue,
                962,
                @"[In Receiving a RopGetPerUserGuid ROP Request] If the operation [RopGetPerUserGuid] is performed against a public folders logon, the server MUST fail the operation with a ReturnValue of 0x80040102.");
            #endregion capture
            #endregion Step4

            #region Step5: RopLogon ROP to log on to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step5

            #region Step6: Call RopWritePerUserInformation ROP to write some data into sever. Then the RopGetPerUserGuid can be returned successfully
            byte[] data = { 0xd8, 0x44, 0xae, 0x73, 0xf9, 0x61, 0x5d, 0x4f, 0xb3, 0xc6, 0x9a, 0x7c, 0x31, 0xfe, 0xc1, 0x23, 0x06, 0x00, 0x00, 0x00, 0x78, 0x2b, 0x33, 0x00 }; // Sample data
            this.writePerUserInformationRequest.FolderId = longTermId; // Pass longTermId of the public folder
            this.writePerUserInformationRequest.HasFinished = 0x1;
            this.writePerUserInformationRequest.DataSize = (ushort)data.Length;
            this.writePerUserInformationRequest.Data = data;
            this.writePerUserInformationRequest.ReplGuid = this.logonResponse.ReplGuid;
            this.oxcstorAdapter.DoRopCall(this.writePerUserInformationRequest, this.outObjHandle, ROPCommandType.RopWritePerUserInformation, out this.outputBuffer);
            this.writePerUserInformationResponse = (RopWritePerUserInformationResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.writePerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step6

            #region Step7: Call RopReadPerUserInformation ROP to make the RopGetPerUserGuid ROP calling successful
            this.readPerUserInformationRequest.FolderId = longTermId;
            this.readPerUserInformationRequest.DataOffset = 0x00; // Set to a random number, to be changed
            this.readPerUserInformationRequest.MaxDataSize = 0x30; // Maximum amount of data to be returned
            this.oxcstorAdapter.DoRopCall(this.readPerUserInformationRequest, this.outObjHandle, ROPCommandType.RopReadPerUserInformation, out this.outputBuffer);
            this.readPerUserInformationResponse = (RopReadPerUserInformationResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0,
                this.readPerUserInformationResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step7

            #region Step8: RopGetPerUserGuid ROP
            this.oxcstorAdapter.DoRopCall(this.getPerUserGuidRequest, this.outObjHandle, ROPCommandType.RopGetPerUserGuid, out this.outputBuffer);
            this.getPerUserGuidResponse = (RopGetPerUserGuidResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R496");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R496
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getPerUserGuidResponse.ReturnValue,
                496,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R497");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R497
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getPerUserGuidResponse.ReturnValue,
                497,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R964");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R964
            Site.CaptureRequirementIfAreEqual<string>(
                this.logonResponse.ReplGuid.ToString(),
                this.getPerUserGuidResponse.DatabaseGuid.ToString(),
                964,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The server returns the associated REPLGUID value in the DatabaseGuid field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R961");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R961
            // If after logging on to the private mailbox the RopGetPerUserGuid ROP is successful 
            // and after logging on to public folder the ROP calling is failed. 
            // It indicates the server will verify the operation is being performed against a private mailbox logon, 
            // and not a public folder logon. 
            bool isVerifiedR961 = this.getPerUserGuidResponse.ReturnValue == 0 && getPerUserGuidResponse1.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR961,
                961,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
            #endregion capture
            #endregion Step8

            #region Step9: RopGetPerUserGuid ROP with an invalid LongTermId
            // Create an invalid LongTermId
            longTermId.DatabaseGuid[0] = 0xFF;
            longTermId.GlobalCounter[0] = 0xFF;
            this.getPerUserGuidRequest.LongTermId = longTermId;
            this.oxcstorAdapter.DoRopCall(this.getPerUserGuidRequest, this.outObjHandle, ROPCommandType.RopGetPerUserGuid, out this.outputBuffer);
            getPerUserGuidResponse1 = (RopGetPerUserGuidResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.getPerUserGuidResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R500");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R500
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                getPerUserGuidResponse1.ReturnValue,
                500,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNotFound: Its value is 0x8004010F.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1135");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1135
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                getPerUserGuidResponse1.ReturnValue,
                1135,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The error code ecNotFound: The public folder identified by the value of the LongTermID field could not be found");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1254");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1254
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                getPerUserGuidResponse1.ReturnValue,
                1254,
                @"[In Receiving a RopGetPerUserGuid ROP Request] If the public folder specified by the LongTermId field cannot be found, the server MUST fail the operation with 0x8004010F (ecNotFound) in the ReturnValue field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R963");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R963
            // Server return failure if pass a non-existed LongTermId to RopGetPerUserGuid;
            // Server return success if pass a valid LongTermId to RopGetPerUserGuid. 
            // This indicates that the server will search row in the table according to LongTermId.
            bool isVerifiedR963 = this.getPerUserGuidResponse.ReturnValue == 0
                && getPerUserGuidResponse1.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifiedR963,
                963,
                @"[In Receiving a RopGetPerUserGuid ROP Request] The server searches the per-user data table for the mailbox for the only row with an FID that is associated with the public folder specified by the LongTermId field in the request.");
            #endregion capture
            #endregion Step9
        }

        /// <summary>
        /// This test case verifies the requirements of logging on to the private mailbox to test that public folder is a ghosted request.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC08_TestRopPublicFolderIsGhosted()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon to log on to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step2

            #region Step3: RopPublicFolderIsGhosted ROP
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[0];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            this.publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R917");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R917
            // 0 means false
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                this.publicFolderIsGhostedResponse.IsGhosted,
                917,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] If the operation is issued against a private mailbox message store, the server MUST return FALSE in the IsGhosted field of the response.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1237");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1237
            // That IsGhosted is false, and Servers is null means no replication state data is returned.
            this.Site.CaptureRequirementIfIsNull(
                this.publicFolderIsGhostedResponse.Servers,
                1237,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] In this case [if the operation is issued against a private mailbox store, the server MUST return FALSE in the IsGhosted field of the response], no replication state data is returned.");
            #endregion capture
            #endregion Step3

            #region Step4: Disconnect server
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            #endregion

            #region Step5: Connect the server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connect successfully");
            #endregion

            #region Step6: RopLogon ROP to log on to the public folder
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step7: RopPublicFolderIsGhosted ROP
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[0];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            this.publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];

            #region Capture
            if (Common.IsRequirementEnabled(400001, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R400001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R400001
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.publicFolderIsGhostedResponse.ReturnValue,
                    400001,
                    @"[In RopPublicFolderIsGhosted ROP Request Buffer] Implementation does only issue this operation [RopPublicFolderIsGhosted ] against a public folders logon. (Exchange 2007 and above follow this behavior.)");
            }

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R431");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R431
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.publicFolderIsGhostedResponse.ReturnValue,
                431,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R432");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R432
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.publicFolderIsGhostedResponse.ReturnValue,
                432,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The error code ecNone: Success.");
            #endregion capture
            #endregion

            #region Step8: RopPublicFolderIsGhosted ROP with invalid FID
            // Set an invalid FID
            this.publicFolderIsGhostedRequest.FolderId = 0xFFFFFFFF;
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            this.publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R43401");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R43401
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                this.publicFolderIsGhostedResponse.ReturnValue,
                43401,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The error code ecNotFound: Its value is 0x8004010F.");
            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements of read-only properties on the private mailbox logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC09_TestOperationsOnReadOnlyProperties()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step2: RopLogon ROP log on to private mailbox.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            #region Step3: Test PidTagExtendedRuleSizeLimit

            #region Variables
            PropertyTag propertyTagExtendedRuleSize;
            uint resultExtendedRuleSize;
            TaggedPropertyValue taggedExtendedRuleSize = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagExtendedRuleSizeLimit's id is 0x0E9B
            const ushort PidTagExtendedRuleSizeLimit = 0x0E9B;
            #endregion

            #region Read a readable property PidTagExtendedRuleSizeLimit
            propertyTagExtendedRuleSize.PropertyId = PidTagExtendedRuleSizeLimit;

            // According to [MS-OXPROPS], PidTagExtendedRuleSizeLimit's data type is 0x0003
            propertyTagExtendedRuleSize.PropertyType = 0x0003;
            uint responseFlag;
            byte[] propertyValue;
            resultExtendedRuleSize = this.GetPropertyValue(propertyTagExtendedRuleSize, out responseFlag, out propertyValue);
            int originalPidTagExtendedRuleSizeLimit = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region RopSetProperties ROP to write a non-writable property PidTagExtendedRuleSizeLimit
            taggedExtendedRuleSize.PropertyTag = propertyTagExtendedRuleSize;
            int changedPidTagExtendedRuleSizeLimit = originalPidTagExtendedRuleSizeLimit + 1;
            taggedExtendedRuleSize.Value = BitConverter.GetBytes(changedPidTagExtendedRuleSizeLimit);
            resultExtendedRuleSize = this.TrySetLogonProperty(taggedExtendedRuleSize);
            #endregion

            #region Read a readable property PidTagExtendedRuleSizeLimit
            resultExtendedRuleSize = this.GetPropertyValue(propertyTagExtendedRuleSize, out responseFlag, out propertyValue);
            int currentPidTagExtendedRuleSizeLimit = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region Capture R3056
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3056");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3056
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlag,
                3056,
                @"[In PidTagExtendedRuleSizeLimit Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");

            #endregion

            #region Capture R3055001

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055001");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055001
            // If the value was not changed by the write operation, the property is read-only.
            Site.CaptureRequirementIfAreNotEqual<int>(
                changedPidTagExtendedRuleSizeLimit,
                currentPidTagExtendedRuleSizeLimit,
                3055001,
                @"[In Read-Only Properties] The PidTagExtendedRuleSizeLimit property is read-only.");
            #endregion

            #endregion

            #region Step4: Test PidTagMaximumSubmitMessageSize

            #region Variables
            PropertyTag propertyTagMaximumSubmitMessageSize;

            // According to [MS-OXPROPS],PidTagMaximumSubmitMessageSize's id is 0x666D
            const ushort PidTagMaximumSubmitMessageSize = 0x666D;
            #endregion

            #region RopGetPropertiesSpecific ROP to read a readable property PidTagMaximumSubmitMessageSize
            propertyTagMaximumSubmitMessageSize.PropertyId = PidTagMaximumSubmitMessageSize;

            // According to [MS-OXPROPS], PidTagMaximumSubmitMessageSize's data type is 0x0003
            propertyTagMaximumSubmitMessageSize.PropertyType = 0x0003;
            uint responseFlagMaximumSubmintMessageSize;
            this.TryGetFlagLogonProperty(propertyTagMaximumSubmitMessageSize, out responseFlagMaximumSubmintMessageSize);
            #endregion

            #region Capture R3057
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3057");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3057
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagMaximumSubmintMessageSize,
                3057,
                @"[In PidTagMaximumSubmitMessageSize Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            #endregion
            #endregion

            #region Step5: Test PidTagProhibitReceiveQuota

            #region Variables
            PropertyTag propertyTagProhibitReceiveQuota;

            // According to [MS-OXPROPS], PidTagProhibitReceiveQuota's id is 0x666A
            const ushort PidTagProhibitReceiveQuota = 0x666A;
            #endregion

            #region RopGetPropertiesSpecific ROP to read a readable property PidTagProhibitReceiveQuota
            propertyTagProhibitReceiveQuota.PropertyId = PidTagProhibitReceiveQuota;

            // According to [MS-OXPROPS], PidTagProhibitReceiveQuota's data type is 0x0003
            propertyTagProhibitReceiveQuota.PropertyType = 0x0003;
            uint responseFlagProhibitReceiveQuota;
            this.TryGetFlagLogonProperty(propertyTagProhibitReceiveQuota, out responseFlagProhibitReceiveQuota);
            #endregion

            #region Capture R3058

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3058");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3058
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagProhibitReceiveQuota,
                3058,
                @"[In PidTagProhibitReceiveQuota Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #endregion

            #region Step6: Test PidTagProhibitSendQuota

            #region Variables
            PropertyTag propertyTagProhibitSendQuota;

            // According to [MS-OXPROPS], PidTagProhibitSendQuota's id is 0x666E
            const ushort PidTagProhibitSendQuota = 0x666E;
            #endregion

            #region Read a readable property PidTagProhibitSendQuota
            propertyTagProhibitSendQuota.PropertyId = PidTagProhibitSendQuota;

            // According to [MS-OXPROPS], PidTagProhibitSendQuota's data type is 0x0003
            propertyTagProhibitSendQuota.PropertyType = 0x0003;
            uint responseFlagProhibitSendQuota;
            this.TryGetFlagLogonProperty(propertyTagProhibitSendQuota, out responseFlagProhibitSendQuota);
            #endregion

            #region Capture R3059

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3059");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3059
            // The value responseFlagProhibitSendQuota is 0 indicates the ROPGetPropertiesSpecific is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagProhibitSendQuota,
                3059,
                @"[In PidTagProhibitSendQuota Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #endregion

            #region Step7: Test PidTagContentCount

            #region Variables
            PropertyTag propertyContentCount;
            uint resultContentCount;
            TaggedPropertyValue taggedContentCount = new TaggedPropertyValue();

            // According to the MS-OXPROPS, PidTagContentCount's id is 0x3602
            const ushort PidTagContentCount = 0x3602;
            #endregion

            #region Read a readable property PidTagContentCount
            propertyContentCount.PropertyId = PidTagContentCount;

            // According to [MS-OXPROPS], PidTagContentCount's data type is 0x0003
            propertyContentCount.PropertyType = 0x0003;
            uint responseContentCount;
            resultContentCount = this.GetPropertyValue(propertyContentCount, out responseContentCount, out propertyValue);
            int originalPidTagContentCount = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region Capture R3061

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3061");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3061
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseContentCount,
                3061,
                @"[In PidTagContentCount Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region Write a non-writable property PidTagContentCount
            taggedContentCount.PropertyTag = propertyContentCount;
            int changedPidTagContentCount = originalPidTagContentCount + 1;
            taggedContentCount.Value = BitConverter.GetBytes(changedPidTagContentCount);
            resultContentCount = this.TrySetLogonProperty(taggedContentCount);

            resultContentCount = this.GetPropertyValue(propertyContentCount, out responseFlag, out propertyValue);
            int currentPidTagContentCount = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region Capture R3055006

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055006");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055006
            // If the value was not changed by the write operation, the property is read-only.
            Site.CaptureRequirementIfAreNotEqual<int>(
                changedPidTagContentCount,
                currentPidTagContentCount,
                3055006,
                @"[In Read-Only Properties] The PidTagContentCount property is read-only.");

            #endregion

            #region Delete a non-deletable property PidTagContentCount
            resultContentCount = this.TryDeleteLogonProperty(propertyContentCount);
            #endregion

            #endregion

            #region Step8: Test PidTagMailboxOwnerEntryId

            #region Variables
            PropertyTag propertyTagMailboxOwnerEntryId;
            uint resultMailboxOwnerEntryId;
            TaggedPropertyValue taggedMailboxOwnerEntryId = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagMailboxOwnerEntryId's id is 0x661B
            const ushort PidTagMailboxOwnerEntryId = 0x661B;
            #endregion

            #region RopGetPropertiesSpecific to read a readable property PidTagMailboxOwnerEntryId
            propertyTagMailboxOwnerEntryId.PropertyId = PidTagMailboxOwnerEntryId;

            // According to [MS-OXPROPS], PidTagMailboxOwnerEntryId's data type is 0x0102
            propertyTagMailboxOwnerEntryId.PropertyType = 0x0102;
            uint responseFlagTagMailBoxOwnerEntryID;
            resultMailboxOwnerEntryId = this.GetPropertyValue(propertyTagMailboxOwnerEntryId, out responseFlagTagMailBoxOwnerEntryID, out propertyValue);
            byte[] originalPidTagMailboxOwnerEntryId = new byte[propertyValue.Length];
            Array.Copy(propertyValue, originalPidTagMailboxOwnerEntryId, propertyValue.Length);
            #endregion

            #region Capture R3062

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3062");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3062
            // The responseFlagTagMailBoxOwnerEntryID value is 0 indicates the RopGetPropertiesSpecific is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagTagMailBoxOwnerEntryID,
                3062,
                @"[In PidTagMailboxOwnerEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region ROP RopSetProperties to write a non-writable property PidTagMailboxOwnerEntryId
            taggedMailboxOwnerEntryId.PropertyTag = propertyTagMailboxOwnerEntryId;
            byte[] mailboxOwnerEntryIdValue = new byte[2] { 0x00, 0x00 };
            taggedMailboxOwnerEntryId.Value = mailboxOwnerEntryIdValue;
            resultMailboxOwnerEntryId = this.TrySetLogonProperty(taggedMailboxOwnerEntryId);

            resultMailboxOwnerEntryId = this.GetPropertyValue(propertyTagMailboxOwnerEntryId, out responseFlagTagMailBoxOwnerEntryID, out propertyValue);
            byte[] currentPidTagMailboxOwnerEntryId = propertyValue;
            #endregion

            #region Capture R3055007
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055007");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055007
            Site.CaptureRequirementIfIsFalse(
                Common.CompareByteArray(currentPidTagMailboxOwnerEntryId, taggedMailboxOwnerEntryId.Value),
                3055007,
                @"[In Read-Only Properties] The PidTagMailboxOwnerEntryId property is read-only.");
            #endregion
            #endregion

            #region Step9: Test PidTagMailboxOwnerName

            #region Variables
            PropertyTag propertyTagMailboxOwnerName;
            uint resultMailboxOwnerName;
            TaggedPropertyValue taggedMailboxOwnerName = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagMailboxOwnerName's id is 0x661C
            const ushort PidTagMailboxOwnerName = 0x661C;
            #endregion

            #region RopGetPropertiesSpecific ROP to read a readable property PidTagMailboxOwnerName
            propertyTagMailboxOwnerName.PropertyId = PidTagMailboxOwnerName;

            // According to [MS-OXPROPS], PidTagMailboxOwnerName's data type is 0x001F
            propertyTagMailboxOwnerName.PropertyType = 0x001F;
            uint responseMailboxOwerName;

            resultMailboxOwnerName = this.GetPropertyValue(propertyTagMailboxOwnerName, out responseMailboxOwerName, out propertyValue);
            byte[] originalPidTagMailboxOwnerName = new byte[propertyValue.Length];
            Array.Copy(propertyValue, originalPidTagMailboxOwnerName, propertyValue.Length);
            #endregion

            #region Capture R3063

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3063");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3063
            // When responseMailboxOwerName value is 0, it indicates RopGetPropertiesSpecifics is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseMailboxOwerName,
                3063,
                @"[In PidTagMailboxOwnerName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region Write a non-writable property PidTagMailboxOwnerName
            taggedMailboxOwnerName.PropertyTag = propertyTagMailboxOwnerName;
            taggedMailboxOwnerName.Value = Encoding.Unicode.GetBytes("50\0");
            resultMailboxOwnerName = this.TrySetLogonProperty(taggedMailboxOwnerName);

            resultMailboxOwnerName = this.GetPropertyValue(propertyTagMailboxOwnerName, out responseMailboxOwerName, out propertyValue);
            byte[] currentPidTagMailboxOwnerName = propertyValue;
            #endregion

            #region Capture R3055008
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055008");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055008
            Site.CaptureRequirementIfIsFalse(
                Common.CompareByteArray(taggedMailboxOwnerName.Value, currentPidTagMailboxOwnerName),
                3055008,
                @"[In Read-Only Properties] The PidTagMailboxOwnerName property is read-only.");

            #endregion
            #endregion

            #region  Step10: Test PidTagMessageSize

            #region Variables
            PropertyTag propertyTagMessageSize;
            uint resultMessageSize;
            TaggedPropertyValue taggedMessageSize = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagMessageSize's id is 0x0E08
            const ushort PidTagMessageSize = 0x0E08;
            #endregion

            #region Read a readable property PidTagMessageSize
            propertyTagMessageSize.PropertyId = PidTagMessageSize;

            // According to [MS-OXPROPS], PidTagMessageSize's data type is 0x0003
            propertyTagMessageSize.PropertyType = 0x0003;
            uint responseFlagTagMessageSize;
            resultMessageSize = this.GetPropertyValue(propertyTagMessageSize, out responseFlagTagMessageSize, out propertyValue);
            int originalPidTagMessageSize = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region Capture R3064

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3064");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3064
            // ResponseFlagTagMessageSize is 0 indicates that the RopGetPropertiesSpecific is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagTagMessageSize,
                3064,
                @"[In PidTagMessageSize Property] Type: PtypInteger32 ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region RopSetProperties ROP to write a non-writable property PidTagMessageSize
            taggedMessageSize.PropertyTag = propertyTagMessageSize;
            int changedPidTagMessageSize = originalPidTagMessageSize + 1;
            taggedMessageSize.Value = BitConverter.GetBytes(changedPidTagMessageSize);
            resultMessageSize = this.TrySetLogonProperty(taggedMessageSize);

            resultMessageSize = this.GetPropertyValue(propertyTagMessageSize, out responseFlagTagMessageSize, out propertyValue);
            int currentPidTagMessageSize = BitConverter.ToInt32(propertyValue, 0);
            #endregion

            #region Capture R3055009

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055009");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055009
            Site.CaptureRequirementIfAreNotEqual<int>(
                changedPidTagMessageSize,
                currentPidTagMessageSize,
                3055009,
                @"[In Read-Only Properties] The PidTagMessageSize property is read-only.");

            #endregion
            #endregion

            #region Step11: Test PidTagMessageSizeExtended

            #region Variables
            PropertyTag propertyTagMessageSizeExtended;
            uint resultMessageSizeExtended;
            TaggedPropertyValue taggedMessageSizeExtended = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagMessageSizeExtended's id is 0x0E08
            const ushort PidTagMessageSizeExtended = 0x0E08;
            #endregion

            #region RopGetPropertiesSpesific to read a readable property PidTagMessageSizeExtended
            propertyTagMessageSizeExtended.PropertyId = PidTagMessageSizeExtended;

            // According to [MS-OXPROPS], PidTagMessageSizeExtended's data type is 0x0014
            propertyTagMessageSizeExtended.PropertyType = 0x0014;
            uint responseFlagPidTagMessageSizeExtended;
            resultMessageSizeExtended = this.GetPropertyValue(propertyTagMessageSizeExtended, out responseFlagPidTagMessageSizeExtended, out propertyValue);
            long originalPidTagMessageSizeExtended = BitConverter.ToInt64(propertyValue, 0);
            #endregion

            #region Capture R635

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3065");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3065
            // ResponseFlagPidTagMessageSizeExtended is 0 indicates that the RopGetPropertiesSpecific ROP is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagPidTagMessageSizeExtended,
                3065,
                @"[In PidTagMessageSizeExtended Property] Type: PtypInteger64 ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region RopSetProperties ROP to write a non-writable property PidTagMessageSizeExtended
            taggedMessageSizeExtended.PropertyTag = propertyTagMessageSizeExtended;
            long messageSizeExtendedValue = originalPidTagMessageSizeExtended + 1;
            taggedMessageSizeExtended.Value = BitConverter.GetBytes(messageSizeExtendedValue);
            resultMessageSizeExtended = this.TrySetLogonProperty(taggedMessageSizeExtended);

            resultMessageSizeExtended = this.GetPropertyValue(propertyTagMessageSizeExtended, out responseFlagPidTagMessageSizeExtended, out propertyValue);
            long currentPidTagMessageSizeExtended = BitConverter.ToInt64(propertyValue, 0);
            #endregion

            #region Capture R3055010
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055010");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055010
            Site.CaptureRequirementIfAreNotEqual<long>(
                messageSizeExtendedValue,
                currentPidTagMessageSizeExtended,
                3055010,
                @"[In Read-Only Properties] The PidTagMessageSizeExtended property is read-only.");
            #endregion
            #endregion

            #region Step12: Test PidTagUserEntryId

            #region Variables
            PropertyTag propertyTagUserEntryId;
            uint resultUserEntryId;
            TaggedPropertyValue taggedUserEntryId = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagUserEntryId's id is 0x6619
            const ushort PidTagUserEntryId = 0x6619;
            #endregion

            #region Read a readable property PidTagUserEntryId
            propertyTagUserEntryId.PropertyId = PidTagUserEntryId;

            // According to [MS-OXPROPS], PidTagUserEntryId's data type is 0x0102
            propertyTagUserEntryId.PropertyType = 0x0102;
            uint responseFlagTagUserEntryId;

            resultUserEntryId = this.GetPropertyValue(propertyTagUserEntryId, out responseFlagTagUserEntryId, out propertyValue);
            byte[] originalPidTagUserEntryId = new byte[propertyValue.Length];
            Array.Copy(propertyValue, originalPidTagUserEntryId, propertyValue.Length);
            #endregion

            #region Capture R3066, R59737

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3066");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3066
            // The value responseFlagTagUserEntryId is 0 indicates that the RopGetPropertiesSpecific is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagTagUserEntryId,
                3066,
                @"[In PidTagUserEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1).");
            #endregion

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59737");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59737
            // The value responseFlagTagUserEntryId is 0 indicates that the RopGetPropertiesSpecific is implemented successfully.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagTagUserEntryId,
                59737,
                @"[in PidTagUserEntryId Property] This property is also available on private mailbox logon (section 2.2.2.1.1.11).");

            #region Write a non-writable property PidTagUserEntryId
            taggedUserEntryId.PropertyTag = propertyTagUserEntryId;
            byte[] userEntryIdValue = new byte[2] { 0x00, 0x00 };
            taggedUserEntryId.Value = userEntryIdValue;
            resultUserEntryId = this.TrySetLogonProperty(taggedUserEntryId);

            resultUserEntryId = this.GetPropertyValue(propertyTagUserEntryId, out responseFlagTagUserEntryId, out propertyValue);
            byte[] currentPidTagUserEntryId = propertyValue;
            #endregion

            #region Capture R3055011

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3055011");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3055011
            Site.CaptureRequirementIfIsFalse(
                Common.CompareByteArray(taggedUserEntryId.Value, currentPidTagUserEntryId),
                3055011,
                @"[In Read-Only Properties] The PidTagUserEntryId property is read-only.");

            #endregion
            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements of read/write properties on the private mailbox logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC10_TestOperationsOnReadWriteProperties()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step2: RopLogon ROP log on to private mailbox.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            #region Step3: Test PidTagComment
            #region Variables
            PropertyTag propertyTagComment = new PropertyTag();

            // According to [MS-OXPROPS], PidTagComment's id is 0x3004
            const ushort PidTagComment = 0x3004;
            #endregion

            #region RopGetProperties ROP to read property PidTagComment
            propertyTagComment.PropertyId = PidTagComment;

            // According to [MS-OXPROPS], PidTagComment's data type is 0x001F
            propertyTagComment.PropertyType = 0x001F;

            TaggedPropertyValue taggedTagComment = new TaggedPropertyValue();
            taggedTagComment.PropertyTag = propertyTagComment;
            string tagCommentValue = "PidTagCommentSample";
            taggedTagComment.Value = Encoding.Unicode.GetBytes(tagCommentValue + "\0");
            this.TrySetLogonProperty(taggedTagComment);

            byte[] valueTagComment;
            uint responseFlagForGetVauleTagComment;
            string tagCurrentCommentValue;
            this.GetPropertyValue(propertyTagComment, out responseFlagForGetVauleTagComment, out valueTagComment);

            // The value responseFlagForGetVauleTagComment is 0 indicates the RopGetPropertiesSpecific ROP is implemented successfully.
            if (0 == responseFlagForGetVauleTagComment)
            {
                tagCurrentCommentValue = System.Text.Encoding.Unicode.GetString(valueTagComment);
                int lastIndex = tagCurrentCommentValue.Length - 1;
                if (tagCurrentCommentValue[lastIndex] == '\0')
                {
                    tagCurrentCommentValue = tagCurrentCommentValue.Remove(lastIndex);
                }
            }
            else
            {
                tagCurrentCommentValue = string.Empty;
            }

            #endregion

            #region Capture code

            if (Common.IsRequirementEnabled(6280001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R6280001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R6280001
                // The value responseFlagForGetVauleTagComment is 0 indicates the RopGetPropertiesSpecific ROP is implemented successfully.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseFlagForGetVauleTagComment,
                    6280001,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagComment property as read-only. <13> Section 2.2.2.1.2.1:  The PidTagComment property is read-only in Microsoft Exchange Server 2013 Service Pack 1 (SP1).");
            }

            if (Common.IsRequirementEnabled(6280002, this.Site))
            {
                // The value responseFlagForGetVauleTagComment is 0 indicates the RopGetPropertiesSpecific ROP is implemented successfully.
                Site.Assert.AreEqual<uint>(0, responseFlagForGetVauleTagComment, "Read property PidTagComment with RopGetPropertiesSpecific should be successful.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R6280002");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R6280002
                Site.CaptureRequirementIfAreEqual<string>(
                    tagCommentValue,
                    tagCurrentCommentValue,
                    6280002,
                    @"[In Appendix A: Product Behavior] Implementation does support the PidTagComment property as read-write. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R650");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R650. Verify the PidTagComment property values set and got are same.
                Site.CaptureRequirementIfAreEqual<string>(
                    tagCommentValue,
                    tagCurrentCommentValue,
                    650,
                    @"[In PidTagComment Property] The PidTagComment property ([MS-OXPROPS] section 2.628) contains a mailbox comment.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3069");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3069
                // The value responseFlagForGetVauleTagComment is 0 indicates the RopGetPropertiesSpecific ROP is implemented successfully.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseFlagForGetVauleTagComment,
                    3069,
                    @"[In PidTagComment Property] Type: PtypString ([MS-OXCDATA] section 2.11.1).");
            }
            #endregion
            #endregion

            #region Step4: Test PidTagDeleteAfterSubmit

            #region Variables
            PropertyTag propertyTagDeleteAfterSubmit;
            uint resultSetDeleteAfterSubmit;
            uint resultDelDeleteAfterSubmit;
            TaggedPropertyValue taggedDeleteAfterSubmit = new TaggedPropertyValue();

            // According to the Open Specification MS-OXPROPS, PidTagDeleteAfterSubmit's id is 0x0E01
            const ushort PidTagDeleteAfterSubmit = 0x0E01;
            #endregion

            #region Write a writable property PidTagDeleteAfterSubmit
            // According to [MS-OXPROPS], PidTagDeleteAfterSubmit's data type is 0x000B
            propertyTagDeleteAfterSubmit.PropertyType = 0x000B;
            propertyTagDeleteAfterSubmit.PropertyId = PidTagDeleteAfterSubmit;
            taggedDeleteAfterSubmit.PropertyTag = propertyTagDeleteAfterSubmit;
            taggedDeleteAfterSubmit.Value = new byte[1];
            taggedDeleteAfterSubmit.Value[0] = 0x00;

            resultSetDeleteAfterSubmit = this.TrySetLogonProperty(taggedDeleteAfterSubmit);
            bool valueofInitial = BitConverter.ToBoolean(this.TryGetLogonPropertyValue(propertyTagDeleteAfterSubmit), 0);

            #endregion

            #region Capture R3068008, R3128
            if (Common.IsRequirementEnabled(3068008, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3068008");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3068008
                bool isVerifyR3068008 = (0 == resultSetDeleteAfterSubmit) && (valueofInitial == false);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR3068008,
                    3068008,
                    @"[In Appendix A: Product Behavior] Implementation support set the PidTagDeleteAfterSubmit property by using the RopSetProperties ROP. (Exchange 2007 and Exchange 2010 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(3128, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3128");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3128
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070005,
                    resultSetDeleteAfterSubmit,
                    3128,
                    @"[In Appendix A: Product Behavior] The implementation returns 0x80070005 (ecAccessDenied) when the client attempts to set PidTagDeleteAfterSubmit property by using the RopSetProperties ROP. (<14> Section 2.2.2.1.2.2: In Exchange 2013 and Exchange 2016 the server returns 0x80070005 (ecAccessDenied) when the client attempts to set this property [PidTagDeleteAfterSubmit] by using the RopSetProperties ROP ([MS-OXCROPS] section 2.2.8.6).)");
            }
            #endregion

            #region Read a readable property PidTagDeleteAfterSubmit
            uint responseFlagDelAfterSubmit;
            this.TryGetFlagLogonProperty(propertyTagDeleteAfterSubmit, out responseFlagDelAfterSubmit);
            #endregion

            if (Common.IsRequirementEnabled(3068008, this.Site))
            {
                #region Capture R3070
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3070");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3070
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseFlagDelAfterSubmit,
                    3070,
                    @"[In PidTagDeleteAfterSubmit Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1)");
                #endregion
            }

            #region Delete a deletable property PidTagDeleteAfterSubmit
            resultDelDeleteAfterSubmit = this.TryDeleteLogonProperty(propertyTagDeleteAfterSubmit);
            #endregion

            #region Capture R654001

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R654001");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R654001
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                resultDelDeleteAfterSubmit,
                654001,
                @"[In PidTagDeleteAfterSubmit Property] The client can also delete this property [PidTagDeleteAfterSubmit].");

            #endregion

            #endregion

            #region Step5: Test PidTagDisplayName

            #region Variables
            PropertyTag propertTagDisplayName;
            TaggedPropertyValue taggedDisplayName = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagDisplayName's id is 0x3001
            const ushort PidTagDisplayName = 0x3001;
            #endregion

            #region ROPGetPropertiesSpecific ROP to read a readable property PidTagDisplayName
            propertTagDisplayName.PropertyId = PidTagDisplayName;

            // According to [MS-OXPROPS], PidTagDisplayName's data type is 0x001F
            propertTagDisplayName.PropertyType = 0x001F;
            uint responseDisplayName;
            this.TryGetFlagLogonProperty(propertTagDisplayName, out responseDisplayName);

            if (Common.IsRequirementEnabled(306800301, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R306800301");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R306800301
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseDisplayName,
                    306800301,
                    @"[In Appendix A: Product Behavior] Implementation does support PidTagDisplayName  property as read-only. <15> Section 2.2.2.1.2.3:  In Exchange 2013 SP1  this property is read-only.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3071");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3071
                // The value responseDisplayName indicates the ROP RopGetPropertiesSpecific is implemented successfully.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseDisplayName,
                    3071,
                    @"[In PidTagDisplayName Property] Type: PtypString ([MS-OXCDATA] section 2.11.1)");
            }
            #endregion

            #region RopSetProperties ROP to write a writable property PidTagDisplayName
            taggedDisplayName.PropertyTag = propertTagDisplayName;
            string strSetDisplayName = "MailboxDisplayName\0";
            taggedDisplayName.Value = Encoding.Unicode.GetBytes(strSetDisplayName);
            this.TrySetLogonProperty(taggedDisplayName);
            string newDisplayName = Encoding.Unicode.GetString(this.TryGetLogonPropertyValue(propertTagDisplayName));

            if (Common.IsRequirementEnabled(306800302, this.Site))
            {
                // The value responseFlagForGetVauleTagComment is 0 indicates the RopGetPropertiesSpecific ROP is implemented successfully.
                Site.Assert.AreEqual<uint>(0, responseDisplayName, "Read property PidTagComment with RopGetPropertiesSpecific should be successful.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R306800302");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R306800302
                // The value of responseDisplayName is 0 indicates the ROPGetPropertiesSpecific ROP is implemented successfully.
                // When the value of strSetDisplayName is the same as newDisplayName, it indicates that the ROPSetProperties ROP is implemented successfully.
                // So the MS-OXCSTOR_R3068003 is verified.
                Site.CaptureRequirementIfIsTrue(
                    newDisplayName.Equals(strSetDisplayName),
                    306800302,
                    @"[In Appendix A: Product Behavior] Implementation does support PidTagDisplayName  property as read-write. (Exchange 2007, Exchange 2010 and Exchange 2016 follow this behavior).");
            }
            #endregion

            #endregion 

            #region Step6: Test PidTagOutOfOfficeState

            #region Variables
            PropertyTag propertyTagOutOfOfficeState;
            TaggedPropertyValue taggedOutOfOfficeState = new TaggedPropertyValue();

            // According to [MS-OXPROPS], PidTagOutOfOfficeState's id is 0x661D
            const ushort PidTagOutOfOfficeState = 0x661D;
            #endregion

            #region RopSetProperties ROP to write a writable property PidTagOutOfOfficeState
            // According to [MS-OXPROPS], PidTagOutOfOfficeState's data type is 0x000B
            propertyTagOutOfOfficeState.PropertyType = 0x000B;
            propertyTagOutOfOfficeState.PropertyId = PidTagOutOfOfficeState;

            uint responseFlag;
            byte[] propertyValue;
            uint resultGetOutOfOfficeState = this.GetPropertyValue(propertyTagOutOfOfficeState, out responseFlag, out propertyValue);
            byte originalPidTagOutOfOfficeState = propertyValue[0];

            taggedOutOfOfficeState.PropertyTag = propertyTagOutOfOfficeState;
            taggedOutOfOfficeState.Value = new byte[1];
            taggedOutOfOfficeState.Value[0] = (byte)(originalPidTagOutOfOfficeState == 0 ? 1 : 0);

            uint resultSetOutOfOfficeState = this.TrySetLogonProperty(taggedOutOfOfficeState);

            // Wait 5 seconds to enable the PidTagOutOfOfficeState
            Thread.Sleep(5000);
            #endregion

            #region RopGetPropertiesSpecific ROP to read a readable property PidTagOutOfOfficeState
            uint responseFlagOutOfOfficeState;
            resultGetOutOfOfficeState = this.GetPropertyValue(propertyTagOutOfOfficeState, out responseFlagOutOfOfficeState, out propertyValue);
            byte currentPidTagOutOfOfficeState = propertyValue[0];
            #endregion

            #region Capture R3072
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3072");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3072
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                responseFlagOutOfOfficeState,
                3072,
                @"[In PidTagOutOfOfficeState Property] Type: PtypBoolean ([MS-OXCDATA] section 2.11.1)");
            #endregion

            #region Capture R3068004
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3068004");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3068004
            Site.Assert.AreEqual<uint>(0, resultSetOutOfOfficeState, "Check if the RopSetProperties ROP is executed successfully");
            Site.Assert.AreEqual<uint>(0, resultGetOutOfOfficeState, "Check if the RopGetProperties ROP is executed successfully");

            // When the operation failed it indicates that the PidTagOutOfOfficeState property cannot be deleted.
            Site.CaptureRequirementIfAreNotEqual<byte>(
                originalPidTagOutOfOfficeState,
                currentPidTagOutOfOfficeState,
                3068004,
                @"[In Read/Write Properties] The PidTagOutOfOfficeState property is read/write.");
            #endregion

            #endregion

            #region Step7: Test PidTagSentMailSvrEID

            #region Variables
            PropertyTag propertyTagSentMailSvrEID;
            uint resultSentMailSvrEID;

            // According to [MS-OXPROPS], PidTagSentMailSvrEID's id is 0x6740
            const ushort PidTagSentMailSvrEID = 0x6740;

            #endregion

            #region  RopGetPropertiesSpecific ROP to read a readable property PidTagSentMailSvrEID
            propertyTagSentMailSvrEID.PropertyId = PidTagSentMailSvrEID;

            // According to [MS-OXPROPS], PidTagSentMailSvrEID's data type is 0x00FB
            propertyTagSentMailSvrEID.PropertyType = 0x00FB;
            #endregion

            #region Delete a deletable property PidTagSentMailSvrEID
            resultSentMailSvrEID = this.TryDeleteLogonProperty(propertyTagSentMailSvrEID);
            #endregion

            #region Capture R59734

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59734");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59734
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                resultSentMailSvrEID,
                59734,
                @"[in PidTagSentMailSvrEID Property] The client can also delete this property.");

            #endregion

            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopLogon ROP.
        /// This test case depends on second SUT. If second SUT is not present, some steps of this test case cannot be executed.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC11_TestRopLogonErrorCodes()
        {
            this.CheckTransportIsSupported();

            Site.Assume.IsFalse(this.transportByMAPIOrNot, "This case doesn't support the mapi_http transport.");

            #region Step1: Connect to the server
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");

            this.oxcstorAdapter.SetAutoRedirect(false);
            #endregion Step1

            #region Step2: Call RopLogon ROP with Essdn that is unknown to the system, expect ecUnknownUser
            this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
            this.logonRequestForPrivateMailBox.OpenFlags = 0x01000000; // Set USE_ADMIN_PRIVILEGE and USE_PER_MDB_REPLID_MAPPING flags.
            string invalidEssdn = "/o=contoso/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=nonexistedUser\0";
            this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(invalidEssdn);
            this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R173");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R173
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000003EB,
                this.logonResponse.ReturnValue,
                173,
                @"[In RopLogon ROP Common Return Codes] The value of return code ecUnknownUser: 0x000003EB.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59746");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59746
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x000003EB,
                this.logonResponse.ReturnValue,
                59746,
                @"[in Private Mailbox Logon] If the user has no configured mailbox database, the ROP MUST fail with a ReturnValue of ecUnknownUser, as specified in section 3.2.5.1.3.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R788");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R788
            Site.CaptureRequirementIfAreEqual<uint>(
                0x000003EB,
                this.logonResponse.ReturnValue,
                788,
                @"[In Private Mailbox Logon] If lookup fails specifically because the ESSDN could not be found, the server MUST fail the operation with a ReturnValue of ecUnknownUser (section 3.2.5.1.3).");

            if (Common.IsRequirementEnabled(1273, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1273");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1273
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0x000003EB,
                    this.logonResponse.ReturnValue,
                    1273,
                    @"[In Appendix A: Product Behavior] Implementation does not return ecLoginFailure. <36> Section 3.2.5.1.3: If the user doesn't exist in the Active Directory forest, Exchange 2007 and Exchange 2010, Exchange 2013 and Exchange 2016 return ecUnknownUser.");
            }
            #endregion Capture
            #endregion Step2

            if (Common.IsRequirementEnabled(3135, this.Site) || Common.IsRequirementEnabled(1346, this.Site) || Common.IsRequirementEnabled(3905001, this.Site))
            {
                #region Step3: Disconnect to server
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
                #endregion Step3

                #region Step4: Connect to server
                this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PrivateMailboxServer, this.essdnOfUser1, this.domain, this.userNameOfMailboxOnServer1, this.passwordOfUserOfMailboxOnServer1);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step4

                #region Step5: Call RopLogon ROP to log on to the non-owner mailbox
                this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping | (uint)OpenFlags.UseAdminPrivilege;
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.userDN + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

                #region Capture
                if (Common.IsRequirementEnabled(3135, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3135");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3135
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80070005,
                        this.logonResponse.ReturnValue,
                        3135,
                        @"[In Appendix A: Product Behavior] Implementation does fail the operation [RopLogon] with ecAccessDenied [if the user does not match the owner of the mailbox]. (<31> Section 3.2.5.1.1: Exchange 2003 and Exchange 2007 return ecAccessDenied.)");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R187");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R187
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80070005,
                        this.logonResponse.ReturnValue,
                        187,
                        @"[In RopLogon ROP Common Return Codes] The value of return code ecAccessDenied: 0x80070005.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R188");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R188
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80070005,
                        this.logonResponse.ReturnValue,
                        188,
                        @"[In RopLogon ROP Common Return Codes] The meaning of return code ecAccessDenied: The user does not have sufficient permissions to the mailbox.");
                }

                if (Common.IsRequirementEnabled(1346, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1346");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1346
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x000003F2,
                        this.logonResponse.ReturnValue,
                        1346,
                        @"[In Appendix A: Product Behavior]  Implementation does not fail the operation [RopLogon] with ecAccessDenied [if the user does not match the owner of the mailbox]. (<31> Section 3.2.5.1.1: Exchange 2010 return ecLoginPerm.)");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R189");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R189
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x000003F2,
                        this.logonResponse.ReturnValue,
                        189,
                        @"[In RopLogon ROP Common Return Codes] The value of return code ecLoginPerm: 0x000003F2.");
                }

                if (Common.IsRequirementEnabled(3905001, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3905001");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3905001
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000000,
                        this.logonResponse.ReturnValue,
                        3905001,
                        @"[In Appendix A: Product Behavior]  Implementation does not fail the operation [RopLogon] with ecAccessDenied [if the user does not match the owner of the mailbox]. (<31> Section 3.2.5.1.1: Exchange 2016 and Exchange 2013 return ecNone.)");
                }
                #endregion Capture
                #endregion Step5
            }

            #region Step6: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
            #endregion Step6

            if (this.secondSUTExistOrNot)
            {
                #region Step7: Connect to server
                this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step7

                #region Step8: RopLogon ROP to log on to the mailbox hosted in another server
                this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping;
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.essdnOfUser2 + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.Others, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

                #region Capture
                if (Common.IsRequirementEnabled(790, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R182");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R182
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000478,
                        this.logonResponse.ReturnValue,
                        182,
                        @"[In RopLogon ROP Common Return Codes] The value of return code ecWrongServer: 0x00000478.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R183");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R183
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000478,
                        this.logonResponse.ReturnValue,
                        183,
                        @"[In RopLogon ROP Common Return Codes] The meaning of return code ecWrongServer: The requested message store for logon is not the user's home message store.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R790");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R790
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000478,
                        this.logonResponse.ReturnValue,
                        790,
                        @"[In Private Mailbox Logon] If the user's configured mailbox is not hosted by this server, the server determines the name of the correct server hosting the user's mailbox and fail the ROP with a ReturnValue of ecWrongServer, as specified in section 3.2.5.1.3");
                }
                #endregion Capture
                #endregion Step8

                #region Step9: Disconnect
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");
                #endregion Step9

                #region Step10: Connect to the server
                this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step10

                #region Step11: Call RopLogon ROP 6 times in 10-second to log on to a mailbox that is not hosted on the server, expect ecServerPaused
                this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = 0x01000000; // Set USE_PER_MDB_REPLID_MAPPING flag.
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(this.essdnOfUser2 + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;

                DateTime begin6TimesLogon = DateTime.Now;
                for (int i = 0; i < 6; i++)
                {
                    this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.Others, out this.outputBuffer);
                    this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                }

                DateTime end6TimesLogon = DateTime.Now;
                TimeSpan elapsedTime = new TimeSpan(end6TimesLogon.Ticks - begin6TimesLogon.Ticks);
                double elapsedSecond = elapsedTime.TotalSeconds;

                Site.Log.Add(LogEntryKind.Comment, "The client has made 6 attempts in a {0}-second period to log on to a mailbox that is not hosted on the server. return value: 0x{1:X}", elapsedSecond, this.logonResponse.ReturnValue);

                // Call RopLogon ROP 6 times in 10 seconds
                if (elapsedSecond < 10)
                {
                    if (Common.IsRequirementEnabled(3103001, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3103001");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3103001
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0,
                            this.logonResponse.ReturnValue,
                            3103001,
                            @"[In Appendix A: Product Behavior] If the client has made more than five attempts within a 10-second period to log on to a mailbox that is not hosted on the server, the implementation does Success the operation with ecServerPaused in the ReturnValue field. <22> Section 3.2.5.1.1:  Exchange 2010 and Exchange 2013 return Success with ecServerPaused.");
                    }

                    if (Common.IsRequirementEnabled(3103002, this.Site))
                    {
                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R191");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R191
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000047F,
                            this.logonResponse.ReturnValue,
                            191,
                            @"[In RopLogon ROP Common Return Codes] The value of return code ecServerPaused: 0x0000047F.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R192");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R192
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000047F,
                            this.logonResponse.ReturnValue,
                            192,
                            @"[In RopLogon ROP Common Return Codes] The meaning of return code ecServerPaused: The client has made more than five attempts within a 10-second period to log on to a mailbox that is not hosted on the server.");

                        // Add the debug information
                        this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3103002");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3103002
                        this.Site.CaptureRequirementIfAreEqual<uint>(
                            0x0000047F,
                            this.logonResponse.ReturnValue,
                            3103002,
                            @"[In Appendix A: Product Behavior] If the client has made more than five attempts within a 10-second period to log on to a mailbox that is not hosted on the server, the implementation does fail the operation with ecServerPaused in the ReturnValue field. (Exchange 2007 and Exchange 2016 follow this behavior).");
                    }
                }
                #endregion Step11

                #region Step12: Disconnect to the server
                this.returnStatus = this.oxcstorAdapter.DisconnectEx();
                Site.Assert.IsTrue(this.returnStatus, "Disconnection  is successful");

                // Sleep 10 seconds to avoid subsequent cases fail with ecServerPaused (0x0000047F)
                Thread.Sleep(10 * 1000);
                #endregion Step12
            }

            #region Capture
            if (Common.IsRequirementEnabled(193, this.Site) ||
                Common.IsRequirementEnabled(1268001, this.Site))
            {
                #region Step13: Connect to the server
                string userForDisableMailbox = Common.GetConfigurationPropertyValue(ConstValues.UserForDisableMailbox, this.Site);
                string passwordForDisableMailbox = Common.GetConfigurationPropertyValue(ConstValues.PasswordForDisableMailbox, this.Site);
                string essdnForDisableMailbox = this.sutControlAdapter.GetUserDN(this.server1Name, userForDisableMailbox);
                Site.Assert.IsFalse(string.IsNullOrEmpty(essdnForDisableMailbox), "The legacyExchangeDN of {0} must not be null or empty. The actual value: {1}", userForDisableMailbox, essdnForDisableMailbox);

                this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PrivateMailboxServer, essdnForDisableMailbox, this.domain, userForDisableMailbox, passwordForDisableMailbox);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion Step13

                #region Step14: RopLogon ROP to log on to the private mailbox that will be disabled in following steps
                this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = 0x01000000; // Set USE_PER_MDB_REPLID_MAPPING flag.
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(essdnForDisableMailbox + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
                #endregion Step14

                #region Step15: Disable the mailbox
                string status = this.sutControlAdapter.DisableMailbox(userForDisableMailbox);
                Site.Assert.AreEqual<string>("success", status.ToLower(), "\"success\" indicates mailbox is disabled, otherwise the mailbox is not disabled and error information is returned.");
                #endregion

                #region Step16: RopLogon to log on to the disabled mailbox
                this.returnStatus = this.oxcstorAdapter.ConnectEx(this.server1Name, ConnectionType.PrivateMailboxServer, essdnForDisableMailbox, this.domain, userForDisableMailbox, passwordForDisableMailbox);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");

                this.logonRequestForPrivateMailBox.LogonFlags = 0x01; // The Private flag is set
                this.logonRequestForPrivateMailBox.OpenFlags = 0x01000000; // Set USE_ADMIN_PRIVILEGE and USE_PER_MDB_REPLID_MAPPING flags.
                this.logonRequestForPrivateMailBox.Essdn = Encoding.ASCII.GetBytes(essdnForDisableMailbox + "\0");
                this.logonRequestForPrivateMailBox.EssdnSize = (ushort)this.logonRequestForPrivateMailBox.Essdn.Length;
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

                if (Common.IsRequirementEnabled(193, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R193");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R193
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x000003EB,
                        this.logonResponse.ReturnValue,
                        193,
                        @"[In Appendix A: Product Behavior] The implementation returns ecUnknownUser [if the client attempts to log on to a mailbox that is disabled]. (<20> Section 3.2.5.1.1: Exchange 2010, Exchange 2013 and Exchange 2016 return ecUnknownUser.)");
                }

                if (Common.IsRequirementEnabled(1268001, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1268001");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1268001
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x80040111,
                        this.logonResponse.ReturnValue,
                        1268001,
                        @"[In Appendix A: Product Behavior] The implementation returns ecLoginFailure [if the client attempts to log on to a mailbox that is disabled]. (<20> Section 3.2.5.1.1: Exchange 2007 returns ecLoginFailure, then, after 5 minutes, the server returns ecUnknownUser.)");

                    // Sleep 5 minutes
                    System.Threading.Thread.Sleep(300 * 1000);

                    this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
                    this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1268001");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1268001
                    this.Site.CaptureRequirementIfAreEqual<uint>(
                        0x000003EB,
                        this.logonResponse.ReturnValue,
                        1268001,
                        @"[In Appendix A: Product Behavior] The implementation returns ecLoginFailure [if the client attempts to log on to a mailbox that is disabled]. (<20> Section 3.2.5.1.1: Exchange 2007 returns ecLoginFailure, then, after 5 minutes, the server returns ecUnknownUser.)");
                }
            }
                #endregion Capture
            #endregion

            #region Step17: Enable the auto redirect
            this.oxcstorAdapter.SetAutoRedirect(true);
            #endregion
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopGetReceiveFolder ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC12_TestRopGetReceiveFolderErrorCodes()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to public folder
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step2: RopLogon to log on to the public folder
            // This operation is only valid when the Logon object refers to a private mailbox logon
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step2

            #region Step3: RopGetReceiveFolder ROP against the public folder logon
            string ipcUpcase = "IPM\0";
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(ipcUpcase);
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            RopGetReceiveFolderResponse getReceiveFolderResponseForPublicFolder = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59783");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59783
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                59783,
                @"[In Receiving a RopGetReceiveFolder ROP Request] If the operation is performed against a public folders logon, the server MUST fail the operation with the ReturnValue field set to 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R238");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R238
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                238,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The value of error code ecNotSupported: Its value is 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R239, The RopGetReceiveFolder ROP was performed against a private mailbox logon.");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R239
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                239,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The meaning of error code ecNotSupported: The ROP was not performed against a private mailbox logon.");
            #endregion Capture
            #endregion Step3

            #region Step4: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            #endregion Step4

            #region Step5: Connect to the private mailbox
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connect successfully");
            #endregion Step5

            #region Step6: RopLogon ROP to log on to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step6

            #region Step7: RopGetReceiveFolder ROP against the private logon
            string ipm = "ipm\0";
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(ipm);
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCSTOR_R59782. Calling the RopGetReceiveFolder ROP against public folder, the return value: 0X{0:X}. Calling the RopGetReceiveFolder ROP against mailbox, the return value: 0X{1:X}.",
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                this.getReceiveFolderResponse.ReturnValue);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59782
            bool isVerifiedR59782 = this.getReceiveFolderResponse.ReturnValue == 0 && getReceiveFolderResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR59782,
                59782,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");

            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCSTOR_R198. Calling the RopGetReceiveFolder ROP against public folder, the return value: 0X{0:X}. Calling the RopGetReceiveFolder ROP against mailbox, the return value: 0X{1:X}.",
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                this.getReceiveFolderResponse.ReturnValue);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R198
            bool isVerifiedR198 = this.getReceiveFolderResponse.ReturnValue == 0 && getReceiveFolderResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR198,
                198,
                @"[In RopGetReceiveFolder ROP Request Buffer] This operation [RopGetReceiveFolder] is only valid when the Logon object refers to a private mailbox logon.");

            // Add the debug information
            this.Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCSTOR_R249. Calling the RopGetReceiveFolder ROP against public folder, the return value: 0X{0:X}. Calling the RopGetReceiveFolder ROP against mailbox, the return value: 0X{1:X}.",
                getReceiveFolderResponseForPublicFolder.ReturnValue,
                this.getReceiveFolderResponse.ReturnValue);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R249
            bool isVerifiedR249 = this.getReceiveFolderResponse.ReturnValue == 0 && getReceiveFolderResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR249,
                249,
                @"[In RopSetReceiveFolder ROP Request Buffer] This operation [RopSetReceiveFolder] MUST be issued against a private mailbox logon.");
            #endregion Capture
            #endregion Step7

            #region Step8: RopGetReceiveFolder ROP with the message class beginning with a period
            string wrongMessageClass = ".My.Class\0";
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(wrongMessageClass);
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59785, The MessageClass: {0}", wrongMessageClass.Replace("\0", string.Empty));

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59785
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.getReceiveFolderResponse.ReturnValue,
                59785,
                @"[In Receiving a RopGetReceiveFolder ROP Request] If the value does not conform to the requirements [MessageClass field in section 2.2.1.2.1], then the server MUST fail the operation with the ReturnValue field set to 0x80070057.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R234");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R234
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.getReceiveFolderResponse.ReturnValue,
                234,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The value of error code ecInvalidParam: 0x80070057.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R235");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R235
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.getReceiveFolderResponse.ReturnValue,
                235,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The meaning of error code ecInvalidParam: The MessageClass value does not conform to the format requirements specified in section 2.2.1.2.1.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59784");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59784
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.getReceiveFolderResponse.ReturnValue,
                59784,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server validates the value of the MessageClass field, as specified in section 2.2.1.2.1.");

            Site.Log.Add(LogEntryKind.Comment, "The RopGetReceiveFolder ROP with MessageClass: {0}", Encoding.ASCII.GetString(this.getReceiveFolderRequest.MessageClass).Replace("\0", string.Empty));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R205");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R205
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.getReceiveFolderResponse.ReturnValue,
                205,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: The string does not begin with a period (""."").");
            #endregion Capture
            #endregion Step8

            #region Step9: RopGetReceiveFolder ROP with message class end with a period
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes("My.Class.\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "The MessageClass in RopGetReceiveFolder ROP: {0}", Encoding.ASCII.GetString(this.getReceiveFolderRequest.MessageClass).Replace("\0", string.Empty));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R206");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R206
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.getReceiveFolderResponse.ReturnValue,
                206,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: The string does not end with a period.");
            #endregion Capture
            #endregion Step9

            #region Step10: RopGetReceiveFolder ROP with the message class that contains adjacent periods
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes("My..Class." + "\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            this.getReceiveFolderResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "The MessageClass in RopGetReceiveFolder ROP: {0}", Encoding.ASCII.GetString(this.getReceiveFolderRequest.MessageClass).Replace("\0", string.Empty));

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R207");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R207
            this.Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                this.getReceiveFolderResponse.ReturnValue,
                207,
                @"[In RopGetReceiveFolder ROP Request Buffer] MessageClass: The string [Contained by MessageClass] MUST meet the following requirement: The string does not contain adjacent periods.");
            #endregion Capture
            #endregion Step10
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopSetReceiveFolder ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC13_TestRopSetReceiveFolderErrorCodes()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to public folder
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon ROP to log on to the public folder
            // This operation is only valid when the Logon object refers to a private mailbox logon
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion Step2

            #region Step3: RopSetReceiveFolder ROP against the public folder logon
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[2];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            RopSetReceiveFolderResponse setReceiveFolderResponseForPublicFolder;
            setReceiveFolderResponseForPublicFolder = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "RopSetReceiveFolder ROP against the public folder.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R279");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R279
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                setReceiveFolderResponseForPublicFolder.ReturnValue,
                279,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The value of error code ecNotSupported: 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R280");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R280
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                setReceiveFolderResponseForPublicFolder.ReturnValue,
                280,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecNotSupported: The ROP was not performed against a private mailbox logon.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R856");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R856
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                setReceiveFolderResponseForPublicFolder.ReturnValue,
                856,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If the operation is performed against a public folders logon, then the server MUST fail the operation with the ReturnValue field set to 0x80040102.");
            #endregion Capture
            #endregion Step3

            #region Step4: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            #endregion Step4

            #region Step5: Connect to the private mailbox
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connect successfully");
            #endregion Step5

            #region Step6: RopLogon ROP to log on to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step6

            #region Step7: RopSetReceiveFolder ROP to set the receive folder
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[3];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(this.messageMyClass + "\0");
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            RopSetReceiveFolderResponse setReceiveFolderResponseForPrivateMailbox;
            setReceiveFolderResponseForPrivateMailbox = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "RopSetReceiveFolder ROP against the private mailbox.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R855. The return value of RopSetReceiveFolder ROP against public folder: 0x{0:x}", setReceiveFolderResponseForPublicFolder.ReturnValue);

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R855
            bool isVerifiedR855 = setReceiveFolderResponseForPrivateMailbox.ReturnValue == 0 && setReceiveFolderResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR855,
                855,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
            #endregion Capture
            #endregion Step7

            #region Step8: RopSetReceiveFolder ROP to set invalid message class
            string wrongMessageClass = ".My.Class\0";
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[2];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(wrongMessageClass);
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R273");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R273
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.setReceiveFolderResponse.ReturnValue,
                273,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The value of error code ecInvalidParam: 0x80070057.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R274");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R274
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.setReceiveFolderResponse.ReturnValue,
                274,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecInvalidParam: The message class string does not conform to the requirements specified in section 2.2.1.3.1.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R858");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R858
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.setReceiveFolderResponse.ReturnValue,
                858,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If the value does not conform to the requirements [MessageClass field in section 2.2.1.2.1], then the server MUST fail the operation with the ReturnValue field set to 0x80070057.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R845");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R845
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070057,
                this.setReceiveFolderResponse.ReturnValue,
                845,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The server validates the value of the MessageClass field, as specified in section 2.2.1.3.1.");
            #endregion Capture
            #endregion Step8

            #region Step9:  RopSetReceiveFolder ROP to set MessageClass is "Report.IPM\0"
            string messageClassSet = "Report.IPM\0";
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[3];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(messageClassSet);
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R271");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R271
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                this.setReceiveFolderResponse.ReturnValue,
                271,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The value of error code ecAccessDenied: 0x80070005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R272");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R272
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                this.setReceiveFolderResponse.ReturnValue,
                272,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecAccessDenied: The client has attempted to change the Receive folder for the ""IPM"" or ""Report.IPM"" classes.");
            #endregion Capture
            #endregion Step9

            #region Step10:  RopSetReceiveFolder ROP to set MessageClass is "IPM\0"
            messageClassSet = "IPM\0";
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[3];
            this.setReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes(messageClassSet);
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R271");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R271
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                this.setReceiveFolderResponse.ReturnValue,
                271,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The value of error code ecAccessDenied: 0x80070005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R272");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R272
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                this.setReceiveFolderResponse.ReturnValue,
                272,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecAccessDenied: The client has attempted to change the Receive folder for the ""IPM"" or ""Report.IPM"" classes.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R859");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R859
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80070005,
                this.setReceiveFolderResponse.ReturnValue,
                859,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If the value of the MessageClass field is a case-insensitive match to either ""IPM"", then the server MUST fail the operation with the ReturnValue field set to 0x80070005.");
            #endregion Capture
            #endregion Step10

            #region Step11:  RopSetReceiveFolder ROP to set MessageClass is empty and folder id is 0
            // Test ecAccessDenied while calling RopGetReceiveFolder ROP
            this.setReceiveFolderRequest.FolderId = 0;
            this.setReceiveFolderRequest.MessageClass = new byte[1] { 0 };
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            this.setReceiveFolderResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R860");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R860
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.setReceiveFolderResponse.ReturnValue,
                860,
                @"[In Receiving a RopSetReceiveFolder ROP Request] If the MessageClass field is set to a zero-length string and the FolderId field is set to zero, then the server MUST fail the operation with the ReturnValue field set to 0x80004005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R275");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R275
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.setReceiveFolderResponse.ReturnValue,
                275,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The value of error code ecError: 0x80004005.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R277");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R277
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80004005,
                this.setReceiveFolderResponse.ReturnValue,
                277,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The meaning of error code ecError: The FID (specified in the FolderId field) is all zeros AND the message class string (specified in the MessageClass field) has a length of zero.");
            #endregion Capture
            #endregion Step11
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopSetReceiveFolderTable ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC14_TestRopGetReceiveFolderTableErrorCodes()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to public folder
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon ROP to log on to public folder
            // This operation is only valid when the Logon object refers to a private mailbox logon
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step2

            #region Step3: RopGetReceiveFolderTable ROP against public folder logon
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderTableRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolderTable, out this.outputBuffer);
            RopGetReceiveFolderTableResponse getReceiveFolderTableResponseForPublicFolder = (RopGetReceiveFolderTableResponse)outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "RopGetReceiveFolderTable ROP against the public folders.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R324");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R324
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderTableResponseForPublicFolder.ReturnValue,
                324,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] the error code ecNotSupported: Its value is 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R325");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R325
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderTableResponseForPublicFolder.ReturnValue,
                325,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] the error code ecNotSupported: The ROP was not performed against a private mailbox logon.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R875");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R875
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderTableResponseForPublicFolder.ReturnValue,
                875,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] If the operation is performed against a public folders logon, the server MUST fail the operation with a ReturnValue of 0x80040102.");
            #endregion Capture
            #endregion Step3

            #region Step4: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            #endregion Step4

            #region Step5: Connect to the private mailbox
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connect successfully");
            #endregion Step5

            #region Step6: RopLogon ROP to log on to private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step6

            #region Step7: RopGetReceiveFolderTable ROP
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderTableRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolderTable, out this.outputBuffer);
            this.getReceiveFolderTableResponse = (RopGetReceiveFolderTableResponse)outputBuffer.RopsList[0];

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R874");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R874
            bool isVerifiedR874 = this.getReceiveFolderTableResponse.ReturnValue == 0 && getReceiveFolderTableResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR874,
                874,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
            #endregion Step6
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopGetPerUserLongTermIds ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC15_TestRopGetPerUserLongTermIdsErrorCodes()
        {
            this.CheckTransportIsSupported();

            #region Step1: Connect to public folder
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion Step1

            #region Step2: RopLogon ROP to log on to the public folder
            // This operation is only valid when the Logon object refers to a private mailbox logon
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step2

            #region Step3: RopGetPerUserLongTermIds ROP against the public folder logon
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsResponseForPublicFolder = (RopGetPerUserLongTermIdsResponse)outputBuffer.RopsList[0];

            #region Capture
            Site.Log.Add(LogEntryKind.Comment, "RopGetPerUserLongTermIds ROP against the public folders.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R482");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R482
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserLongTermIdsResponseForPublicFolder.ReturnValue,
                482,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The error code ecNotSupported: Its value is 0x80040102.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R483");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R483
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserLongTermIdsResponseForPublicFolder.ReturnValue,
                483,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The error code ecNotSupported: The ROP was attempted against a public folders logon.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59771");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59771
            this.Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getPerUserLongTermIdsResponseForPublicFolder.ReturnValue,
                59771,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] If the operation is performed against a public folders logon, the server MUST fail the operation with a ReturnValue of 0x80040102.");
            #endregion Capture
            #endregion Step3

            #region Step4: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnect successfully");
            #endregion Step4

            #region Step5: Connect to the private mailbox
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connect successfully");
            #endregion Step5

            #region Step6: RopLogon ROP to logon to the private mailbox
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion Step6

            #region Step7: RopGetPerUserLongTermIds ROP
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            this.getPerUserLongTermIdsResponse = (RopGetPerUserLongTermIdsResponse)outputBuffer.RopsList[0];

            Site.Log.Add(LogEntryKind.Comment, "RopGetPerUserLongTermIds ROP against to the private mailbox, the return value: {0}", this.getPerUserLongTermIdsResponse.ReturnValue);

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59770");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59770
            bool isVerifiedR59770 = this.getPerUserLongTermIdsResponse.ReturnValue == 0 && getPerUserLongTermIdsResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR59770,
                59770,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The server verifies that the operation is being performed against a private mailbox logon, not against a public folders logon.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R469");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R469
            bool isVerifiedR469 = this.getPerUserLongTermIdsResponse.ReturnValue == 0 && getPerUserLongTermIdsResponseForPublicFolder.ReturnValue != 0;

            this.Site.CaptureRequirementIfIsTrue(
                isVerifiedR469,
                469,
                @"[In RopGetPerUserLongTermIds ROP Request Buffer] This ROP [RopGetPerUserLongTermIds] MUST be issued against a logon that was made to a private mailbox.");
            #endregion Step7
        }

        /// <summary>
        /// This test case verifies the requirements related to the error codes when calling RopLongTermIdFromId ROP.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S01_TC16_TestRopLongTermIdFromIdErrorCodes()
        {
            this.CheckTransportIsSupported();

            #region Step 1: Connect the server via user configured by "AdminUserName".
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step 2: Logon to a private mailbox.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step 3: Call RopLongTermIdFromId ROP to get the LongTermId of the inbox.

            this.longTermIdFromIdRequest.ObjectId = this.logonResponse.FolderIds[4];
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];
            LongTermId longTermIdForInbox = this.longTermIdFromIdResponse.LongTermId;

            #endregion

            #region Step 4: Call RopIdFromLongTermId by setting global counter component in LongTermId to zeros.

            this.getIdFromLongTermIdRequest.LongTermId.DatabaseGuid = this.longTermIdFromIdResponse.LongTermId.DatabaseGuid;
            this.getIdFromLongTermIdRequest.LongTermId.GlobalCounter = new byte[this.longTermIdFromIdResponse.LongTermId.GlobalCounter.Length];
            this.oxcstorAdapter.DoRopCall(this.getIdFromLongTermIdRequest, this.outObjHandle, ROPCommandType.RopIdFromLongTermId, out this.outputBuffer);
            this.getIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R948");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R948
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getIdFromLongTermIdResponse.ReturnValue,
                948,
                @"[In Receiving a RopIdFromLongTermId ROP Request] If the LongTermId field of the request contains zeros for the global counter component, the server MUST fail the operation with 0 in the ReturnValue field.");
            #endregion

            #region Step 5: Call RopIdFromLongTermId by setting replica GUID component in LongTermId to zeros.
            this.getIdFromLongTermIdRequest.LongTermId.DatabaseGuid = new byte[this.longTermIdFromIdResponse.LongTermId.DatabaseGuid.Length];
            this.getIdFromLongTermIdRequest.LongTermId.GlobalCounter = this.longTermIdFromIdResponse.LongTermId.GlobalCounter;
            this.oxcstorAdapter.DoRopCall(this.getIdFromLongTermIdRequest, this.outObjHandle, ROPCommandType.RopIdFromLongTermId, out this.outputBuffer);
            this.getIdFromLongTermIdResponse = (RopIdFromLongTermIdResponse)this.outputBuffer.RopsList[0];

            if (Common.IsRequirementEnabled(94801001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R94801001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R94801001
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getIdFromLongTermIdResponse.ReturnValue,
                    94801001,
                    @"[In Appendix A: Product Behavior] When the LongTermId field of the request contains zeros for the replica GUID (REPLGUID) component, the implementation does return 0 in the ReturnValue field.  <44> Section 3.2.5.9:  Exchange 2010 returns 0 for this condition.");
            }

            if (Common.IsRequirementEnabled(94801002, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R94801002");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R94801002
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.getIdFromLongTermIdResponse.ReturnValue,
                    94801002,
                    @"[In Appendix A: Product Behavior] When the LongTermId field of the request contains zeros for the replica GUID (REPLGUID) component, the implementation does fail the operation with 0x80070057 (ecInvalidParam) in the ReturnValue field.  (Exchange 2007, Exchange 2013 and above follow this behavior).");

                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1120");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1120
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x80070057,
                    this.getIdFromLongTermIdResponse.ReturnValue,
                    1120,
                    @"[In Receiving a RopIdFromLongTermId ROP Request] The error code ecInvalidParam: Its value is 0x80070057.");
            }
            #endregion
        }
    }
}