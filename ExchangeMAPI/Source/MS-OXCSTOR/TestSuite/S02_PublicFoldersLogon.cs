namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// MS-OXCSTOR test suite for testing the public folder
    /// </summary>
    [TestClass]
    public class S02_PublicFoldersLogon : TestSuiteBase
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
        /// This test case verifies the requirements of logging on to a public folder.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC01_TestLogonToPublicFolder()
        {
            this.CheckTransportIsSupported();

            Site.Assume.IsFalse(this.transportByMAPIOrNot, "This case doesn't support the mapi_http transport.");
            
            #region Step1: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step2:  RopLogon ROP log on to public folder with the Ghosted flag

            // Set Ghosted flag
            this.logonRequestForPublicFolder.LogonFlags = (byte)LogonFlags.Ghosted;

            // Set the IGNORE_HOME_MDB flag
            this.logonRequestForPublicFolder.OpenFlags = this.logonRequestForPublicFolder.OpenFlags | (uint)OpenFlags.IgnoreHomeMDB;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse1 = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, logonResponse1.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            #region Capture

            if (Common.IsRequirementEnabled(5979301, this.Site))
            {
                // Do not check NNTP Article Index and EForms Registry for the user's locale
                // NNTP Article Index: It is supported on Exchange 2003 and below version.
                // EForms Registry for the user's locale: Needs to create the Organizational Forms Library 
                //     1. New-PublicFolder -Path "\NON_IPM_SUBTREE\EFORMS REGISTRY" -Name "Organizational Forms Library"
                //     2. Download MFCMAPI tool and add property PR_EFORMS_LOCALE_ID
                //     3. Set-PublicFolder "\NON_IPM_SUBTREE\EFORMS REGISTRY\Organizational Forms Library" -EformsLocaleID EN-US
                bool isR5979301Verified = logonResponse1.FolderIds[0] != 0 // Public Folders Root Folder
                    && logonResponse1.FolderIds[1] != 0 // Interpersonal messages subtree
                    && logonResponse1.FolderIds[2] != 0 // Non-interpersonal messages subtree
                    && logonResponse1.FolderIds[3] != 0 // EForms Registry
                    && logonResponse1.FolderIds[4] != 0 // Free/Busy Data
                    && logonResponse1.FolderIds[5] != 0 // Offline address book Data
                    && logonResponse1.FolderIds[7] != 0 // Local Site's Free/Busy Data
                    && logonResponse1.FolderIds[8] != 0 // Local Site's Offline Address Book Data
                    && logonResponse1.FolderIds[10] == 0 // Empty
                    && logonResponse1.FolderIds[11] == 0 // Empty
                    && logonResponse1.FolderIds[12] == 0; // Empty

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R5979301");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R5979301
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonResponse1.ReturnValue,
                    5979301,
                    @"[In Appendix A: Product Behavior]  Implementation returns the Folder ID of all of the following folders in this field [FolderIds in RopLogon ROP Success Response Buffer for Public Folders]: Public Folders Root Folder (All other folders listed here are direct or indirect children of this folder), Interpersonal messages subtree, Non-interpersonal messages subtree, EForms Registry, Free/Busy Data, Offline address book Data, EForms Registry for the user's locale, Local Site's Free/Busy Data, Local Site's Offline Address Book Data, NNTP Article Index, Empty, Empty, Empty. (Exchange 2003, Exchange 2007 and Exchange 2010 follow this behavior)");
            }

            if (Common.IsRequirementEnabled(155, this.Site))
            {
                bool isR155Verified = logonResponse1.FolderIds[0] != 0 // Public Folders Root Folder
                    && logonResponse1.FolderIds[1] != 0 // Interpersonal messages subtree
                    && logonResponse1.FolderIds[2] != 0 // Non-interpersonal messages subtree
                    && logonResponse1.FolderIds[3] != 0 // EForms Registry
                    && logonResponse1.FolderIds[4] == 0 // Free/Busy Data
                    && logonResponse1.FolderIds[5] == 0 // Offline address book Data
                    && logonResponse1.FolderIds[7] == 0 // Local Site's Free/Busy Data
                    && logonResponse1.FolderIds[8] == 0 // Local Site's Offline Address Book Data
                    && logonResponse1.FolderIds[9] == 0 // NNTP Article Index
                    && logonResponse1.FolderIds[10] == 0 // Empty
                    && logonResponse1.FolderIds[11] == 0 // Empty
                    && logonResponse1.FolderIds[12] == 0; // Empty

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R155");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R155
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonResponse1.ReturnValue,
                    155,
                    @"[In Appendix A: Product Behavior] The implementation returns the empty Folder ID structures for the following folders: Free/Busy Data, Offline Address Book Data, Local Site's Free/Busy Data, Local Site's Offline Address Book Data, and NNTP Article Index. (<8> Section 2.2.1.1.4: Exchange 2013 and Exchange 2016 returns the empty Folder ID structures for the following folders: Free/Busy Data, Offline Address Book Data, Local Site's Free/Busy Data, Local Site's Offline Address Book Data, and NNTP Article Index.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R7202");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R7202
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                logonResponse1.ReturnValue,
                7202,
                @"[In RopLogon ROP Request Buffer] [OpenFlags] The description of flag PUBLIC: A request to open a public folders message store.");

            if (Common.IsRequirementEnabled(1266001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1266001");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1266001
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonResponse1.ReturnValue,
                    1266001,
                    @"[In Appendix A: Product Behavior] The implementation requires passing 0x00 in the EssdnSize field for public folder logon and that the Essdn field be empty. (<6> Section 2.2.1.1.1: Exchange 2003, Exchange 2007, and Exchange 2010 require passing 0x00 in the EssdnSize field for public folder logon and that the Essdn field be empty.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3110001");

            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                logonResponse1.ReturnValue,
                3110001,
                @"[In Public Folders Logon] The OpenFlags field MUST have the PUBLIC bit set to log on to the public folder container (for either database or mailbox).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R786");

            Site.Assert.AreNotEqual<uint>((uint)LogonFlags.Private, (uint)(logonResponse1.LogonFlags & (uint)LogonFlags.Private), "Check the Private flag is not set");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R786
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                logonResponse1.ReturnValue,
                786,
                @"[In Receiving a RopLogon ROP Request] Otherwise [if the LogonFlags field has not the Private bit set], the logon is going to the public folders.");
            #endregion

            this.longTermIdFromIdRequest.ObjectId = (ulong)BitConverter.ToUInt16(logonResponse1.ReplId, 0);
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(0x00000000, this.longTermIdFromIdResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R148");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R148
            Site.CaptureRequirementIfAreEqual<string>(
                logonResponse1.ReplGuid.ToString(),
                this.longTermIdFromIdResponse.LongTermId.DatabaseGuid.ToString(),
                148,
                @"[In RopLogon ROP Success Response Buffer for Public Folders] ReplId: This field [ReplId] contains the short form of the value specified in the ReplGuid field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R153");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R153
            this.Site.CaptureRequirementIfAreEqual<string>(
                this.longTermIdFromIdResponse.LongTermId.DatabaseGuid.ToString(),
                logonResponse1.ReplGuid.ToString(),
                153,
                @"[In RopLogon ROP Success Response Buffer for Public Folders] ReplGuid: It [ReplGuid] is used in forming LongTermIDs structures, as defined in [MS-OXCDATA] section 2.2.1.3.1.");

            #endregion

            #region Step3: RopLogon ROP to log on to public folder without the Ghosted flag
            this.logonRequestForPublicFolder.LogonFlags = 0x00;

            // OpenFlags is set USE_PER_MDB_REPLID_MAPPING and Public flags
            this.logonRequestForPublicFolder.OpenFlags = (uint)OpenFlags.UsePerMDBReplipMapping | (uint)OpenFlags.Public;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            this.oxcstorAdapter.SetAutoRedirect(false);

            #region Step4: Call RopLogon ROP to log on to a public folder with ALTERNATE_SERVER bit set

            // Set ALTERNATE_SERVER and PUBLIC flags
            this.logonRequestForPublicFolder.OpenFlags = (uint)OpenFlags.AlternateServer | (uint)OpenFlags.Public;

            // The Ghosted flag is not set
            this.logonRequestForPublicFolder.LogonFlags = (byte)LogonFlags.PublicFolder;
            RopLogonRequest logonRequest_ALTERNATE_SERVER = this.logonRequestForPublicFolder;
            this.oxcstorAdapter.DoRopCall(logonRequest_ALTERNATE_SERVER, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse_ALTERNATESERVER_NOGHOST = (RopLogonResponse)this.outputBuffer.RopsList[0];

            // Set the Ghosted flag
            this.logonRequestForPublicFolder.LogonFlags = (byte)LogonFlags.Ghosted;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse_ALTERNATESERVER_GHOST = (RopLogonResponse)this.outputBuffer.RopsList[0];

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R837");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R837
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000478,
                logonResponse_ALTERNATESERVER_GHOST.ReturnValue,
                837,
                @"[In Public Folders Logon] Otherwise [If a suitable server can be found], the operation MUST fail with a ReturnValue of 0x00000478 (ecWrongServer).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R838001");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R838001
            // 1 The Ghosted flag is not set, the response of the server is logonResponse_ALTERNATESERVER_NOGHOST.
            // 2 The Ghosted flag is set, the response of the server is logonResponse_ALTERNATESERVER_GHOST.
            // 3 Compare if logonResponse_ALTERNATESERVER_NOGHOST and logonResponse_ALTERNATESERVER_GHOST are the same, 
            // if the responses are the same MS-OXCSTOR_R838001 is verified.
            bool isVerifyR838001 = this.CompareLogonResponse(logonResponse_ALTERNATESERVER_NOGHOST, logonResponse_ALTERNATESERVER_GHOST);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR838001,
                838001,
                @"[In Public Folders Logon] If the OpenFlags field has the ALTERNATE_SERVER bit set, the server will return the same response no matter whether the Ghosted bit of the LogonFlags field is set or not.");
            #endregion capture
            #endregion

            #region Step5: Call RopLogon ROP to log on to a public folder without ALTERNATE_SERVER bit set
            this.logonRequestForPublicFolder.OpenFlags = (uint)OpenFlags.Public;

            // The Ghost bit flag is not set.
            this.logonRequestForPublicFolder.LogonFlags = 0x00;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse logonResponse_NOALTERNATESERVER = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0x00000000, logonResponse_NOALTERNATESERVER.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.oxcstorAdapter.SetAutoRedirect(true);

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R833");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R833
            // Set the user who has a mailbox on server 1, but with the default public folder data base on server 2.
            // If server search for another public folder database server in the organization when the bit  ALTERNATE_SERVER bit is set in the OpenFlags field, next scenario will occur:
            // 1, Log on to the server 1 with ALTERNATE_SERVER bit set, server will return the server name to indicate user the valid server to log on.
            // 2, Log on to the server 2 without ALTERNATE_SERVER bit set, server will not return the server name.
            // MS_OXCSTOR_R833 is verified if the ServerName filed is not null.
            bool isALTERNATESERVERConsistent = (logonResponse_ALTERNATESERVER_GHOST.ServerName != null) && (logonResponse_NOALTERNATESERVER.ServerName == null);
            bool isVerifyR833 = isALTERNATESERVERConsistent;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR833,
                833,
                @"[In Public Folders Logon] If the OpenFlags field has the ALTERNATE_SERVER bit set, the server searches for another public folder container server in the organization which is not the configured preferred server.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R90");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R90
            bool isVerifyR90 = isALTERNATESERVERConsistent;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR90,
                90,
                @"[In RopLogon ROP Request Buffer] [OpenFlags] The description of flag ALTERNATE_SERVER: Requests a private server to provide an alternate public server.");

            #endregion capture

            #endregion

            #region Step6: Call RopLogon ROP to log on to a public folder with set OpenFlags to USE_PER_MDB_REPLID_MAPPING
            // Set PUBLIC and USE_PER_MDB_REPLID_MAPPING
            this.logonRequestForPublicFolder.OpenFlags = (uint)OpenFlags.Public | (uint)OpenFlags.UsePerMDBReplipMapping;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

            // OpenFlags is not set USE_PER_MDB_REPLID_MAPPING and set PUBLIC flags
            this.logonRequestForPublicFolder.OpenFlags = (uint)OpenFlags.Public;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            RopLogonResponse ropResponse_NO_USER_PER_MDB_REPLID_MAPPING = (RopLogonResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(0x00000000, ropResponse_NO_USER_PER_MDB_REPLID_MAPPING.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3540");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3540
            // In step 6, set the USE_PER_MDB_REPLID_MAPPING flag, and the response is logonResponse.
            // In capture R3540 not set the USE_PER_MDB_REPLID_MAPPING flag, and the response is ropResponse_NO_USER_PER_MDB_REPLID_MAPPING.
            // So, need to check whether their responses are the same.
            bool isVerifyR3540 = this.CompareLogonResponse(this.logonResponse, ropResponse_NO_USER_PER_MDB_REPLID_MAPPING);
            Site.CaptureRequirementIfIsTrue(
                isVerifyR3540,
                3540,
                @"[In RopLogon ROP Request Buffer] [OpenFlags] The description of flag USE_PER_MDB_REPLID_MAPPING: For logons to a public folder message store, the server reply is the same, no matter whether the flag is set or not.");
            #endregion
            #endregion
        }

        /// <summary>
        /// This test case verifies the logon redirect response.
        /// This test case depends on second SUT. If second SUT is not present, this test case cannot be executed.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC02_TestLogonWithecWrongServer()
        {
            this.CheckTransportIsSupported();

            Site.Assume.IsFalse(this.transportByMAPIOrNot, "This case doesn't support the mapi_http transport.");
            
            if (this.secondSUTExistOrNot)
            {
                this.oxcstorAdapter.SetAutoRedirect(false);
                #region Step1: Connect to server
                string server1 = this.server1Name;
                string userDN = this.essdnOfUser2 + "\0";
                string domain = this.domain;
                string password = this.passwordOfUserOfMailboxOnServer2;

                this.returnStatus = this.oxcstorAdapter.ConnectEx(server1, ConnectionType.PublicFolderServer, userDN, domain, this.userNameOfMailboxOnServer2, password);
                Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
                #endregion

                #region Step2: Call RopLogon ROP to log on to the public folder
                this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
                this.oxcstorAdapter.SetAutoRedirect(true);
                this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
                #endregion

                #region Capture

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1222");

                Site.Assume.AreNotEqual<uint>((uint)OpenFlags.AlternateServer, (uint)(this.logonRequestForPublicFolder.OpenFlags & (uint)OpenFlags.AlternateServer), "Check if the ALTERNATESERVER is not set.");
                Site.Assume.AreNotEqual<uint>((uint)OpenFlags.IgnoreHomeMDB, (uint)(this.logonRequestForPublicFolder.OpenFlags & (uint)OpenFlags.IgnoreHomeMDB), "Check if the IGNORE_HOME_MDB is not set.");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1222
                Site.CaptureRequirementIfAreEqual<uint>(
                    0x00000478,
                    this.logonResponse.ReturnValue,
                    1222,
                    @"[In Public Folders Logon] Otherwise [If the OpenFlags field has neither the ALTERNATE_SERVER bit nor the IGNORE_HOME_MDB bit set], the server has the following behavior: If the server does not host that container, the server MUST fail the operation with a ReturnValue of 0x00000478 (ecWrongServer).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R116");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R116
                // If the logonResponse serverNameSize field equals the actual ServerName length the MS-OXCSTOR_R116 is verified.
                Site.CaptureRequirementIfAreEqual<byte>(
                    (byte)logonResponse.ServerName.Length,
                    this.logonResponse.ServerNameSize,
                    116,
                    @"[In RopLogon ROP Redirect Response Buffer] ServerName: The string length (including the terminating NULL character) MUST be equal to the value specified by the ServerNameSize field.");

                if (Common.IsRequirementEnabled(1266002, this.Site))
                {
                    // Use the same connection and enable auto redirect, expect RopLogon ROP return success
                    this.oxcstorAdapter.SetAutoRedirect(true);
                    this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
                    RopLogonResponse logonResponseUseEssdn = (RopLogonResponse)this.outputBuffer.RopsList[0];

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R114");

                    // In the step 2, RopLogon ROP to log on to the public folder with a redirect response logonResponse.
                    // Use the ESSDN from the logonResponse ServerName field to Logon a public folder, if the logonResponseUserEssdn
                    // return value is 0 it indicates the ESSDN from the redirect response is a valid ESSDN for the client to connect to.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        0,
                        logonResponseUseEssdn.ReturnValue,
                        114,
                        @"[In RopLogon ROP Redirect Response Buffer] ServerName: This field contains the enterprise/site/server distinguished name (ESSDN) of server for the client to connect to, as the server included in the request no longer hosts the requested mailbox (it was moved), or was the wrong server to connect to for access to public folders.");
                }
                #endregion capture
            }
            else
            {
                Site.Assert.Inconclusive("This case runs only when the second system under test exists.");
            }
        }

        /// <summary>
        /// This test case verifies that the operation about public folder is ghosted.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC03_TestPublicFolderIsGhosted()
        {
            this.CheckTransportIsSupported();

            #region Step1: connect
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            this.oxcstorAdapter.SetAutoRedirect(true);
            #endregion

            #region Step2: RopLogon ROP to log on to the private mailbox
            this.logonRequestForPrivateMailBox.LogonFlags = (byte)LogonFlags.Private;
            this.logonRequestForPrivateMailBox.OpenFlags = 0x01000000;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step3: RopPublicFolderIsGhosted ROP against the private mailbox logon
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[0];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            this.publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.publicFolderIsGhostedResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R917");

            Site.CaptureRequirementIfAreEqual<byte>(
                0x00,
                this.publicFolderIsGhostedResponse.IsGhosted,
                917,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] If the operation is issued against a private mailbox message store, the server MUST return FALSE in the IsGhosted field of the response.");

            // The Servers field of the RopPublicFolderIsGhosted ROP response is the replica information
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1237");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1237
            Site.CaptureRequirementIfIsNull(
                this.publicFolderIsGhostedResponse.Servers,
                1237,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] In this case [if the operation is issued against a private mailbox store, the server MUST return FALSE in the IsGhosted field of the response], no replication state data is returned.");

            #endregion

            #region Step4: Disconnect server1
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
            #endregion

            #region Step5: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region Step6: RopLogon ROP to log on to the public folder
            this.logonRequestForPublicFolder.OpenFlags = 0x01000202;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step7: RopPublicFolderIsGhosted ROP against the public folder logon
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[0];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);

            RopPublicFolderIsGhostedResponse publicFolderIsGhostedResponse1;
            publicFolderIsGhostedResponse1 = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.publicFolderIsGhostedResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion this.publicFolderIsGhostedRequest

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R405");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R405
            // The default setting (after setup the test environment) is that the queried server is an active replica of the folder.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                publicFolderIsGhostedResponse1.IsGhosted,
                405,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] otherwise [if the server is an active replica of the folder the IsGhosted flag] is FALSE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R930");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R930
            // The default setting (after setup the test environment) is that the queried server is an active replica of the folder,  
            // Check if "IsGhost" is false to indirectly verify the queried server is listed.
            Site.CaptureRequirementIfAreEqual<byte>(
                0,
                publicFolderIsGhostedResponse1.IsGhosted,
                930,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The value of the IsGhosted field MUST be FALSE if the queried server is listed as an active replica of the folder.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R941");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R941
            bool isVerifiedR941 = (publicFolderIsGhostedResponse1.Servers == null && publicFolderIsGhostedResponse1.ServersCount == null) ||
                (publicFolderIsGhostedResponse1.Servers != null && publicFolderIsGhostedResponse1.ServersCount == publicFolderIsGhostedResponse1.Servers.Length);

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR941,
                941,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request]  The current total list length constitutes the ServersCount return value.");
            #endregion

            #region Step8: RopPublicFolderIsGhosted ROP with invalid FolderId
            // A non-existing folder
            this.publicFolderIsGhostedRequest.FolderId = 0x01;
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            RopPublicFolderIsGhostedResponse publicFolderIsGhostedInvalidResponse;
            publicFolderIsGhostedInvalidResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];
            #endregion

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1236");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1236
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                publicFolderIsGhostedInvalidResponse.ReturnValue,
                1236,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] If the public folder specified by the FolderId field cannot be found in the public folder container (either public folder database or public folder mailbox), the server MUST fail the operation with 0x8004010F (ecNotFound) in the ReturnValue field.");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R436
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                publicFolderIsGhostedInvalidResponse.ReturnValue,
                436,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The error code ecNotFound: Its value is 0x8004010F.");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R437
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                publicFolderIsGhostedInvalidResponse.ReturnValue,
                437,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] The error code ecNotFound: The FID could not be found in the public folder container.");

            #endregion

            #region Step9:  RopPublicFolderIsGhosted ROP with the public folder = the IPM subtree root folders
            // According to the section 2.2.1.1.4 in the Open Specification, FolderIds[1] is Interpersonal Messages Sub-tree(IPM)
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[1];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            this.publicFolderIsGhostedResponse = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.publicFolderIsGhostedResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Step10: RopPublicFolderIsGhosted ROP with the public folder = non-IPM subtree
            // According to the section 2.2.1.1.4 in the Open Specification, FolderIds[2] is Non-interpersonal messages subtree(non-IPM)
            this.publicFolderIsGhostedRequest.FolderId = this.logonResponse.FolderIds[2];
            this.oxcstorAdapter.DoRopCall(this.publicFolderIsGhostedRequest, this.outObjHandle, ROPCommandType.RopPublicFolderIsGhosted, out this.outputBuffer);
            RopPublicFolderIsGhostedResponse pubFolderIsGhostResponse_non_IPM = (RopPublicFolderIsGhostedResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                pubFolderIsGhostResponse_non_IPM.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1239");

            // In step 9 the public folder is an IPM subtree and the response is publicFolderIsGhostedResponse.
            // When public folder is changed into a Non-interpersonal messages subtree, the response is pubFolderIsGhostResponse_non_IPM.
            // If both publicFolderIsGhostedResponse IsGhosted and pubFolderIsGhostResponse_non_IPM IsGhosted fields are 0, then MS-OXCSTOR_R1239 is verified.
            bool isVerifyR1239 = pubFolderIsGhostResponse_non_IPM.IsGhosted == 0 && this.publicFolderIsGhostedResponse.IsGhosted == 0;

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1239
            Site.CaptureRequirementIfIsTrue(
                isVerifyR1239,
                1239,
                @"[In Receiving a RopPublicFolderIsGhosted ROP Request] If the client issues this operation against the IPM subtree or the non-IPM subtree public folders, the value of the IsGhosted field MUST be FALSE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1098");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1098
            // ServersCount is not present means the value of it is null.
            Site.Assert.AreEqual(0, this.publicFolderIsGhostedResponse.IsGhosted, "Check if the IsGhosted is 0(FALSE)");

            Site.CaptureRequirementIfIsNull(
                this.publicFolderIsGhostedResponse.ServersCount,
                1098,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ServersCount: This field [ServersCount] is not present if the IsGhosted field is set to FALSE.");
            #endregion
        }

        /// <summary>
        /// This test case verifies the operations that are invalid when the Logon object refers to a public folder logon.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC04_TestInvalidOperationsLogonPublicfolders()
        {
            this.CheckTransportIsSupported();

            #region step1: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PrivateMailboxServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region step2: RopLogon to log on to the private mailbox

            this.oxcstorAdapter.DoRopCall(this.logonRequestForPrivateMailBox, this.insideObjHandle, ROPCommandType.RopLogonPrivateMailbox, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];

            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Capture

            // Local Site's Free/Busy Data Folder
            this.getOwningServersRequest.FolderId = this.logonResponse.FolderIds[7];
            this.oxcstorAdapter.DoRopCall(this.getOwningServersRequest, this.outObjHandle, ROPCommandType.RopGetOwningServers, out this.outputBuffer);
            this.getOwningServersResponse = (RopGetOwningServersResponse)this.outputBuffer.RopsList[0];

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R892");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R892
            Site.CaptureRequirementIfAreNotEqual<uint>(
                0,
                getOwningServersResponse.ReturnValue,
                892,
                @"[In Receiving a RopGetOwningServers ROP Request] If the operation is performed against a private mailbox message store, the server can fail the operation, or it can compute a correct answer for the client.");
            #endregion

            #region step3: RopGetReceiveFolder ROP
            this.getReceiveFolderRequest.MessageClass = Encoding.ASCII.GetBytes("IPM\0");
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            RopGetReceiveFolderResponse getReceiveFolderForPrivateMailboxResponse;
            getReceiveFolderForPrivateMailboxResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            #endregion RopGetReceiveFolder ROP

            #region step4: RopSetReceiveFolder ROP
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[4];
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            RopSetReceiveFolderResponse setReceiveFolderForPrivateMailboxResponse;
            setReceiveFolderForPrivateMailboxResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            #endregion RopSetReceiveFolder ROP

            #region step5: RopGetReceiveFolderTable ROP
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderTableRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolderTable, out this.outputBuffer);
            RopGetReceiveFolderTableResponse getReceiveFolderTableForPrivateMailboxResponse = (RopGetReceiveFolderTableResponse)this.outputBuffer.RopsList[0];
            #endregion RopGetReceiveFolderTable ROP

            #region Step6: RopGetPerUserLongTermIds ROP
            this.longTermIdFromIdRequest.ObjectId = this.logonResponse.FolderIds[0];
            this.oxcstorAdapter.DoRopCall(this.longTermIdFromIdRequest, this.outObjHandle, ROPCommandType.RopLongTermIdFromId, out this.outputBuffer);
            this.longTermIdFromIdResponse = (RopLongTermIdFromIdResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.longTermIdFromIdResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

            this.getPerUserLongTermIdsRequest.DatabaseGuid = this.longTermIdFromIdResponse.LongTermId.DatabaseGuid;

            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsForPrivateMailboxResponse;
            getPerUserLongTermIdsForPrivateMailboxResponse = (RopGetPerUserLongTermIdsResponse)this.outputBuffer.RopsList[0];

            Site.Assert.AreEqual<uint>(
                0x00000000,
                getPerUserLongTermIdsForPrivateMailboxResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion RopGetPerUserLongTermIds ROP

            #region Step7: Disconnect
            this.returnStatus = this.oxcstorAdapter.DisconnectEx();
            Site.Assert.IsTrue(this.returnStatus, "Disconnection is successful");
            #endregion

            #region step8: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region step9: RopLogon ROP to public folders
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region step10: RopGetReceiveFolder ROP
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolder, out this.outputBuffer);
            #endregion RopGetReceiveFolder ROP

            RopGetReceiveFolderResponse getReceiveFolderForPublicFoldersResponse;
            getReceiveFolderForPublicFoldersResponse = (RopGetReceiveFolderResponse)this.outputBuffer.RopsList[0];

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59782");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59782
            // In the step 3, call RopGetReceiveFolder for a private mailbox and in the step 10 call RopGetReceiveFolder for public folders.
            // If successful in the step 3 and failed in the step 8, MS-OXCSTOR_R59782 can be verified.
            bool isVerifyR59782 = getReceiveFolderForPrivateMailboxResponse.ReturnValue == 0 && getReceiveFolderForPublicFoldersResponse.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR59782,
                59782,
                @"[In Receiving a RopGetReceiveFolder ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
            #endregion

            #region step11: RopSetReceiveFolder ROP
            this.setReceiveFolderRequest.FolderId = this.logonResponse.FolderIds[4];
            this.oxcstorAdapter.DoRopCall(this.setReceiveFolderRequest, this.outObjHandle, ROPCommandType.RopSetReceiveFolder, out this.outputBuffer);
            RopSetReceiveFolderResponse setReceiveFolderForPublicFoldersResponse;
            setReceiveFolderForPublicFoldersResponse = (RopSetReceiveFolderResponse)this.outputBuffer.RopsList[0];
            #endregion RopSetReceiveFolder ROP

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R855");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R855
            // In the step 4, call RopSetReceiveFolder for a private mailbox and in the step 11 call RopSetReceiveFolder for public folders.
            // If successful in the step 4 and failed in the step 11, this requirement can be verified.
            bool isVerify_R855 = setReceiveFolderForPrivateMailboxResponse.ReturnValue == 0 && setReceiveFolderForPublicFoldersResponse.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R855,
                855,
                @"[In Receiving a RopSetReceiveFolder ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");
            #endregion

            #region step12: RopGetReceiveFolderTable ROP
            this.oxcstorAdapter.DoRopCall(this.getReceiveFolderTableRequest, this.outObjHandle, ROPCommandType.RopGetReceiveFolderTable, out this.outputBuffer);
            RopGetReceiveFolderTableResponse getReceiveFolderTableForPublicFosderResponse;
            getReceiveFolderTableForPublicFosderResponse = (RopGetReceiveFolderTableResponse)this.outputBuffer.RopsList[0];
            #endregion RopGetReceiveFolderTable ROP

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R874");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R874
            // In the step 5, call RopGetReceiveFolderTable for a private mailbox and in the step 12 call RopGetReceiveFolderTable for public folders.
            // If successful in the step 5 and failed in the step 12, this requirement can be verified.
            bool isVerify_R874 = getReceiveFolderTableForPrivateMailboxResponse.ReturnValue == 0 && getReceiveFolderTableForPublicFosderResponse.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerify_R874,
                874,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] The server verifies that the operation is being performed against a private mailbox logon, and not a public folders logon.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R875");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R875
            Site.CaptureRequirementIfAreEqual<uint>(
                0x80040102,
                getReceiveFolderTableForPublicFosderResponse.ReturnValue,
                875,
                @"[In Receiving a RopGetReceiveFolderTable ROP Request] If the operation is performed against a public folders logon, the server MUST fail the operation with a ReturnValue of 0x80040102.");
            #endregion

            #region step13: RopGetPerUserLongTermIds ROP against the public folder logon
            this.getPerUserLongTermIdsRequest.DatabaseGuid = new byte[16];
            this.oxcstorAdapter.DoRopCall(this.getPerUserLongTermIdsRequest, this.outObjHandle, ROPCommandType.RopGetPerUserLongTermIds, out this.outputBuffer);
            RopGetPerUserLongTermIdsResponse getPerUserLongTermIdsForPublicFoldersResponse;
            getPerUserLongTermIdsForPublicFoldersResponse = (RopGetPerUserLongTermIdsResponse)this.outputBuffer.RopsList[0];
            #endregion RopGetPerUserLongTermIds ROP

            #region Capture
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59770");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59770
            // Step 6 calls RopGetPerUserLongTermIds for a private mailbox and step 13 calls RopGetPerUserLongTermIds for public folders.
            // If successful in the step 6 and failed in the step 13, this requirement can be verified.
            bool isVerifyR59770 = getPerUserLongTermIdsForPrivateMailboxResponse.ReturnValue == 0 && getPerUserLongTermIdsForPublicFoldersResponse.ReturnValue != 0;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR59770,
                59770,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] The server verifies that the operation is being performed against a private mailbox logon, not against a public folders logon.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59771");

            // Step 13 calls RopGetPerUserLongTermIds for public folders.
            // If the return value is 0x80040102, R59771 is verified.
            Site.CaptureRequirementIfAreEqual(
                0x80040102,
                getPerUserLongTermIdsForPublicFoldersResponse.ReturnValue,
                59771,
                @"[In Receiving a RopGetPerUserLongTermIds ROP Request] If the operation is performed against a public folders logon, the server MUST fail the operation with a ReturnValue of 0x80040102.");

            #endregion
        }

        /// <summary>
        /// This test case tests ROP buffers of RopGetOwningServers.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC05_TestRopGetOwningServers()
        {
            this.CheckTransportIsSupported();

            string publicFolderName = Common.GenerateResourceName(Site, "PublicFolder");

            #region step1: Connect to server1
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            #endregion

            #region step2: Call RopLogon ROP to log on public folders
            // Set Public USE_PER_MDB_REPLID_MAPPING IGNORE_HOME_MDB flags
            this.logonRequestForPublicFolder.OpenFlags = 0x01000202;
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion
            
            #region step3: Open a public folder
            RopOpenFolderRequest openFolderRequest;
            RopOpenFolderResponse openFolderResponse;

            openFolderRequest.RopId = 0x02;
            openFolderRequest.LogonId = 0x0;
            openFolderRequest.InputHandleIndex = 0x0;
            openFolderRequest.OutputHandleIndex = 0x01;

            // Root folder
            openFolderRequest.FolderId = this.logonResponse.FolderIds[1];

            // Opening an existing folder
            openFolderRequest.OpenModeFlags = 0x0;

            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            uint openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];
            #endregion

            #region step4: Create a public folder
            string publicFolder = publicFolderName;
            string ghostPublicFolder = publicFolder + "\0";
            RopCreateFolderRequest createFolderRequest;
            RopCreateFolderResponse createFolderResponse;
            createFolderRequest.RopId = 0x1C;
            createFolderRequest.LogonId = 0x0;
            createFolderRequest.InputHandleIndex = 0x0;
            createFolderRequest.OutputHandleIndex = 0x01;
            createFolderRequest.FolderType = 0x01; // Generic folder
            createFolderRequest.UseUnicodeStrings = 0x0; // FALSE
            createFolderRequest.OpenExisting = 0xFF; // non-zero(TRUE)
            createFolderRequest.Reserved = 0x0; // FALSE
            createFolderRequest.DisplayName = Encoding.ASCII.GetBytes(ghostPublicFolder);
            createFolderRequest.Comment = Encoding.ASCII.GetBytes(ghostPublicFolder);
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            ulong folderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                createFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region step5: Call RopGetOwningServers ROP request with the existed public folder
            this.getOwningServersRequest.FolderId = folderId;
            this.oxcstorAdapter.DoRopCall(this.getOwningServersRequest, this.outObjHandle, ROPCommandType.RopGetOwningServers, out this.outputBuffer);
            this.getOwningServersResponse = (RopGetOwningServersResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                this.getOwningServersResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region Capture
            if (Common.IsRequirementEnabled(3490001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3490001");
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    this.getOwningServersResponse.ReturnValue,
                    3490001,
                    @"[In Appendix A: Product Behavior] Implementation does issue this operation [RopGetOwningServers] against a public folders logon. (Exchange 2013 follows this behavior.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R915");

            Site.Assert.AreNotEqual<int>(0, this.getOwningServersResponse.OwningServersCount, "Check if the number of the OwningServers field is not 0");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R915
            Site.CaptureRequirementIfIsNotNull(
                this.getOwningServersResponse.OwningServers,
                915,
                "[In Receiving a RopGetOwningServers ROP Request] The list contents of server identifiers constitute the value in the OwningServers field.");

            if (Common.IsRequirementEnabled(1266002, this.Site))
            {
                RopLogonRequest logonRequestForPublicFolderUseResponseESSDN;
                logonRequestForPublicFolderUseResponseESSDN = this.logonRequestForPublicFolder;
                logonRequestForPublicFolderUseResponseESSDN.EssdnSize = (ushort)this.getOwningServersResponse.OwningServers[0].Length;
                logonRequestForPublicFolderUseResponseESSDN.Essdn = System.Text.Encoding.ASCII.GetBytes(this.getOwningServersResponse.OwningServers[0]);
                RopOutputBuffer ropBufUseEssdn;
                this.oxcstorAdapter.DoRopCall(logonRequestForPublicFolderUseResponseESSDN, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out ropBufUseEssdn);
                RopLogonResponse logonResponseUseEssdn = (RopLogonResponse)ropBufUseEssdn.RopsList[0];
                this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];

                // Use the ESSDN from the logonResponse ServerName field to Logon a public folder, if the logonResponseUserEssdn
                // return value is 0 it indicates the ESSDN from the GetOwningServer response is a valid ESSDN for the client to connect to.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1266002");

                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonResponseUseEssdn.ReturnValue,
                    1266002,
                    @"[In Appendix A: Product Behavior] In the case of a public folder logon, this field [Essdn] contains an ASCII string that uniquely identifies the server that contains the public folder mailbox to log on to. (Exchange 2013 follows this behavior)");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R370");

                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    logonResponseUseEssdn.ReturnValue,
                    370,
                    @"[In RopGetOwningServers ROP Success Response Buffer] OwningServers: Each string is the ESSDN of a public folder database that hosts an active replica of the content of the folder.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1061");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1061
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    logonRequestForPublicFolderUseResponseESSDN.LogonFlags,
                    logonResponseUseEssdn.LogonFlags,
                    1061,
                    @"[In RopLogon ROP Redirect Response Buffer] LogonFlags: The server returns these flags unchanged from the LogonFlags field of the RopLogon request (section 2.2.1.1.1).");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R914");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_914
            Site.CaptureRequirementIfAreEqual<int>(
                this.getOwningServersResponse.OwningServersCount,
                this.getOwningServersResponse.OwningServers.Length,
                914,
                @"[In Receiving a RopGetOwningServers ROP Request] The current total list length constitutes the OwningServersCount value returned in the response.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R390");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R390
            // The error code ecNone indicates the operation is implemented successfully, so if the ReturnValue
            // is 0, MS-OXCSTOR_R390 is verified with valid FID.
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getOwningServersResponse.ReturnValue,
                390,
                @"[In Receiving a RopGetOwningServers ROP Request] The error code ecNone: Its value is 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R391");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R391
            Site.CaptureRequirementIfAreEqual<uint>(
                0,
                this.getOwningServersResponse.ReturnValue,
                391,
                @"[In Receiving a RopGetOwningServers ROP Request] The error code ecNone: Success.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R916");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R916 
            // Check if the returned server's identifier moniker map into userDN's string
            string expectedMoniker = this.userDN.Remove(this.userDN.IndexOf("/cn="));
            bool isVerifiedR916 = this.getOwningServersResponse.OwningServers[0].ToUpper().Contains(expectedMoniker.ToUpper());

            Site.CaptureRequirementIfIsTrue(
                isVerifiedR916,
                916,
                @"[In Receiving a RopGetOwningServers ROP Request] The server MUST map whatever identifier moniker for each server it has into an ESSDN string to return to the client.");

            #endregion

            #region step6: RopGetOwningServers ROP
            this.oxcstorAdapter.DoRopCall(this.getOwningServersRequest, this.outObjHandle, ROPCommandType.RopGetOwningServers, out this.outputBuffer);
            RopGetOwningServersResponse getOwningServersResponse1;
            getOwningServersResponse1 = (RopGetOwningServersResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0,
                getOwningServersResponse1.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            #endregion

            #region step7: RopGetOwningServers ROP with invalid FID
            // Set an invalid FID
            this.getOwningServersRequest.FolderId = 0x1;
            this.oxcstorAdapter.DoRopCall(this.getOwningServersRequest, this.outObjHandle, ROPCommandType.RopGetOwningServers, out this.outputBuffer);
            this.getOwningServersResponse = (RopGetOwningServersResponse)this.outputBuffer.RopsList[0];
            #endregion

            #region Capture

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1232");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1232
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                this.getOwningServersResponse.ReturnValue,
                1232,
                @"[In Receiving a RopGetOwningServers ROP Request] If the public folder specified by the FolderId field cannot be found in the public folder database, the server MUST fail the operation with 0x8004010F (ecNotFound) in the ReturnValue field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R394");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R394
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                this.getOwningServersResponse.ReturnValue,
                394,
                @"[In Receiving a RopGetOwningServers ROP Request] The error code ecNotFound: Its value is 0x8004010F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R395");

            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R395
            Site.CaptureRequirementIfAreEqual<uint>(
                0x8004010F,
                this.getOwningServersResponse.ReturnValue,
                395,
                @"[In Receiving a RopGetOwningServers ROP Request] The error code ecNotFound: The FID could not be found in the public folder database.");
            #endregion

            #region step8: Softly delete the folder
            // Log on to the public folder.
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(0, this.logonResponse.ReturnValue, "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            openFolderRequest.FolderId = this.logonResponse.FolderIds[1];
            openFolderRequest.InputHandleIndex = 0x00;
            openFolderRequest.OutputHandleIndex = 0x01;

            // Opening an existing folder
            openFolderRequest.OpenModeFlags = 0x0;
            this.oxcstorAdapter.DoRopCall(openFolderRequest, this.outObjHandle, ROPCommandType.Others, out this.outputBuffer);

            openFolderResponse = (RopOpenFolderResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                openFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            openedFolderHandle = this.outputBuffer.ServerObjectHandleTable[openFolderRequest.OutputHandleIndex];

            // Get the folder Id.
            this.oxcstorAdapter.DoRopCall(createFolderRequest, openedFolderHandle, ROPCommandType.Others, out this.outputBuffer);
            createFolderResponse = (RopCreateFolderResponse)this.outputBuffer.RopsList[0];
            folderId = createFolderResponse.FolderId;
            Site.Assert.AreEqual<uint>(
                0x00000000,
                createFolderResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");

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
        /// This test case tests the operations of Property PidTagUserEntryId.
        /// </summary>
        [TestCategory("MSOXCSTOR"), TestMethod()]
        public void MSOXCSTOR_S02_TC06_TestOperationsOnProperties()
        {
            this.CheckTransportIsSupported();

            #region Step1: connect
            this.returnStatus = this.oxcstorAdapter.ConnectEx(ConnectionType.PublicFolderServer);
            Site.Assert.IsTrue(this.returnStatus, "Connection is successful");
            this.oxcstorAdapter.SetAutoRedirect(true);
            #endregion

            #region Step2: Call RopLogon ROP to log on to the public folder
            this.oxcstorAdapter.DoRopCall(this.logonRequestForPublicFolder, this.insideObjHandle, ROPCommandType.RopLogonPublicFolder, out this.outputBuffer);
            this.logonResponse = (RopLogonResponse)this.outputBuffer.RopsList[0];
            Site.Assert.AreEqual<uint>(
                0x00000000,
                this.logonResponse.ReturnValue,
                "0 indicates the ROP succeeds, other value indicates error occurs.");
            this.outObjHandle = this.outputBuffer.ServerObjectHandleTable[0];
            #endregion

            #region Step3: test PidTagUserEntryId

            #region Variable
            PropertyTag propertyTagUserEntryId;

            // According to [MS-OXPROPS], PidTagUserEntryId's id is 0x6619
            const ushort PropertyIdForUserEntryId = 0x6619;
            #endregion

            #region Read a readable property PidTagUserEntryId
            propertyTagUserEntryId.PropertyId = PropertyIdForUserEntryId;

            // According to [MS-OXPROPS], PidTagUserEntryId's data type is 0x0102
            propertyTagUserEntryId.PropertyType = 0x0102;
            uint responseFlag;
            this.TryGetFlagLogonProperty(propertyTagUserEntryId, out responseFlag);
            #endregion

            #region Capture
            
            if (Common.IsRequirementEnabled(3077, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3077");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3077
                // When the operation failed, the server return error.
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    responseFlag,
                    3077,
                    @"[In PidTagUserEntryId Property] Type: PtypBinary ([MS-OXCDATA] section 2.11.1)");
            }
            #endregion

            #endregion
        }
    }
}