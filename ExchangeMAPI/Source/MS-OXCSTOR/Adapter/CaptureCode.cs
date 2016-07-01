namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Server role and both role Adapter requirements capture code for MS-OXCSTOR.
    /// </summary>
    public partial class MS_OXCSTORAdapter : ManagedAdapterBase, IMS_OXCSTORAdapter
    {
        /// <summary>
        /// Verify the ROP runs on the underlying remote procedure call (RPC) transport
        /// </summary>
        private void VerifyRPC()
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R4");

            // If connect and execute through RPC call successfully, indicates the ROP response buffers are received from the server through RPC transport.
            this.Site.CaptureRequirement(
                4,
                @"[In Transport] The ROP response buffers specified by this protocol [MS-OXCSTOR] are received from the server by using the underlying Remote Operations (ROP) List and Encoding Protocol, as specified in [MS-OXCROPS].");
        }

        /// <summary>
        /// Verify transport.
        /// </summary>
        private void VerifyTransport()
        {
            string transport = Common.GetConfigurationPropertyValue("TransportSeq", this.Site);
            if (transport.ToLower() == "mapi_http" && Common.IsRequirementEnabled(300000001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R300000001");

                // Verify requirement MS-OXCSTOR_R300000001
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                    300000001,
                    @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }

            if ((transport.ToLower() == "ncacn_ip_tcp") && Common.IsRequirementEnabled(1911, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1911");

                // Verify requirement MS-OXCSTOR_R1911
                // If the transport sequence is ncacn_ip_tcp and the code can reach here, it means that the implementation does support ncacn_ip_tcp transport.
                Site.CaptureRequirement(
                    1911,
                    @"[In Appendix B: Product Behavior] Implementation does support this given protocol sequence [ncacn_ip_tcp]. ( Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopIdFromLongTermId
        /// </summary>
        /// <param name="request">The structure of ROP RopIdFromLongTermId request.</param>
        /// <param name="response">The structure of ROP RopIdFromLongTermId response.</param>
        private void VerifyRopIdFromLongTermId(RopIdFromLongTermIdRequest request, RopIdFromLongTermIdResponse response)
        {
            if (response.ReturnValue == 0)
            {
                ulong globalCounter = 0;
                for (int i = 0; i < 6; i++)
                {
                    globalCounter += ((ulong)request.LongTermId.GlobalCounter[i]) << (i * 8);
                }

                globalCounter = globalCounter << 16;

                // The last 2 bytes of the Folder ID or Message ID are not fixed.
                ulong maskedObjectId = response.ObjectId & ~(ulong)0xFFFF;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59721");

                // If the returned ObjectId is associated with the requested GlobalCounter, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirementIfAreEqual<ulong>(
                    globalCounter,
                    maskedObjectId,
                    59721,
                    @"[in RopIdFromLongTermId ROP] The RopIdFromLongTermId ROP ([MS-OXCROPS] section 2.2.3.9) is used to obtain the Folder ID or Message ID, as specified in [MS-OXCDATA] sections 2.2.4.1 or 2.2.4.2, given the LongTermID, as specified in [MS-OXCDATA] section 2.2.1.3.1.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R458");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R458
                this.Site.CaptureRequirementIfAreEqual<ulong>(
                    globalCounter,
                    maskedObjectId,
                    458,
                    @"[In RopIdFromLongTermId ROP Success Response Buffer] ObjectId: Contains the Folder ID or Message ID, as specified in [MS-OXCDATA] sections 2.2.1.1 or 2.2.1.2, that is mapped from the given LongTermID as specified in [MS-OXCDATA] section 2.2.1.3.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R459");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R459
                // The underlying structure for LongTermId is parsed as structure as REPLGUID(128-bit) with the specified sequence, if GlobalCounter is associated with the given short-term ID,this requirement can be verified.
                Site.CaptureRequirementIfAreEqual<ulong>(
                    globalCounter,
                    maskedObjectId,
                    459,
                    @"[In RopIdFromLongTermId ROP Success Response Buffer] ObjectId: The Folder ID or Message ID is a 64-bit value composed of the 16-bit replica ID (REPLID) followed by the 48-bit global counter portion of the given LongTermID.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1253");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1253
                this.Site.CaptureRequirementIfAreEqual<ulong>(
                    globalCounter,
                    maskedObjectId,
                    1253,
                    @"[In Receiving a RopIdFromLongTermId ROP Request] The Folder ID or Message ID consists of the 16-bit REPLID followed by the 48-bit global counter portion of the given LongTermID.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopLogon for the private mailbox.
        /// </summary>
        /// <param name="request">The structure of ROP RopLogon request.</param>
        /// <param name="response">The structure of ROP RopLogon response.</param>
        private void VerifyRopLogonForPrivateMailbox(RopLogonRequest request, RopLogonResponse response)
        {
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R13");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R13
                // If calling RopLogon ROP successfully, indicates the logon session is established.
                this.Site.CaptureRequirement(
                    13,
                    @"[In RopLogon ROP] The RopLogon ROP ([MS-OXCROPS] section 2.2.3.1) establishes a logon session between the client and the server.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R106");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R106
                // If calling RopLogon ROP successfully, indicates the Essdn is right identifier to log on to mailbox.
                this.Site.CaptureRequirement(
                    106,
                    @"[In RopLogon ROP Request Buffer] Essdn: In the case of a private mailbox logon, this field contains an ASCII string that uniquely identifies a mailbox to log on to.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R123");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R123
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    0x01,
                    (byte)(response.ResponseFlags & (byte)0x01),
                    123,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The description of flag Reserved: This bit [Reserved] MUST be set.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R122");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R122
                // Check if the flag Reserved is 0x01. 
                Site.CaptureRequirementIfAreEqual<byte>(
                    0x01,
                    (byte)(response.ResponseFlags & (byte)0x01),
                    122,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] [Response Flags] The value of flag Reserved: 0x01.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1062");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1062
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    request.LogonFlags,
                    response.LogonFlags,
                    1062,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] LogonFlags: The server returns these flags [Private, Undercover, and Ghostedflags] unchanged from the LogonFlags field of the RopLogon request (section 2.2.1.1.1).");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R118");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R118
                this.Site.CaptureRequirementIfAreEqual<byte>(
                    request.LogonFlags,
                    response.LogonFlags,
                    118,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] LogonFlags: This field is composed of the Private, Undercover, and Ghosted flags.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R117");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R117
                // The RopLogon ROP response structure for public folders is designated
                Site.CaptureRequirement(
                    117,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailbox] The following field values [LogonFlags, FolderIds, Response Flags, ReplId, ReplGuid, Logon Time, GwartTime, StoreState] are included in the RopLogon response only when the Private bit is set in the LogonFlags field of the RopLogon request (section 2.2.1.1.1).");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopLogon for the public folders.
        /// </summary>
        /// <param name="request">The structure of ROP RopLogon request.</param>
        /// <param name="response">The structure of ROP ROPLogonForPublicFolder response.</param>
        private void VerifyRopLogonForPublicFolder(RopLogonRequest request, RopLogonResponse response)
        {
            // The returnValue is 0 to indicate that the response is successful.
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R13");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R13
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    response.ReturnValue,
                    13,
                    @"[In RopLogon ROP] The RopLogon ROP ([MS-OXCROPS] section 2.2.3.1) establishes a logon session between the client and the server.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R146, LogonFlags:{0}", response.LogonFlags);

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R146
                bool isVerifyR146 = ((LogonFlags)response.LogonFlags & ~(LogonFlags.Private | LogonFlags.PublicFolder | LogonFlags.Ghosted | LogonFlags.Undercover)) == 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR146,
                    146,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] LogonFlags: This field [LogonFlags] is composed of the Private, Undercover, and Ghosted flags.");

                #region Capture 1071
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1071");

                // In the step2 logon a Public Folders.
                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1071
                Site.CaptureRequirementIfAreEqual<byte>(
                    request.LogonFlags,
                    response.LogonFlags,
                    1071,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] LogonFlags: The server returns these flags [Private, Undercover, and Ghosted] unchanged from the LogonFlags field of the RopLogon request (section 2.2.1.1.1).");
                #endregion

                Guid perUserGuid = new Guid(response.PerUserGuid);

                if (Common.IsRequirementEnabled(1332, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1332");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1332
                    this.Site.CaptureRequirementIfAreNotEqual<Guid>(
                        Guid.Empty,
                        perUserGuid,
                        1332,
                        @"[In Appendix A: Product Behavior] The implementation does not set the PerUserGuid field to an empty GUID. (<9> Section 2.2.1.1.4: Exchange 2007 does not set the PerUserGuid field to an empty GUID.)");
                }

                if (Common.IsRequirementEnabled(163, this.Site))
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R163");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R163
                    this.Site.CaptureRequirementIfAreEqual<Guid>(
                        Guid.Empty,
                        perUserGuid,
                        163,
                        @"Implementation does set this field [PerUserGuid in RopLogon ROP Success Response Buffer for Public Folders] to an empty GUID (all zeroes). (Exchange 2003, Exchange 2010 and above follow this behavior.)");
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R3012");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R3012
                // Check if the LogonFlag is set.
                bool isVerifyR3012 = (response.LogonFlags & (byte)LogonFlags.Private) == (byte)LogonFlags.Private;

                // The server returns the LogonFlags unchanged from the LogonFlag filed of the RopLogon request , So
                // if the response LgonFlag don't contain Private flag, it is not set in the request, then this requirement 
                // will be verified.
                Site.CaptureRequirementIfIsFalse(
                    isVerifyR3012,
                    3012,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] The success response buffer for public folders is sent only when the Private bit is not set in the LogonFlags field of the RopLogon request (section 2.2.1.1.1).");
            }
            else if (response.ReturnValue == 0x00000478)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R112");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R112
                bool isVerifyR112 = ((LogonFlags)response.LogonFlags & ~(LogonFlags.Private | LogonFlags.PublicFolder | LogonFlags.Ghosted | LogonFlags.Undercover)) == 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR112,
                    112,
                    @"[In RopLogon ROP Redirect Response Buffer] LogonFlags: This field contains the Private, Undercover, and Ghosted flags.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R115");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R115
                bool isVerifyR115 = Common.IsNullTerminatedASCIIStr(response.ServerName);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR115,
                    115,
                    @"[In RopLogon ROP Redirect Response Buffer] ServerName: The string includes the terminating NULL character.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopSetReceiveFolder.
        /// </summary>
        /// <param name="response">The structure of ROP RopSetReceiveFolder response.</param>
        private void VerifyRopSetReceiveFolder(RopSetReceiveFolderResponse response)
        {
            // The returnValue is 0 to indicate that the response is successful.
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R241");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirement(
                    241,
                    @"[In RopSetReceiveFolder ROP] RopSetReceiveFolder is used to establish the Receive folder for messages of a given message class.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopGetReceiveFolder.
        /// </summary>
        /// <param name="response">The structure of ROP RopGetReceiveFolder response.</param>
        private void VerifyRopGetReceiveFolder(RopGetReceiveFolderResponse response)
        {
            // The return value is 0 to indicate that the response is successful.
            if (response.ReturnValue == 0)
            {
                if (response.MessageClass != null)
                {
                    // If there is one character value in the string is not in the numeric range of 0 to 127,  this value will be false.   
                    bool isASCII = Common.IsNullTerminatedASCIIStr(response.MessageClass);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R212");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R212
                    // The method ToString is used to decode the stream with ASCII, so if the result is not null, this stream must use ASCII encoding.
                    Site.CaptureRequirementIfIsTrue(
                        isASCII,
                        212,
                        @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: The string uses ASCII encoding.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R213, the length of the MessageClass is: {0}.", response.MessageClass.Length);

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R213
                    // Check if the MessageClass string length is greater than zero and less than or equal to 255.
                    bool isVerifyR213 =
                        (response.MessageClass.Length > 0) &&
                        (response.MessageClass.Length <= 255);

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR213,
                        213,
                        @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: The length (including the terminating NULL character) is greater than zero and less than or equal to 255.");

                    // If MessageClass field only contains a terminating NULL character, ignore this requirement R214
                    if (response.MessageClass[0] != 0x00)
                    {
                        // If there is one character value in the string is not in the numeric range of 32 to 126,  this value will be false.
                        bool isValidateValue = true;

                        // The last element in MessageClass is a terminating NULL character, so the length minus one here. 
                        for (int i = 0; i < response.MessageClass.Length - 1; i++)
                        {
                            byte elem = response.MessageClass[i];

                            // Check if each character value in the string is in the numeric range of 32 to 126
                            if ((elem < 32) || (elem > 126))
                            {
                                isValidateValue = false;
                                break;
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R214");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R214
                        Site.CaptureRequirementIfIsTrue(
                            isValidateValue,
                            214,
                            @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: Each character value in the string is in the numeric range of 32 to 126, inclusive.");
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R215");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R215
                    // Check if the first byte in the byte array (response.MessageClass[0]) is not the character ".".
                    Site.CaptureRequirementIfAreNotEqual<byte>(
                       (byte)'.',
                        response.MessageClass[0],
                        215,
                        @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: The string does not begin with a period (""."").");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R216");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R216
                    // Check if the last byte in the byte array (response.MessageClass[response.MessageClass.Length - 1]) is not the character ".".
                    Site.CaptureRequirementIfAreNotEqual<byte>(
                        (byte)'.',
                        response.MessageClass[response.MessageClass.Length - 1],
                        216,
                        @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: The string does not end with a period.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R217, the MessageClass is: {0}.", response.MessageClass.ToString());

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R217
                    bool isVerifyR217 = true;

                    // Check if the element contained in the MessageClass buffer does not contain the adjacent period character "-".
                    foreach (byte elem in response.MessageClass)
                    {
                        if (elem == (byte)'-')
                        {
                            isVerifyR213 = false;
                            break;
                        }
                    }

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R217");

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR217,
                        217,
                        @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass: The string [Contained by ExplicitMessageClass] MUST meet the following requirement: The string does not contain adjacent periods.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R197");

                    // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                    Site.CaptureRequirement(
                        197,
                        @"[In RopGetReceiveFolder ROP] This ROP [RopGetReceiveFolder] also returns the specific parent message class configured to deliver to that folder.");
                }
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopGetReceiveFolderTable.
        /// </summary>
        /// <param name="response">The structure of ROP RopGetReceiveFolderTable response.</param>
        private void VerifyRopGetReceiveFolderTable(RopGetReceiveFolderTableResponse response)
        {
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R282");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirement(
                    282,
                    @"[In RopGetReceiveFolderTable ROP] The RopGetReceiveFolderTable is used to obtain a comprehensive list of all configured message classes and their associated Receive folders.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R297, the rows count:{0}", response.Rows == null ? 0 : response.Rows.Count);

                uint actualRowsCount = 0;

                // Check if server returns row information.
                if (response.Rows != null && response.Rows.Count > 0)
                {
                    // Get each row from the Rows structure.
                    for (int rowIndex = 0; rowIndex < response.Rows.Count; rowIndex++)
                    {
                        PropertyRow row = response.Rows.PropertyRows[rowIndex];
                        actualRowsCount++;

                        #region Verify each row's format

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R297, the flag is: {0}.", row.Flag);
                        bool isVerifiedR297 = row.Flag == 0x00 || row.Flag == 0x01;
                        Site.CaptureRequirementIfIsTrue(
                            isVerifiedR297,
                            297,
                            @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows: Each row is returned in either a StandardPropertyRow structure or a FlaggedPropertyRow structure, both of which [StandardPropertyRow, FlaggedPropertyRow] are specified in [MS-OXCDATA] section 2.8.1.1 and [MS-OXCDATA] section 2.8.1.2, respectively.");

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R877, the flag is: {0}.", row.Flag);
                        bool isVerifiedR877 = row.Flag == 0x00 || row.Flag == 0x01;

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R877
                        Site.CaptureRequirementIfIsTrue(
                            isVerifiedR877,
                            877,
                            @"[In Receiving a RopGetReceiveFolderTable ROP Request] The Rows field of the RopGetReceiveFolderTable response (section 2.2.1.4.2) contains either a StandardPropertyRow structure or a FlaggedPropertyRow structure for each row of the Receive folder table.");

                        // According to the description in [MS-OXCDATA], the structure of StandardPropertyRow and FlaggedPropertyRow 
                        // consists of a Flag field (1 byte) and a ValueArray field (byte[]). And the Row field which is contained in the Rows 
                        // field has been designed as the structure above in this test suite. So it's only need to verify the Flag field to 
                        // be 0x00 (which indicates a StandardPropertyRow structure), or 0x01(which indicates a FlaggedPropertyRow structure.).
                        #region StandardPropertyRow

                        // If this ValueArray is a StandardPropertyRow structure (0x00).
                        if (row.Flag == 0x00)
                        {
                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R299");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R299
                            this.Site.CaptureRequirementIfAreEqual<byte>(
                                0x00,
                                row.Flag,
                                299,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows: The value of each structure's [StandardPropertyRow structure, FlaggedPropertyRow] Flag field indicates which structure is being used: 0x00 for the StandardPropertyRow structure.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R879");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R879
                            Site.CaptureRequirementIfAreEqual<byte>(
                                0x00,
                                row.Flag,
                                879,
                                @"[In Receiving a RopGetReceiveFolderTable ROP Request] Otherwise [if there is not an error retrieving any data of a row from the Receive folder table], the server returns the row formatted as a StandardPropertyRow structure.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R301");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R301
                            // According to the description in [MS-OXCDATA], the structure of StandardPropertyRow and FlaggedPropertyRow consists of a Flag field (1 byte) and a ValueArray field (byte[]).
                            // And the Row field which is contained in the Rows field has been designed as the structure above in this test suite.
                            // So it's only need to verify the count of the structures that each row contains, then this requirement will be verified.
                            Site.CaptureRequirementIfAreEqual<int>(
                                3,
                                row.PropertyValues.Count,
                                301,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows: The ValueArray field of either StandardPropertyRow or FlaggedPropertyRow MUST include only the following properties[PidTagFolderId property, PidTagMessageClass property, PidTagLastModificationTime property], in the order given, and no other properties.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R72");

                            // Verify MS-OXCDATA requirement: MS-OXCDATA_R72
                            Site.CaptureRequirementIfAreEqual<Type>(
                                typeof(byte),
                                row.Flag.GetType(),
                                "MS-OXCDATA",
                                72,
                                @"[In StandardPropertyRow Structure] Flag (1 byte): An unsigned integer.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R73");

                            // If server returns 0x00 for the Flag field, which indicates that  all property
                            // values are present and without error, then this requirement will be verified.
                            Site.CaptureRequirement(
                                "MS-OXCDATA",
                                73,
                                @"[In StandardPropertyRow Structure] Flag (1 byte): This value MUST be set to 0x00 to indicate that all property values are present and without error.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R75");

                            // This requirement is verified using SDK, so it can be captured directly.
                            Site.CaptureRequirement(
                                "MS-OXCDATA",
                                75,
                                @"[In StandardPropertyRow Structure] ValueArray (variable): At each position of the array, the structure will either be a PropertyValue structure, as specified in section 2.11.2.1, if the type of the corresponding property tag was specified, or a TypedPropertyValue structure, as specified in section 2.11.3, if the type of the corresponding property tag was PtypUnspecified (section 2.11.1).");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R79");

                            // If server returns 0x00 for the Flag field, which indicates that  all property 
                            // values are present and without error, then this requirement will be verified.
                            Site.CaptureRequirement(
                                "MS-OXCDATA",
                                79,
                                @"[In FlaggedPropertyRow Structure] Flag (1 byte): Otherwise [when PtypUnspecified was not used in the ROP request and the ROP response includes a type], this value MUST be set to 0x00.");

                            // PidTagMessageClass property
                            byte[] messageClassBytes = row.PropertyValues[1].Value;
                            string messageClass = Encoding.ASCII.GetString(messageClassBytes);

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R305");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R305
                            // If the PidTagMessageClass property can be parsed as ASCII string, and not null, indicates it's the ASCII string.
                            this.Site.CaptureRequirementIfIsNotNull(
                                messageClass,
                                305,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property ([MS-OXPROPS] section 2.778): A PtypString8 value.");

                            // Add the debug information
                            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R307");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R307
                            this.Site.CaptureRequirementIfAreEqual<byte>(
                                0,
                                messageClassBytes[messageClassBytes.Length - 1],
                                307,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property:  The string [PidTagMessageClass] includes the terminating NULL character.");
                        }
                        #endregion

                        #region Verify each row's message class

                        // According to the  [MS-OXCSTOR] section 2.2.1.4.2, each row structure contains three elements: PidTagFolderId, PidTagMessageClass, PidTagLastModificationTime.
                        // Get the second array in the row structure(row.PropertyValues[1].value) which represents the PidTagMessageClass structure.
                        // If server returns MessageClass information for this row.
                        if (row.PropertyValues[1].Value != null)
                        {
                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R308");

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R308
                            Site.CaptureRequirementIfIsTrue(
                                Common.IsNullTerminatedASCIIStr(row.PropertyValues[1].Value),
                                308,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property: The string [PidTagMessageClass] MUST meet the following requirements: The string uses ASCII encoding.");

                            // Add the debug information
                            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R309, the size of the string is: {0}.", row.PropertyValues[1].Value.Length);

                            // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R309
                            bool isVerifyR309 =
                                row.PropertyValues[1].Value.Length > 0 && row.PropertyValues[1].Value.Length <= 255;

                            Site.CaptureRequirementIfIsTrue(
                                isVerifyR309,
                                309,
                                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property:  The string [PidTagMessageClass] MUST meet the following requirements: The length (including the terminating NULL character) is greater than zero and less than or equal to 255.");

                            // If the length of the "row.PropertyValues[1].value" is 1 ,the byte is terminator
                            if (row.PropertyValues[1].Value.Length != 1)
                            {
                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R310, the MessageClass string is: {0}.", row.PropertyValues[1].Value.ToString());

                                // If there is one character value in the string is not in the numeric range of 32 to 126, this value will be false.
                                bool isValidateValue = true;

                                // Get each character value in the string.
                                // The last one byte is terminator
                                for (int i = 0; i < row.PropertyValues[1].Value.Length - 1; i++)
                                {
                                    // Check if each character value in the string is in the numeric range of 32 to 126
                                    if ((row.PropertyValues[1].Value[i] < 32) || (row.PropertyValues[1].Value[i] > 126))
                                    {
                                        isValidateValue = false;
                                        break;
                                    }
                                }

                                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R310
                                Site.CaptureRequirementIfIsTrue(
                                    isValidateValue,
                                    310,
                                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property:   The string [PidTagMessageClass] MUST meet the following requirements: Each character value in the string is in the numeric range of 32 to 126, inclusive.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R311");

                                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R311
                                // Check the first character in the string (row.PropertyValues[1].Value[0]) is not ".".
                                Site.CaptureRequirementIfAreNotEqual<byte>(
                                    (byte)'.',
                                    row.PropertyValues[1].Value[0],
                                    311,
                                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property:  The string [PidTagMessageClass] MUST meet the following requirements: The string does not begin with a period (""."").");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R312");

                                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R312
                                // Check the last character in the string (row.PropertyValues[1].value[row.PropertyValues[1].value.Length - 1]) is not ".".
                                Site.CaptureRequirementIfAreNotEqual<byte>(
                                    (byte)'.',
                                    row.PropertyValues[1].Value[row.PropertyValues[1].Value.Length - 1],
                                    312,
                                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property: The string [PidTagMessageClass] MUST meet the following requirements: The string does not end with a period.");

                                // Add the debug information
                                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R313, the string is: {0}.", row.PropertyValues[1].Value.ToString());

                                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R313
                                // if the character in the string is "-", the value will be false.
                                bool isVerifyR313 = true;

                                // Get each character in the string.
                                foreach (byte elem in row.PropertyValues[1].Value)
                                {
                                    // If there is one character to be "-", R313 will be failed.
                                    if (elem == (byte)'-')
                                    {
                                        isVerifyR313 = false;
                                        break;
                                    }
                                }

                                Site.CaptureRequirementIfIsTrue(
                                    isVerifyR313,
                                    313,
                                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] [Rows] PidTagMessageClass property: The string [PidTagMessageClass] MUST meet the following requirements: The string does not contain adjacent periods.");
                            }
                        }

                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R301");

                        // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R301
                        // According to the description in [MS-OXCDATA], the structure of StandardPropertyRow and FlaggedPropertyRow consists of a Flag field (1 byte) and a ValueArray field (byte[]).
                        // And the Row field which is contained in the Rows field has been designed as the structure above in this test suite.
                        // So it's only need to verify the count of the structures that each row contains, then this requirement will be verified.
                        Site.CaptureRequirementIfAreEqual<int>(
                            3,
                            row.PropertyValues.Count,
                            301,
                            @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows: The ValueArray field of either StandardPropertyRow or FlaggedPropertyRow MUST include only the following properties[PidTagFolderId property, PidTagMessageClass property, PidTagLastModificationTime property], in the order given, and no other properties.");
                        #endregion
                        #endregion
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R296");

                    // If parse the Rows field as the array successfully, indicates the Rows field is consistent with its description.
                    this.Site.CaptureRequirement(
                        296,
                        @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows: An array that contains the rows of the Receive folder table.");
                }

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R294");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R294
                this.Site.CaptureRequirementIfAreEqual<uint>(
                    response.RowCount,
                    actualRowsCount,
                    294,
                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] RowCount: The number of rows in the table.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopGetOwningServers.
        /// </summary>
        /// <param name="response">The structure of ROP RopGetOwningServers response.</param>
        private void VerifyRopGetOwningServers(RopGetOwningServersResponse response)
        {
            // The this.returnValue is 0 to indicate that the response is successful.
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R343");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirement(
                    343,
                    @"[In RopGetOwningServers ROP] RopGetOwningServers is used to obtain the set of servers that host content for a replicated public folder.");

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCSTOR_R367, the CheapServersCount is: {0}, and the OwningServersCount is: {1}.",
                    response.CheapServersCount,
                    response.OwningServersCount);

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R367
                Site.CaptureRequirementIfIsTrue(
                    response.CheapServersCount <= response.OwningServersCount,
                    367,
                    @"[In RopGetOwningServers ROP Success Response Buffer] CheapServersCount: This value [CheapServersCount] MUST be less than or equal to OwningServersCount.");

                if (response.OwningServersCount > 0)
                {
                    // Add the debug information
                    this.Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-OXCSTOR_R368, the CheapServersCount is: {0}, and the OwningServersCount is: {1}.",
                        response.CheapServersCount,
                        response.OwningServersCount);

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R368
                    bool isVerifiedR368 = response.CheapServersCount > 0;

                    this.Site.CaptureRequirementIfIsTrue(
                        isVerifiedR368,
                        368,
                        @"[In RopGetOwningServers ROP Success Response Buffer] CheapServersCount: This value [CheapServersCount] MUST be greater than zero if OwningServersCount is greater than zero.");
                }

                bool isServerNameWithNullTerminated = true;
                foreach (string serverName in response.OwningServers)
                {
                    if (serverName[serverName.Length - 1] != '\0')
                    {
                        isServerNameWithNullTerminated = false;
                        break;
                    }
                }

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R369");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R369
                this.Site.CaptureRequirementIfIsTrue(
                    isServerNameWithNullTerminated,
                    369,
                    @"[In RopGetOwningServers ROP Success Response Buffer] OwningServers: Contains an array of null-terminated ASCII strings.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R371");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R371
                Site.CaptureRequirementIfAreEqual<int>(
                    (int)response.OwningServersCount,
                    response.OwningServers.Length,
                    371,
                    @"[In RopGetOwningServers ROP Success Response Buffer] OwningServers: The number of strings MUST be equal to the value specified in the OwningServersCount field.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R351");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R351
                this.Site.CaptureRequirementIfAreEqual<int>(
                    (int)response.OwningServersCount,
                    response.OwningServers.Length,
                    351,
                    @"[In RopGetOwningServers ROP Success Response Buffer] OwningServersCount: Identifies the number of strings contained in the OwningServers field.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopPublicFolderIsGhosted.
        /// </summary>
        /// <param name="response">The structure of ROP RopPublicFolderIsGhosted response.</param>
        private void VerifyRopPublicFolderIsGhosted(RopPublicFolderIsGhostedResponse response)
        {
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R396");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirement(
                    396,
                    @"[In RopPublicFolderIsGhosted ROP] The RopPublicFolderIsGhosted is used to obtain the replication state for a folder on the current server.");

                if (response.IsGhosted == 0x00)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1100");

                    // This block is to verify the requirements when the IsGhosted field is set to FALSE.  
                    // If the response can be parsed successfully when the IsGhosted field is set to FALSE, indicates the field[CheapServersCount] is not present
                    this.Site.CaptureRequirement(
                        1100,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] CheapServersCount: This field [CheapServersCount] is not present if the IsGhosted field is set to FALSE.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1101");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1101
                    // If the Servers field is not present, the Servers will be null.
                    this.Site.CaptureRequirementIfIsNull(
                        response.Servers,
                        1101,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] Servers: This field[Servers] is not present if the IsGhosted field is set to FALSE");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R59712");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R59712
                    // If the Servers field is not present, the Servers will be null.
                    this.Site.CaptureRequirementIfIsNull(
                        response.Servers,
                        59712,
                        @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] Servers: This field is not present if the IsGhosted field is set to FALSE.");
                }
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopLongTermIdFromId
        /// </summary>
        /// <param name="request">The structure of ROP RopLongTermIdFromId request.</param>
        /// <param name="response">The structure of ROP RopLongTermIdFromId response.</param>
        private void VerifyRopLongTermIdFromId(RopLongTermIdFromIdRequest request, RopLongTermIdFromIdResponse response)
        {
            if (response.ReturnValue == 0)
            {
                ulong globalCounter = 0;
                for (int i = 0; i < 6; i++)
                {
                    globalCounter += (ulong)(((ulong)response.LongTermId.GlobalCounter[i]) << (i * 8));
                }

                globalCounter = globalCounter << 16;

                // The last 2 bytes of the Folder ID or Message ID are not fixed.
                ulong maskedObjectId = request.ObjectId & ~(ulong)0xFFFF;

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R69742");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                // The returned LongTermID is associated with the given a Folder ID or Message ID.
                this.Site.CaptureRequirementIfAreEqual<ulong>(
                    maskedObjectId,
                    globalCounter,
                    69742,
                    @"[In RopLongTermIdFromId ROP] The RopLongTermIdFromId ROP ([MS-OXCROPS] section 2.2.3.8) is used to obtain a LongTermID structure, as specified in [MS-OXCDATA] section 2.2.1.3.1, given a Folder ID structure or Message ID structure, as specified in [MS-OXCDATA] section 2.2.1.1 or section 2.2.1.2. ");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R440");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R440 
                // If the high 48 bits in ObjectId and GlobalCounter are same, indicates the LongTermId related to the requested REPLID is valid in the REPLID and REPLGUID to-and-from mapping table.
                Site.CaptureRequirementIfAreEqual<ulong>(
                    maskedObjectId,
                    globalCounter,
                    440,
                    @"[In RopLongTermIdFromId ROP Request Buffer] ObjectId: The 16-bit REPLID portion of the Folder ID or Message ID MUST be a valid entry in the REPLID and REPLGUID to-and-from mapping table.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R442");

                // The LongTermId field is verified in the RopLongTermIdFromIdResponse structure.
                // If the LongTermId field is parsed successfully, this requirement can be captured directly.
                Site.CaptureRequirementIfIsNotNull(
                    response.LongTermId,
                    442,
                    @"[In RopLongTermIdFromId ROP Success Response Buffer] LongTermId: Contains the LongTermID structure, as specified in [MS-OXCDATA] section 2.2.1.3.1.");

                // The longTermId is null indicates its invalid object ID.
                if (response.LongTermId.DatabaseGuid != null)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R1244");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R1244
                    // The underlying structure for LongTermId is parsed as structure as REPLGUID(128-bit) with the specified sequence, if GlobalCounter is associated with the given short-term ID,this requirement can be verified.
                    Site.CaptureRequirementIfAreEqual<ulong>(
                        maskedObjectId,
                        globalCounter,
                        1244,
                        @"[In Receiving a RopLongTermIdFromId ROP Request] The LongTermID consists of the 128-bit REPLGUID, followed by the 48-bit global counter portion of the given Folder ID or Message ID, followed by 16 bits of padding set to 0x0000.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2188");

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2188
                    // The LongTermID is a GID, which is verified in "MS-OXCDATA", so here only need to verify the length of Pad field.
                    bool isVerifyR2188 = response.LongTermId.Size() == 24;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR2188,
                        "MS-OXCDATA",
                        2188,
                        @"[In LongTermID Structure] A LongTermID structure is a Global Identifier structure, as specified in section 2.2.1.3, plus a 2-byte Pad field that has the value 0x0000.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2189");

                    // The longtermID consists of DatabaseGuid, GlobalCounter and Padding, so its length is the sum of the three.
                    int longTermIdLength = response.LongTermId.DatabaseGuid.Length +
                        response.LongTermId.GlobalCounter.Length + 2; // +response.LongTermId.Padding.Length;

                    // Verify MS-OXCDATA requirement: MS-OXCDATA_R2189
                    Site.CaptureRequirementIfAreEqual<int?>(
                        24,
                        longTermIdLength,
                        "MS-OXCDATA",
                        2189,
                        @"[In LongTermID Structure] The total length of the LongTermID structure is 24 bytes.");
                }
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopGetPerUserLongTermIds.
        /// </summary>
        /// <param name="response">The structure of ROP RopGetPerUserLongTermIds response.</param>
        private void VerifyRopGetPerUserLongTermIds(RopGetPerUserLongTermIdsResponse response)
        {
            if (response.ReturnValue == 0)
            {
                if (response.LongTermIds != null && response.LongTermIds.Length > 0)
                {
                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R468");

                    // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description. 
                    this.Site.CaptureRequirement(
                        468,
                        @"[In RopGetPerUserLongTermIds ROP] The RopGetPerUserLongTermIds ROP ([MS-OXCROPS] section 2.2.3.10) is used to obtain the LongTermIDs (as specified in [MS-OXCDATA] section 2.2.1.3.1) of folders in a public folders message store that contain per-user read/unread data identified by a REPLGUID.");

                    // Add the debug information
                    this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R474");

                    // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R474
                    // If the response can be parsed successfully and the return value is success, this requirement can be verified directly.
                    this.Site.CaptureRequirement(
                        474,
                        @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIds: Contains an array of LongTermID structures, as specified in [MS-OXCDATA] section 2.2.1.3.1, of folders in the public message store for which this user has cached read/unread information.");
                }

                bool isLongTermIdCountConsistent = ((response.LongTermIds == null) && (response.LongTermIdCount == 0)) ||
                    ((response.LongTermIds != null) && (response.LongTermIdCount == (ushort)response.LongTermIds.Length));

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R475, the LongTermIdCount is: {0}.", response.LongTermIdCount);

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R475
                bool isVerify475 = isLongTermIdCountConsistent;
                Site.CaptureRequirementIfIsTrue(
                    isVerify475,
                    475,
                    @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIds: The number of items in this array MUST be the same as the value of the LongTermIdCount field (section 2.2.1.10.2).");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R472, the LongTermIdCount is: {0}.", response.LongTermIdCount);

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R472
                bool isVerify472 = isLongTermIdCountConsistent;
                Site.CaptureRequirementIfIsTrue(
                    isVerify472,
                    472,
                    @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIdCount: Specifies the number of entries in the following array.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopGetPerUserGuid.
        /// </summary>
        /// <param name="response">The structure of ROP RopGetPerUserGuid response.</param>
        private void VerifyRopGetPerUserGuid(RopGetPerUserGuidResponse response)
        {
            if (response.ReturnValue == 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R484");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R484
                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description. 
                this.Site.CaptureRequirement(
                    484,
                    @"[In RopGetPerUserGuid ROP] The RopGetPerUserGuid ROP ([MS-OXCROPS] section 2.2.3.11) obtains the REPLGUID of the public message store that previously provided the now cached per-user read/unread data for a specific public folder.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopReadPerUserInformation.
        /// </summary>
        /// <param name="response">The structure of ROP RopReadPerUserInformation response.</param>
        private void VerifyRopReadPerUserInformation(RopReadPerUserInformationResponse response)
        {
            // If server returns Data, change numbers associated with a message are returned.
            if (response.Data != null && response.DataSize != 0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R502");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description. 
                this.Site.CaptureRequirement(
                    502,
                    @"[In RopReadPerUserInformation ROP] RopReadPerUserInformation is used to obtain a set of change numbers, each of which is associated with a message that the user has read in a specific public folder.");

                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R542");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R542
                this.Site.CaptureRequirementIfAreEqual<int>(
                    response.Data.Length,
                    (int)response.DataSize,
                    542,
                    @"[In RopReadPerUserInformation ROP Success Response Buffer] DataSize: Contains the size, in bytes, of the data being returned.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R546");

                // Verify MS-OXCSTOR requirement: MS-OXCSTOR_R546
                Site.CaptureRequirementIfAreEqual<int>(
                    (int)response.DataSize,
                    response.Data.Length,
                    546,
                    @"[In RopReadPerUserInformation ROP Request Buffer] Data: The size of the BLOB MUST be equal to the value specified in the DataSize field.");
            }
        }

        /// <summary>
        /// Verify the response by sending the ROP RopWritePerUserInformation.
        /// </summary>
        /// <param name="response">The structure of ROP RopWritePerUserInformation response.</param>
        private void VerifyRopWritePerUserInformation(RopWritePerUserInformationResponse response)
        {
            if (response.ReturnValue == 0x0)
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCSTOR_R560");

                // If the response can be parsed successfully and the return value is success, indicates the ROP's functionality is consistent with its description.
                this.Site.CaptureRequirement(
                    560,
                    @"[In RopWritePerUserInformation ROP] RopWritePerUserInformation is used to establish the set of change numbers of messages the user has read in a specific public folder.");
            }
        }
    }
}