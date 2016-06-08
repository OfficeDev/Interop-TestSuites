namespace Microsoft.Protocols.TestSuites.MS_OXCROPS
{
    using System;
    using System.Collections.Generic;
    using System.Runtime.InteropServices;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Adapter requirements capture code for MS-OXCROPS server role.
    /// </summary>
    public partial class MS_OXCROPSAdapter
    {
        #region Verify transport 
        
        /// <summary>
        /// Verify transport.
        /// </summary>
        private void VerifyTransport()
        {
            if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http" && Common.IsRequirementEnabled(5000, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R5000");

                // Verify requirement MS-OXCROPS_R5000
                // If the transport sequence is MAPIHTTP and the code can reach here, it means that the implementation does support MAPIHTTP transport.
                Site.CaptureRequirement(
                    5000,
                    @"[In Appendix A: Product Behavior] Implementation does support this specification [MS-OXCMAPIHTTP]. (Exchange Server 2013 Service Pack 1 (SP1) follows this behavior.)");
            }
            else if (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "ncacn_ip_tcp" && Common.IsRequirementEnabled(5001, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R5001");

                // Verify requirement MS-OXCROPS_R5001
                // If the transport sequence is ncacn_ip_tcp and the code can reach here, it means that the implementation does support ncacn_ip_tcp transport.
                Site.CaptureRequirement(
                    5001,
                    @"[In Appendix B: Product Behavior] Implementation does support this given protocol sequence [ncacn_ip_tcp]. ( Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
            }
        }
        #endregion

        #region Message Syntax both role
        /// <summary>
        /// Verify each ROP request and response buffer structure 
        /// </summary>
        /// <param name="ropSize">The size of Request buffer or Response buffer</param>
        /// <param name="ropList">The RopList of request or response</param>
        /// <param name="serverOHTable">The ROP server object handle table of request or response</param>
        /// <param name="rawData">The ROP response payload</param>
        private void VerifyMessageSyntaxRequestAndResponseBuffer(object ropSize, byte[][] ropList, List<uint> serverOHTable, byte[] rawData)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R6");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R6
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropSize.GetType(),
                6,
                @"[In ROP Input and Output Buffers] RopSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R11");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R11
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(List<uint>),
                serverOHTable.GetType(),
                11,
                @"[In ROP Input and Output Buffers] ServerObjectHandleTable (variable): An array of 32-bit values.");

            int serverOHTableLength = serverOHTable.Count * sizeof(uint);

            // The actual bytes count of ServerObjectHandleTable. RPC_HEADER_EXT_LEN:the length of RPC_HEADER_EXT
            int actualsOHTableLength = rawData.Length - (ushort)ropSize - RPCHEADEREXTLEN;

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4675");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4675
            Site.CaptureRequirementIfAreEqual<int>(
                serverOHTableLength,
                actualsOHTableLength,
                4675,
                @"[In ROP Input and Output Buffers] ServerObjectHandleTable (variable):The size of this field is equal to the number of bytes of data remaining in the ROP Input and Output Buffers after the RopsList field.");

            // Mark byte count of ropList
            ushort ropsBytesCount = 0;
            foreach (byte[] bytes in ropList)
            {
                // Get byte count of ropList
                ropsBytesCount += (ushort)bytes.Length;
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4673");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4673
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)ropSize,
                (ushort)(ropsBytesCount + 2),
                4673,
                @"[In ROP Input and Output Buffers] RopsList (variable): The size of this field is 2 bytes less than the value specified in the RopSize field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4671");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4671
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[][]),
                ropList.GetType(),
                4671,
                @"[In ROP Input and Output Buffers] RopsList (variable): For a ROP output buffer, this field contains an array of ROP response buffers.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3
            // If the previous requirements can be verified correctly, that is, the response message is parsed in little-endian byte order, this requirement can verified directly.
            Site.CaptureRequirement(
                3,
                @"[In Messages] Unless otherwise specified, buffers and fields in this section are depicted in little-endian byte order.");
        }
        #endregion

        #region Verify error for the RPC
        /// <summary>
        /// Verify RPC error for Reserved RopIds
        /// </summary>
        /// <param name="ropId">The field of RopId in request</param>
        /// <param name="status">The status of the RPC operation</param>
        private void VerifyRPCErrorEncounterReservedRopIds(byte ropId, uint status)
        {
            // Check whether the ropId is one of the reserved RopIds.
            bool isReservedRopId = this.oxcropsClient.IsReservedRopId(ropId);

            // If RopId is associated with a reserved ROP 
            // And the RPC status does not equal to 0 (0 represents success response) and the transport is RPC, then verify this requirement.
            bool isVerifyR213 = isReservedRopId && (status != 0);
            if (Common.IsRequirementEnabled(213, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R213");

                // Verify requirement MS-MS-OXCROPS_R213
                Site.CaptureRequirementIfIsTrue(
                        isVerifyR213,
                        213,
                        @"[In Appendix B: Product Behavior] Implementation does return an error for the RPC ,as specified in [MS-OXCRPC] section 3.1.4.12, when it encounters a RopId value that is associated with a reserved ROP. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            // If RopId is associated with a reserved ROP 
            // And the returned status does not equal to 0 (0 represents success response) and the transport is mapi_http, then verify this requirement.
            if (Common.IsRequirementEnabled(180005, this.Site) && (Common.GetConfigurationPropertyValue("TransportSeq", this.Site).ToLower() == "mapi_http"))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R180005");

                // Verify requirement MS-MS-OXCROPS_R213
                Site.CaptureRequirementIfIsTrue(
                        isReservedRopId && (status != 0),
                        180005,
                        @"[In Appendix A: Product Behavior] Implementation does return an error for the Execute request type when the server encounters a RopId value that is associated with a reserved ROP. (<4> Section 2.2.2: Exchange 2013 SP1 and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify RPC error when server is unable to parse the ROP requests in the input ROP buffer
        /// </summary>
        /// <param name="status">The status of the RPC operation</param>
        private void VerifyRPCErrorEncounterUnableParseRequest(uint status)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4551");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4551
            // The return value is 0x000004B6 indicating the server is unable to parse the ROP request.
            Site.CaptureRequirementIfAreEqual<uint>(
                OxcRpcErrorCode.ECRpcFormat,
                status,
                4551,
                @"[In Processing a ROP Input Buffer] If the server is unable to parse the ROP requests in the  ROP input buffer, the RPC MUST fail by returning 0x000004B6.");
        }
        #endregion

        #region Logon ROPs

        #region RopLogon
        /// <summary>
        /// Verify RopLogon Success Response
        /// </summary>
        /// <param name="ropLogonResponse">The response of the RopLogon request</param>
        private void VerifyRopLogonSuccessResponse(RopLogonResponse ropLogonResponse)
        {
            byte logonFlags = ropLogonResponse.LogonFlags;

            // This bit is set for logon to a private mailbox and is not set for logon to public folders.
            if (logonFlags == (byte)LogonFlags.Private)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R121,LogonFlags:{0}", ropLogonResponse.LogonFlags);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R121
                bool isVerifyR121 = ((ropLogonResponse.LogonFlags & 0x01) == (byte)LogonFlags.Private)
                                    || ((ropLogonResponse.LogonFlags & 0x02) == (byte)LogonFlags.Undercover)
                                    || ((ropLogonResponse.LogonFlags & 0x04) == (byte)LogonFlags.Ghosted);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR121,
                    121,
                    @"[In RopLogon ROP Redirect Response Buffer] LogonFlags (1 byte): The possible values[the value of Private is 0x01, the value of Undercover is 0x02, the value of Ghosted is 0x04] are specified in [MS-OXCSTOR] Section 2.2.1.1.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R43");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R43
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.RopId.GetType(),
                    43,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] RopId (1 byte): ): An unsigned integer .");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R45");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R45
                Site.CaptureRequirementIfAreEqual<byte>(
                    (byte)RopId.RopLogon,
                    ropLogonResponse.RopId,
                    45,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] RopId (1 byte): For this operation[RopLogon], this field[RopId (1 byte)] is set to 0xFE.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R46");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R46
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.OutputHandleIndex.GetType(),
                    46,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] OutputHandleIndex (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R48");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R48
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    ropLogonResponse.ReturnValue.GetType(),
                    48,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ReturnValue (4 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R50");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R50
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropLogonResponse.ReturnValue,
                    50,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ReturnValue (4 bytes): For this response[Success Response Buffer for Private Mailboxes], this field is set to 0x00000000.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R51");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R51
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonFlags.GetType(),
                    51,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] LogonFlags (1 byte): ): A flags structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R70,LogonFlags:{0}", ropLogonResponse.LogonFlags);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R70
                bool isVerifyR70 = ((ropLogonResponse.LogonFlags & (byte)LogonFlags.Private) == (byte)LogonFlags.Private)
                                   || ((ropLogonResponse.LogonFlags & (byte)LogonFlags.Undercover) == (byte)LogonFlags.Undercover)
                                   || ((ropLogonResponse.LogonFlags & (byte)LogonFlags.Ghosted) == (byte)LogonFlags.Ghosted);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR70,
                    70,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] LogonFlags (1 byte): The possible values[the value of Private is 0x01, the value of Undercover is 0x02 and the value of Ghosted is 0x04] are specified in [MS-OXCSTOR] Section2.2.1.1.3.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R53, the actual size of element in FolderIds is {0} and length of FolderIds is {1}", Marshal.SizeOf(ropLogonResponse.FolderIds[0]), ropLogonResponse.FolderIds.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R53
                // If size of element in FolderIds is 8 byte and the length of  FolderIds is 13, then verify this requirement
                bool isVerifyR53 = (Marshal.SizeOf(ropLogonResponse.FolderIds[0]) == 8) && (ropLogonResponse.FolderIds.Length == 13);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR53,
                    53,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] FolderIds (104 bytes): 13 64-bit identifiers.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R55");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R55
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.ResponseFlags.GetType(),
                    55,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ResponseFlags (1 byte): A flags structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R71,ResponseFlags:{0}", ropLogonResponse.ResponseFlags);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R71
                bool isVerifyR71 = ((ropLogonResponse.ResponseFlags & 0x01) == (byte)ResponseFlags.Reserved)
                                   || ((ropLogonResponse.ResponseFlags & 0x02) == (byte)ResponseFlags.OwnerRight)
                                   || ((ropLogonResponse.ResponseFlags & 0x04) == (byte)ResponseFlags.SendAsRight)
                                   || ((ropLogonResponse.ResponseFlags & 0x10) == (byte)ResponseFlags.OOF);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR71,
                    71,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ResponseFlags (1 byte): The possible values[the value of Reserved is 0x01, the value of OwnerRight is 0x02, the value of SendAsRight is 0x04, the value of OOF is 0x10] are specified in [MS-OXCSTOR] Section 2.2.1.1.3.");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R57
                bool isVerifyR57 = this.IsGUID(ropLogonResponse.MailboxGuid);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R57, the actual type of MailboxGuid type is :{0}", ropLogonResponse.MailboxGuid.GetType().Name);
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR57,
                    57,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] MailboxGuid (16 bytes): A GUID.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R59, the actual type of RepId is {0} and length of it is {1}", ropLogonResponse.ReplId.GetType().Name, ropLogonResponse.ReplId.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R59
                // The 16-bit is 2-element byte array. If the RepId is byte array and the length of it is 2, the R59 is verified
                bool isVerifyR59 = (ropLogonResponse.ReplId.GetType() == typeof(byte[])) &&
                                   (ropLogonResponse.ReplId.Length == 2);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR59,
                    59,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ReplId (2 bytes): An identifier.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R61, the actual type of ReplGuid is:{0}", ropLogonResponse.ReplGuid.GetType().Name);
                
                // Verify MS-OXCROPS requirement: MS-OXCROPS_R61

                // Determine whether ReplGuid is GUID or not. 
                bool isVerifyR61 = this.IsGUID(ropLogonResponse.ReplGuid);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR61,
                    61,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ReplGuid (16 bytes): ): A GUID.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R62, the actual type of ReplGuid is {0}", ropLogonResponse.ReplGuid.GetType().Name);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R62

                // Determine whether ReplGuid is GUID or not. 
                // If ReplGuid is GUID, it indicates this field specifies the replica GUID.
                bool isVerifyR62 = this.IsGUID(ropLogonResponse.ReplGuid);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR62,
                    62,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] ReplGuid (16 bytes): This field specifies the replica GUID that is associated with the replica ID, which is specified in the ReplId field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R63");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R63
                // This LogonTime structure is generated based on the IDL from MS-OXCROPS,
                // If this field is type of LogonTime, the requirement is verified.
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(LogonTime),
                    ropLogonResponse.LogonTime.GetType(),
                    63,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] LogonTime (8 bytes): A LogonTime structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R65");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R65
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ulong),
                    ropLogonResponse.GwartTime.GetType(),
                    65,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] GwartTime (8 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R74");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R74
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    ropLogonResponse.StoreState.GetType(),
                    74,
                    @"[In RopLogon ROP Success Response Buffer for Private Mailboxes] StoreState (4 bytes): A flags structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R77");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R77
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.Seconds.GetType(),
                    77,
                    @"[In LogonTime Structure] Seconds (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R79");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R79
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.Minutes.GetType(),
                    79,
                    @"[In LogonTime Structure] Minutes (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R81");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R81
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.Hour.GetType(),
                    81,
                    @"[In LogonTime Structure] Hour (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R83");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R83
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.DayOfWeek.GetType(),
                    83,
                    @"[In LogonTime Structure] DayOfWeek (1 byte): An enumeration.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R85");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R85
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.Day.GetType(),
                    85,
                    @"[In LogonTime Structure] Day (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R87");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R87
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonTime.Month.GetType(),
                    87,
                    @"[In LogonTime Structure] Month (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R89");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R89
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropLogonResponse.LogonTime.Year.GetType(),
                    89,
                    @"[In LogonTime Structure] Year (2 bytes): An unsigned integer.");
            }
            else if (logonFlags == (byte)LogonFlags.PublicFolder)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R91");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R91
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.RopId.GetType(),
                    91,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] RopId (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R93");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R93
                Site.CaptureRequirementIfAreEqual<byte>(
                    (byte)RopId.RopLogon,
                    ropLogonResponse.RopId,
                    93,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] RopId (1 byte): For this operation[RopLogon], this field[RopId (1 byte)] is set to 0xFE.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R94");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R94
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.OutputHandleIndex.GetType(),
                    94,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] OutputHandleIndex (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R96");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R96
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    ropLogonResponse.ReturnValue.GetType(),
                    96,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] ReturnValue (4 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R98");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R98
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropLogonResponse.ReturnValue,
                    98,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] ReturnValue (4 bytes): For this response[Success Response Buffer for Public Folders], this field is set to 0x00000000.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R99");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R99
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropLogonResponse.LogonFlags.GetType(),
                    99,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] LogonFlags (1 byte): A flags structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R100");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R100
                Site.CaptureRequirementIfAreEqual<byte>(
                    (byte)LogonFlags.PublicFolder,
                    ropLogonResponse.LogonFlags,
                    100,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] LogonFlags (1 byte): The possible values[the value of Public is 0x00] are specified in [MS-OXCSTOR]2.2.1.1.4.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R102, the actual type of FolderIds is {0} and length of array is {1}", ropLogonResponse.FolderIds.GetType().Name, ropLogonResponse.FolderIds.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R102
                // If the type of FolderIds is long[] or ulong[] and the length of array is 13, then verify this requirement
                bool isVerifyR102 = (ropLogonResponse.FolderIds.GetType() == typeof(long[]) ||
                                    ropLogonResponse.FolderIds.GetType() == typeof(ulong[])) &&
                                    (ropLogonResponse.FolderIds.Length == 13);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR102,
                    102,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] FolderIds (104 bytes): 13 64-bit identifiers.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R104, the actual type of ReqlId is {0} and length of it is {1}", ropLogonResponse.ReplId.GetType().Name, ropLogonResponse.ReplId.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R104
                bool isVerifyR104 = (ropLogonResponse.ReplId.GetType() == typeof(byte[])) &&
                                    (ropLogonResponse.ReplId.Length == 2);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR104,
                    104,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] ReplId (2 bytes): An identifier.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R106, the actual type of ReplGuid is :{0}", ropLogonResponse.ReplGuid.GetType().Name);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R106
                bool isVerifyR106 = this.IsGUID(ropLogonResponse.ReplGuid);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR106,
                    106,
                    @"[In RopLogon ROP Success Response Buffer for Public Folders] ReplGuid (16 bytes): A GUID that specifies the replica GUID associated with the replica ID that is specified in the ReplId field.");

                if (Common.IsRequirementEnabled(4720, this.Site))
                {
                    bool isVerifyR4720 = false;
                    for (int i = 0; i < ropLogonResponse.PerUserGuid.Length; i++)
                    {
                        bool verify = ropLogonResponse.PerUserGuid[i] != 0x00;

                        // If one byte of PerUserGuid is zero then break the iteration 
                        if (verify)
                        {
                            isVerifyR4720 = true;
                            break;
                        }
                    }

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4720, the actual value of PerUserGuid is {0}", BitConverter.ToString(ropLogonResponse.PerUserGuid));

                    // Verify MS-OXCROPS requirement: MS-OXCROPS_R4720
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR4720,
                        4720,
                        @"[In Appendix B: Product Behavior] Implementation does not set the PerUserGuid field to an empty GUID. (<5> Section 2.2.3.1.3: Exchange 2007 does not set the PerUserGuid field to an empty GUID.)");
                }

                if (Common.IsRequirementEnabled(214, this.Site))
                {
                    bool isVerifyR214 = true;
                    for (int i = 0; i < ropLogonResponse.PerUserGuid.Length; i++)
                    {
                        isVerifyR214 = isVerifyR214 & (ropLogonResponse.PerUserGuid[i] == 0x00);

                        // If one byte of PerUserGuid not zero then break the iteration 
                        if (!isVerifyR214)
                        {
                            break;
                        }
                    }

                    // Add the comment information for debugging
                    Site.Log.Add(LogEntryKind.Comment, "Verify MS-OXCROPS_R214,PerUserGuid:{0}", ropLogonResponse.PerUserGuid);

                    // Verify MS-OXCROPS requirement: 214
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR214,
                        214,
                        @"[In Appendix B: Product Behavior] PerUserGuid (16 bytes): Implementation does set this field to an empty GUID (all zeroes). (Microsoft Exchange Server 2010 and above follow this behavior.)");
                }               
            }
        }

        /// <summary>
        /// Verify RopLogon Redirect Response
        /// </summary>
        /// <param name="ropLogonResponse">The response of the RopLogon Request</param>
        private void VerifyRopLogonRedirectResponse(RopLogonResponse ropLogonResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R112");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R112
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.RopId.GetType(),
                112,
                @"[In RopLogon ROP Redirect Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R114");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R114
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopLogon,
                ropLogonResponse.RopId,
                114,
                @"[In RopLogon ROP Redirect Response Buffer] RopId (1 byte): For this operation[RopLogon], this field[RopId (1 byte)] is set to 0xFE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R115");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R115
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.OutputHandleIndex.GetType(),
                115,
                @"[In RopLogon ROP Redirect Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R117");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R117
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropLogonResponse.ReturnValue.GetType(),
                117,
                @"[In RopLogon ROP Redirect Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R119");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R119
            Site.CaptureRequirementIfAreEqual<uint>(
                WrongServer,
                ropLogonResponse.ReturnValue,
                119,
                @"[In RopLogon ROP Redirect Response Buffer] ReturnValue (4 bytes): For this response[Redirect Response], this field is set to 0x00000478.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R120");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R120
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.LogonFlags.GetType(),
                120,
                @"[In RopLogon ROP Redirect Response Buffer] LogonFlags (1 byte): A flags structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R123");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R123
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.ServerNameSize.GetType(),
                123,
                @"[In RopLogon ROP Redirect Response Buffer] ServerNameSize (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R125,ServerName:{0}", ropLogonResponse.ServerName);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R125
            bool isVerifyR125 = this.IsNullTerminatedASCIIStr(ropLogonResponse.ServerName);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR125,
                125,
                @"[In RopLogon ROP Redirect Response Buffer] ServerName (variable): A null-terminated ASCII string.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R126");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R126
            Site.CaptureRequirementIfAreEqual<int>(
                (int)ropLogonResponse.ServerNameSize,
                ropLogonResponse.ServerName.Length,
                126,
                @"[In RopLogon ROP Redirect Response Buffer] ServerName (variable): The number of characters (including the terminating null character) contained in this field is specified by the ServerNameSize field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R124");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R124
            Site.CaptureRequirementIfAreEqual<int>(
                (int)ropLogonResponse.ServerNameSize,
                ropLogonResponse.ServerName.Length,
                124,
                @"[In RopLogon ROP Redirect Response Buffer] ServerNameSize (1 byte): This value specifies the length of the ServerName field.");
        }

        /// <summary>
        /// Verify RopLogon Failure Response
        /// </summary>
        /// <param name="ropLogonResponse">The response of the RopLogon Request</param>
        private void VerifyRopLogonFailureResponse(RopLogonResponse ropLogonResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R128");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R128
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.RopId.GetType(),
                128,
                @"[In RopLogon ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R130");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R130
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopLogon,
                ropLogonResponse.RopId,
                130,
                @"[In RopLogon ROP Failure Response Buffer] RopId (1 byte): For this operation[RopLogon], this field[RopId (1 byte)] is set to 0xFE.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R131");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R131
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLogonResponse.OutputHandleIndex.GetType(),
                131,
                @"[In RopLogon ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R133");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R133
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropLogonResponse.ReturnValue.GetType(),
                133,
                @"[In RopLogon ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R135,ReturnValue:{0}", ropLogonResponse.ReturnValue);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R135
            bool isVerifyR135 = (ropLogonResponse.ReturnValue != SuccessReturnValue) &&
                                (ropLogonResponse.ReturnValue != WrongServer);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR135,
                135,
                @"[In RopLogon ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000 or 0x00000478.");
        }
        #endregion

        #region RopGetReceiveFolder
        /// <summary>
        /// Verify RopGetReceiveFolder Success Response
        /// </summary>
        /// <param name="ropGetReceiveFolderResponse">The success response of RopGetReceiveFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetReceiveFolder request</param>
        private void VerifyRopGetReceiveFolderSuccessResponse(RopGetReceiveFolderResponse ropGetReceiveFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R148");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R148
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderResponse.RopId.GetType(),
                148,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R150");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R150
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetReceiveFolder,
                ropGetReceiveFolderResponse.RopId,
                150,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetReceiveFolder], this field[RopId (1 byte)] is set to 0x27.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R151");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R151
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderResponse.InputHandleIndex.GetType(),
                151,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer index.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R152");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R152
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetReceiveFolderResponse.InputHandleIndex,
                152,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R153");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R153
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetReceiveFolderResponse.ReturnValue.GetType(),
                153,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R155");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R155
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetReceiveFolderResponse.ReturnValue,
                155,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R156");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R156
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropGetReceiveFolderResponse.FolderId.GetType()),
                156,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] FolderId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R158, the actual value of ExplicitMessageClass is :{0}", BitConverter.ToString(ropGetReceiveFolderResponse.MessageClass));

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R158
            bool isVerifyR158 = this.IsNullTerminatedASCIIStr(ropGetReceiveFolderResponse.MessageClass);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR158,
                158,
                @"[In RopGetReceiveFolder ROP Success Response Buffer] ExplicitMessageClass (variable): A null-terminated ASCII string.");
        }

        /// <summary>
        /// Verify RopGetReceiveFolder Failure Response
        /// </summary>
        /// <param name="ropGetReceiveFolderResponse">The failure response of RopGetReceiveFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetReceiveFolder request</param>
        private void VerifyRopGetReceiveFolderFailureResponse(RopGetReceiveFolderResponse ropGetReceiveFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R160");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R160
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderResponse.RopId.GetType(),
                160,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] RopId (1 byte): An unsigned integer .");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R162");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R162
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetReceiveFolder,
                ropGetReceiveFolderResponse.RopId,
                162,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetReceiveFolder], this field[RopId (1 byte)] is set to 0x27.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R163");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R163
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderResponse.InputHandleIndex.GetType(),
                163,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] InputHandleIndex (1 byte): ): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R164");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R164
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetReceiveFolderResponse.InputHandleIndex,
                164,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R165");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R165
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetReceiveFolderResponse.ReturnValue.GetType(),
                165,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R167");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R167
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetReceiveFolderResponse.ReturnValue,
                167,
                @"[In RopGetReceiveFolder ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSetReceiveFolder
        /// <summary>
        /// Verify RopSetReceiveFolder Response
        /// </summary>
        /// <param name="ropSetReceiveFolderResponse">The response of RopSetReceiveFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetReceiveFolder request</param>
        private void VerifyRopSetReceiveFolderResponse(RopSetReceiveFolderResponse ropSetReceiveFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R181");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetReceiveFolderResponse.RopId.GetType(),
                181,
                @"[In RopSetReceiveFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R183");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R183
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetReceiveFolder,
                ropSetReceiveFolderResponse.RopId,
                183,
                @"[In RopSetReceiveFolder ROP Response Buffer] RopId (1 byte): For this operation[RopSetReceiveFolder], this field[RopId (1 byte)] is set to 0x26.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R184");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R184
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetReceiveFolderResponse.InputHandleIndex.GetType(),
                184,
                @"[In RopSetReceiveFolder ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R185");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R185
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetReceiveFolderResponse.InputHandleIndex,
                185,
                @"[In RopSetReceiveFolder ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R186");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R186
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetReceiveFolderResponse.ReturnValue.GetType(),
                186,
                @"[In RopSetReceiveFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetReceiveFolderTable
        /// <summary>
        /// Verify RopGetReceiveFolderTable Success Response 
        /// </summary>
        /// <param name="ropGetReceiveFolderTableResponse">The success response of RopGetReceiveFolderTable request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetReceiveFolderTable request</param>
        private void VerifyRopGetReceiveFolderTableSuccessResponse(RopGetReceiveFolderTableResponse ropGetReceiveFolderTableResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R197");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R197
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderTableResponse.RopId.GetType(),
                197,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R199");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R199
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetReceiveFolderTable,
                ropGetReceiveFolderTableResponse.RopId,
                199,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetReceiveFolderTable], this field[RopId (1 byte)] is set to 0x68.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R200");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R200
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderTableResponse.InputHandleIndex.GetType(),
                200,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R201");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R201
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetReceiveFolderTableResponse.InputHandleIndex,
                201,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R202");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R202
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetReceiveFolderTableResponse.ReturnValue.GetType(),
                202,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] ReturnValue (4 bytes): ): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R204");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R204
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetReceiveFolderTableResponse.ReturnValue,
                204,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R205");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R205
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetReceiveFolderTableResponse.RowCount.GetType(),
                205,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] RowCount (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R207");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R207
            // The class PropertyRowSet is List of PropertRow, if the type of Rows is PropertyRowSet can verify R207 
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyRowSet),
                ropGetReceiveFolderTableResponse.Rows.GetType(),
                207,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows (variable):An array of row structures.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCROPS_R209,the rows count:{0}",
                ropGetReceiveFolderTableResponse.Rows == null ? 0 : ropGetReceiveFolderTableResponse.Rows.Count);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R209
            bool isStandardPropertyRowOrFlaggedPropertyRow = true;
            if (ropGetReceiveFolderTableResponse.Rows != null)
            {
                // The rows count.If the count greater than 0, then the PropertyRows is not null and the array of PropertyRow length greater than 0.
                if (ropGetReceiveFolderTableResponse.Rows.Count > 0)
                {
                    foreach (PropertyRow propertyRow in ropGetReceiveFolderTableResponse.Rows.PropertyRows)
                    {
                        // According to the de-serialize approach of ropGetReceiveFolderTableResponse,flag indicates the type of PropertyRow
                        // From MS-OXCDATA,Flag is 0x00, representing StandardPropertyRow. Flag is 0x01, representing FlaggedPropertyRow. 
                        if ((propertyRow.Flag != ((byte)PropertyRowFlag.StandardPropertyRow)) &&
                           (propertyRow.Flag != (byte)PropertyRowFlag.FlaggedPropertyRow))
                        {
                            isStandardPropertyRowOrFlaggedPropertyRow = false;
                            break;
                        }
                    }
                }

                bool isVerifyR209 = isStandardPropertyRowOrFlaggedPropertyRow;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR209,
                    209,
                    @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows (variable): Each row is returned in either a StandardPropertyRow structure, as specified in[MS-OXCDATA] section 2.8.1.1, or a FlaggedPropertyRow structure, as specified in [MS-OXCDATA] section 2.8.1.2. ");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R210");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R210
            Site.CaptureRequirementIfAreEqual<uint>(
                ropGetReceiveFolderTableResponse.RowCount,
                (uint)ropGetReceiveFolderTableResponse.Rows.Count,
                210,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows (variable): The number of row structures contained in this field is specified by the RowCount field.");

            #region Verify R211
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R211");

            bool isVerifyR211 = true;

            PropertyRowSet propertyRowSet = ropGetReceiveFolderTableResponse.Rows;

            if (propertyRowSet != null)
            {
                foreach (PropertyRow propertyRow in propertyRowSet.PropertyRows)
                {
                    // ValueArray include only the PidTagFolderId, PidTagMessageClass, and PidTagLastModificationTime properties
                    isVerifyR211 &= propertyRow.PropertyValues.Count == 3;

                    if (!isVerifyR211)
                    {
                        break;
                    }

                    // De-serialize the property PidTagFolderId as PtypInteger64.
                    // Specified in MS-OXPROPS.
                    isVerifyR211 &= propertyRow.PropertyValues[0].Size() == sizeof(long);
                    if (!isVerifyR211)
                    {
                        break;
                    }

                    BitConverter.ToInt64(propertyRow.PropertyValues[0].Value, 0);

                    // De-serialize the property PidTagMessageClass as PtypString.
                    // Specified in MS-OXPROPS.
                    isVerifyR211 &= this.IsNullTerminatedASCIIStr(propertyRow.PropertyValues[1].Value);
                    if (!isVerifyR211)
                    {
                        break;
                    }

                    Encoding.Unicode.GetString(propertyRow.PropertyValues[1].Value);

                    // De-serialize the property PidTagLastModificationTime as PtypTime.
                    // Specified in MS-OXPROPS.
                    isVerifyR211 &= propertyRow.PropertyValues[2].Size() == sizeof(long);
                    if (!isVerifyR211)
                    {
                        break;
                    }

                    DateTime.FromFileTimeUtc(BitConverter.ToInt64(propertyRow.PropertyValues[2].Value, 0));
                }
            }

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R211
            // The de-serialize process has verified that the ValueArray include only the PidTagFolderId, PidTagMessageClass, 
            // and PidTagLastModificationTime properties, in that order, and no other properties
            Site.CaptureRequirementIfIsTrue(
                isVerifyR211,
                211,
                @"[In RopGetReceiveFolderTable ROP Success Response Buffer] Rows (variable): The ValueArray field of either the StandardPropertyRow structure or the FlaggedPropertyRow structure MUST include only the PidTagFolderId ([MS-OXCFOLD] section 2.2.2.2.1.6), PidTagMessageClass ([MS-OXCMSG] section 2.2.1.3), and PidTagLastModificationTime ([MS-OXCMSG] section 2.2.2.2) properties, in that order, and no other properties.");
            #endregion
        }

        /// <summary>
        /// Verify RopGetReceiveFolderTable Failure Response 
        /// </summary>
        /// <param name="ropGetReceiveFolderTableResponse">The failure response of RopGetReceiveFolderTable request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetReceiveFolderTable request</param>
        private void VerifyRopGetReceiveFolderTableFailureResponse(RopGetReceiveFolderTableResponse ropGetReceiveFolderTableResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R215");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R215
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderTableResponse.RopId.GetType(),
                215,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R217");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R217
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetReceiveFolderTable,
                ropGetReceiveFolderTableResponse.RopId,
                217,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetReceiveFolderTable], this field[RopId (1 byte)] is set to 0x68.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R218");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R218
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetReceiveFolderTableResponse.InputHandleIndex.GetType(),
                218,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R219");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R219
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetReceiveFolderTableResponse.InputHandleIndex,
                219,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R220");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R220
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetReceiveFolderTableResponse.ReturnValue.GetType(),
                220,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R222");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R222
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetReceiveFolderTableResponse.ReturnValue,
                222,
                @"[In RopGetReceiveFolderTable ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetStoreState
        /// <summary>
        /// Verify RopGetStoreState Success Response
        /// </summary>
        /// <param name="ropGetStoreStateResponse">The success response of RopGetStoreState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStoreState request</param>
        private void VerifyRopGetStoreStateSuccessResponse(RopGetStoreStateResponse ropGetStoreStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R232");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R232
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStoreStateResponse.RopId.GetType(),
                232,
                @"[In RopGetStoreState ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R234");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R234
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStoreState,
                ropGetStoreStateResponse.RopId,
                234,
                @"[In RopGetStoreState ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetStoreState], this field[RopId (1 byte)] is set to 0x7B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R235");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R235
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStoreStateResponse.InputHandleIndex.GetType(),
                235,
                @"[In RopGetStoreState ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R236");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R236
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStoreStateResponse.InputHandleIndex,
                236,
                @"[In RopGetStoreState ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R237");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R237
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStoreStateResponse.ReturnValue.GetType(),
                237,
                @"[In RopGetStoreState ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R239");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R239
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetStoreStateResponse.ReturnValue,
                239,
                @"[In RopGetStoreState ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R240");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R240
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ropGetStoreStateResponse.StoreState.GetType()),
                240,
                @"[In RopGetStoreState ROP Success Response Buffer] StoreState (4 bytes): A flags structure that contains flags that indicate the state of the mailbox for the logged on user.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R241,StoreState:{0}", ropGetStoreStateResponse.StoreState);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R241
            // The possible values' condition is in MS-OXCSTOR, verify StoreState is either 0x01000000 or 0x00000000. 
            bool isVerifyR241 = (ropGetStoreStateResponse.StoreState == (uint)StoreState.StoreHasSearches)
                || (ropGetStoreStateResponse.StoreState == (uint)StoreState.None);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR241,
                241,
                @"[In RopGetStoreState ROP Success Response Buffer] StoreState (4 bytes): The possible values[0x01000000(STORE_HAS_SEARCHES), if the mailbox currently has any active search folders. Otherwise, 0x00000000.]are specified in [MS-OXCSTOR] section 2.2.1.5.2.");
        }

        /// <summary>
        /// Verify RopGetStoreState Failure Response
        /// </summary>
        /// <param name="ropGetStoreStateResponse">The failure response of RopGetStoreState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStoreState request</param>
        private void VerifyRopGetStoreStateFailureResponse(RopGetStoreStateResponse ropGetStoreStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R243");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R243
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStoreStateResponse.RopId.GetType(),
                243,
                @"[In RopGetStoreState ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R245");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R245
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStoreState,
                ropGetStoreStateResponse.RopId,
                245,
                @"[In RopGetStoreState ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetStoreState], this field[RopId (1 byte)] is set to 0x7B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R246");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R246
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStoreStateResponse.InputHandleIndex.GetType(),
                246,
                @"[In RopGetStoreState ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R247");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R247
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStoreStateResponse.InputHandleIndex,
                247,
                @"[In RopGetStoreState ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R248");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R248
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStoreStateResponse.ReturnValue.GetType(),
                248,
                @"[In RopGetStoreState ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R250");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R250
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetStoreStateResponse.ReturnValue,
                250,
                @"[In RopGetStoreState ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetOwningServers
        /// <summary>
        /// Verify RopGetOwningServers Success Response
        /// </summary>
        /// <param name="ropGetOwningServerResponse">The success response of RopGetOwningServers request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetOwningServers request</param>
        private void VerifyRopGetOwningServersSuccessResponse(RopGetOwningServersResponse ropGetOwningServerResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R262");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R262
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetOwningServerResponse.RopId.GetType(),
                262,
                @"[In RopGetOwningServers ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R264");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R264
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetOwningServers,
                ropGetOwningServerResponse.RopId,
                264,
                @"[In RopGetOwningServers ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetOwningServers], this field[RopId (1 byte)] is set to 0x42.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R265");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R265
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetOwningServerResponse.InputHandleIndex.GetType(),
                265,
                @"[In RopGetOwningServers ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R266");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R266
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetOwningServerResponse.InputHandleIndex,
                266,
                @"[In RopGetOwningServers ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R267");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R267
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetOwningServerResponse.ReturnValue.GetType(),
                267,
                @"[In RopGetOwningServers ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R269");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R269
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetOwningServerResponse.ReturnValue,
                269,
                @"[In RopGetOwningServers ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response Buffer], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R270");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R270
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetOwningServerResponse.OwningServersCount.GetType(),
                270,
                @"[In RopGetOwningServers ROP Success Response Buffer] OwningServersCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R272");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R272
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetOwningServerResponse.CheapServersCount.GetType(),
                272,
                @"[In RopGetOwningServers ROP Success Response Buffer] CheapServersCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R274,the length of the OwningServers is {0}", ropGetOwningServerResponse.OwningServers.Length);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R274
            // Assume the result is true. If one of List is not null-terminated ASCII string, then false.
            bool isVerifyR274 = true;
            foreach (string service in ropGetOwningServerResponse.OwningServers)
            {
                byte[] buffer = Encoding.ASCII.GetBytes(service);
                isVerifyR274 = isVerifyR274 && this.IsNullTerminatedASCIIStr(buffer);

                // If one of List is not null-terminated ASCII string then break iteration
                if (!isVerifyR274)
                {
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR274,
                274,
                @"[In RopGetOwningServers ROP Success Response Buffer] OwningServers (variable):A list of null-terminated ASCII strings.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R275");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R275
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropGetOwningServerResponse.OwningServersCount,
                (ushort)ropGetOwningServerResponse.OwningServers.Length,
                275,
                @"[In RopGetOwningServers ROP Success Response Buffer] OwningServers (variable): The number of strings contained in this field is specified by the OwningServersCount field.");
        }

        /// <summary>
        /// Verify RopGetOwningServers Failure Response
        /// </summary>
        /// <param name="ropGetOwningServerResponse">The failure response of RopGetOwningServers request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetOwningServers request</param>
        private void VerifyRopGetOwningServersFailureResponse(RopGetOwningServersResponse ropGetOwningServerResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R277");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R277
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetOwningServerResponse.RopId.GetType(),
                277,
                @"[In RopGetOwningServers ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R279");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R279
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetOwningServers,
                ropGetOwningServerResponse.RopId,
                279,
                @"[In RopGetOwningServers ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetOwningServers], this field[RopId (1 byte)] is set to 0x42.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R280");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R280
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetOwningServerResponse.InputHandleIndex.GetType(),
                280,
                @"[In RopGetOwningServers ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R281");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R281
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetOwningServerResponse.InputHandleIndex,
                281,
                @"[In RopGetOwningServers ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R282");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R282
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetOwningServerResponse.ReturnValue.GetType(),
                282,
                @"[In RopGetOwningServers ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R284");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R284
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetOwningServerResponse.ReturnValue,
                284,
                @"[In RopGetOwningServers ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response Buffer], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopPublicFolderIsGhosted
        /// <summary>
        /// Verify RopPublicFolderIsGhosted Success Response
        /// </summary>
        /// <param name="ropPublicFolderIsGhostedResponse">The success response of RopPublicFolderIsGhosted request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopPublicFolderIsGhosted request</param>
        private void VerifyRopPublicFolderIsGhostedSuccessResponse(RopPublicFolderIsGhostedResponse ropPublicFolderIsGhostedResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R296");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R296
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropPublicFolderIsGhostedResponse.RopId.GetType(),
                296,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R298");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R298
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopPublicFolderIsGhosted,
                ropPublicFolderIsGhostedResponse.RopId,
                298,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] RopId (1 byte): For this operation[RopPublicFolderIsGhosted], this field[RopId (1 byte)] is set to 0x45.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R299");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R299
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropPublicFolderIsGhostedResponse.InputHandleIndex.GetType(),
                299,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R300");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R300
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropPublicFolderIsGhostedResponse.InputHandleIndex,
                300,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R301");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R301
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropPublicFolderIsGhostedResponse.ReturnValue.GetType(),
                301,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R303");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R303
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropPublicFolderIsGhostedResponse.ReturnValue,
                303,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R304,IsGhosted:{0}", ropPublicFolderIsGhostedResponse.IsGhosted);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R304
            // If IsGhosted is of the byte type and can be represented by the boolean type, then Verify R304
            bool isVerifyR304 = (typeof(byte) == ropPublicFolderIsGhostedResponse.IsGhosted.GetType()) &&
                                ((Convert.ToBoolean(ropPublicFolderIsGhostedResponse.IsGhosted) == true) ||
                                (Convert.ToBoolean(ropPublicFolderIsGhostedResponse.IsGhosted) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR304,
                304,
                @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] IsGhosted (1 byte): A Boolean.");

            // The field ServersCount,CheapServersCount and Servers are present if IsGhosted is non-zero and are not present if IsGhosted is zero
            if (ropPublicFolderIsGhostedResponse.IsGhosted != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R306");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R306
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropPublicFolderIsGhostedResponse.ServersCount.GetType(),
                    306,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] ServersCount (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R311");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R311
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropPublicFolderIsGhostedResponse.CheapServersCount.GetType(),
                    311,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] CheapServersCount (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R315,the length of the Servers is {0}", ropPublicFolderIsGhostedResponse.Servers.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R315
                // Assume the result is true. If one of List is not null-terminated ASCII string, then false.
                bool isVerifyR315 = true;
                foreach (string s in ropPublicFolderIsGhostedResponse.Servers)
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(s);
                    isVerifyR315 = isVerifyR315 && this.IsNullTerminatedASCIIStr(buffer);

                    // If one of List is not null-terminated ASCII string then break iteration
                    if (!isVerifyR315)
                    {
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR315,
                    315,
                    @"[In RopPublicFolderIsGhosted ROP Success Response Buffer] Servers (optional) (variable): A list of null-terminated ASCII strings.");
            }
        }

        /// <summary>
        /// Verify RopPublicFolderIsGhosted Failure Response
        /// </summary>
        /// <param name="ropPublicFolderIsGhostedResponse">The failure response of RopPublicFolderIsGhosted request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopPublicFolderIsGhosted request</param>
        private void VerifyRopPublicFolderIsGhostedFailureResponse(RopPublicFolderIsGhostedResponse ropPublicFolderIsGhostedResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R321");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R321
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropPublicFolderIsGhostedResponse.RopId.GetType(),
                321,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R323");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R323
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopPublicFolderIsGhosted,
                ropPublicFolderIsGhostedResponse.RopId,
                323,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] RopId (1 byte): For this operation[RopPublicFolderIsGhosted], this field[RopId (1 byte)] is set to 0x45.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R324");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R324
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropPublicFolderIsGhostedResponse.InputHandleIndex.GetType(),
                324,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R325");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R325
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropPublicFolderIsGhostedResponse.InputHandleIndex,
                325,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value of the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R326");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R326
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropPublicFolderIsGhostedResponse.ReturnValue.GetType(),
                326,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R328");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R328
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropPublicFolderIsGhostedResponse.ReturnValue,
                328,
                @"[In RopPublicFolderIsGhosted ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopLongTermIdFromId
        /// <summary>
        /// Verify RopLongTermIdFromId Success Response
        /// </summary>
        /// <param name="ropLongTermIdFromIdResponse">The success response of RopLongTermIdFromId request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopLongTermIdFromId request</param>
        private void VerifyRopLongTermIdFromIdSuccessResponse(RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R340");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R340
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLongTermIdFromIdResponse.RopId.GetType(),
                340,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] RopId (1 byte): An unsigned integer .");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R342");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R342
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopLongTermIdFromId,
                ropLongTermIdFromIdResponse.RopId,
                342,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] RopId (1 byte): For this operation[RopLongTermIdFromId], this field[RopId (1 byte)] is set to 0x43.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R343");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R343
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLongTermIdFromIdResponse.InputHandleIndex.GetType(),
                343,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R344");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R344
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropLongTermIdFromIdResponse.InputHandleIndex,
                344,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value of the InputHandleIndex field in the request. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R345");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R345
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropLongTermIdFromIdResponse.ReturnValue.GetType(),
                345,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R347");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R347
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropLongTermIdFromIdResponse.ReturnValue,
                347,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R348");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R348
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(LongTermId),
                ropLongTermIdFromIdResponse.LongTermId.GetType(),
                348,
                @"[In RopLongTermIdFromId ROP Success Response Buffer] LongTermId (24 bytes): A LongTermID structure.");
        }

        /// <summary>
        /// Verify RopLongTermIdFromId Failure Response
        /// </summary>
        /// <param name="ropLongTermIdFromIdResponse">The failure response of RopLongTermIdFromId request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopLongTermIdFromId request</param>
        private void VerifyRopLongTermIdFromIdFailureResponse(RopLongTermIdFromIdResponse ropLongTermIdFromIdResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R351");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R351
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLongTermIdFromIdResponse.RopId.GetType(),
                351,
                @"[In RopLongTermIdFromId ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R353");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R353
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopLongTermIdFromId,
                ropLongTermIdFromIdResponse.RopId,
                353,
                @"[In RopLongTermIdFromId ROP Failure Response Buffer] RopId (1 byte): For this operation[RopLongTermIdFromId], this field[RopId (1 byte)] is set to 0x43.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R354");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R354
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLongTermIdFromIdResponse.InputHandleIndex.GetType(),
                354,
                @"[In RopLongTermIdFromId ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R355");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R355
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropLongTermIdFromIdResponse.InputHandleIndex,
                355,
                @"[In RopLongTermIdFromId ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R356");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R356
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropLongTermIdFromIdResponse.ReturnValue.GetType(),
                356,
                @"[In RopLongTermIdFromId ROP Failure Response Buffer] ReturnValue (4 bytes):An unsigned integer.");

            // Refer to MS-OXCROPS: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve the Server object and, 
            // therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4713");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4713
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropLongTermIdFromIdResponse.ReturnValue,
                    4713,
                    @"[In Appendix B: Product Behavior] For some ROPs, Implementation does use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid. (<23> Section 3.2.5.1: For some ROPs, Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.)");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R358");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R358
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropLongTermIdFromIdResponse.ReturnValue,
                    358,
                    @"[In RopLongTermIdFromId ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
            }
        }
        #endregion

        #region RopIdFromLongTermId
        /// <summary>
        /// Verify RopIdFromLongTermId Success Response
        /// </summary>
        /// <param name="ropIdFromLongTermIdResponse">The success response of RopIdFromLongTermId request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopIdFromLongTermId request</param>
        private void VerifyRopIdFromLongTermIdSuccessResponse(RopIdFromLongTermIdResponse ropIdFromLongTermIdResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R371");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R371
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropIdFromLongTermIdResponse.RopId.GetType(),
                371,
                @"[In RopIdFromLongTermId ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R373");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R373
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopIdFromLongTermId,
                ropIdFromLongTermIdResponse.RopId,
                373,
                @"[In RopIdFromLongTermId ROP Success Response Buffer]RopId (1 byte): For this operation[RopIdFromLongTermId], this field[RopId (1 byte)] is set to 0x44.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R374");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R374
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropIdFromLongTermIdResponse.InputHandleIndex.GetType(),
                374,
                @"[In RopIdFromLongTermId ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R375");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R375
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropIdFromLongTermIdResponse.InputHandleIndex,
                375,
                @"[In RopIdFromLongTermId ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R376");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R376
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropIdFromLongTermIdResponse.ReturnValue.GetType(),
                376,
                @"[In RopIdFromLongTermId ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R378");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R378
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropIdFromLongTermIdResponse.ReturnValue,
                378,
                @"[In RopIdFromLongTermId ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R379");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R379
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropIdFromLongTermIdResponse.ObjectId.GetType()),
                379,
                @"[In RopIdFromLongTermId ROP Success Response Buffer] ObjectId (8 bytes): ): An identifier.");
        }

        /// <summary>
        /// Verify RopIdFromLongTermId Failure Response
        /// </summary>
        /// <param name="ropIdFromLongTermIdResponse">The failure response of RopIdFromLongTermId request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopIdFromLongTermId request</param>
        private void VerifyRopIdFromLongTermIdFailureResponse(RopIdFromLongTermIdResponse ropIdFromLongTermIdResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R381");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R381
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropIdFromLongTermIdResponse.RopId.GetType(),
                381,
                @"[In RopIdFromLongTermId ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R383");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R383
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopIdFromLongTermId,
                ropIdFromLongTermIdResponse.RopId,
                383,
                @"[In RopIdFromLongTermId ROP Failure Response Buffer] RopId (1 byte): For this operation[RopIdFromLongTermId], this field[RopId (1 byte)] is set to 0x44.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R384");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R384
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropIdFromLongTermIdResponse.InputHandleIndex.GetType(),
                384,
                @"[In RopIdFromLongTermId ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R385");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R385
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropIdFromLongTermIdResponse.InputHandleIndex,
                385,
                @"[In RopIdFromLongTermId ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R386");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R386
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropIdFromLongTermIdResponse.ReturnValue.GetType(),
                386,
                @"[In RopIdFromLongTermId ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Refer to MS-OXCROPS: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve 
            // the Server object and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4713");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4713
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropIdFromLongTermIdResponse.ReturnValue,
                    4713,
                    @"[In Appendix B: Product Behavior] For some ROPs, Implementation does use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid. (<23> Section 3.2.5.1: For some ROPs, Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.)");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R388");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R388
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropIdFromLongTermIdResponse.ReturnValue,
                    388,
                    @"[In RopIdFromLongTermId ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
            }
        }
        #endregion

        #region RopGetPerUserLongTermIds
        /// <summary>
        /// Verify RopGetPerUserLongTermIds Success Response
        /// </summary>
        /// <param name="ropGetPerUserLongTermIdsResponse">The success response of RopGetPerUserLongTermIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPerUserLongTermIds request</param>
        private void VerifyRopGetPerUserLongTermIdsSuccessResponse(RopGetPerUserLongTermIdsResponse ropGetPerUserLongTermIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R400");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R400
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserLongTermIdsResponse.RopId.GetType(),
                400,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R402");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R402
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPerUserLongTermIds,
                ropGetPerUserLongTermIdsResponse.RopId,
                402,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetPerUserLongTermIds], this field[RopId (1 byte)] is set to 0x60.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R403");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R403
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserLongTermIdsResponse.InputHandleIndex.GetType(),
                403,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R404");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R404
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPerUserLongTermIdsResponse.InputHandleIndex,
                404,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R405");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R405
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPerUserLongTermIdsResponse.ReturnValue.GetType(),
                405,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R407");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R407
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPerUserLongTermIdsResponse.ReturnValue,
                407,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R408");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R408
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetPerUserLongTermIdsResponse.LongTermIdCount.GetType(),
                408,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIdCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R410");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R410
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(LongTermId[]),
                ropGetPerUserLongTermIdsResponse.LongTermIds.GetType(),
                410,
                @"[In RopGetPerUserLongTermIds ROP Success Response Buffer] LongTermIds (variable): An array of LongTermID structures, of which the format is specified in [MS-OXCDATA] section 2.2.1.3.1. ");
        }

        /// <summary>
        /// Verify RopGetPerUserLongTermIds Failure Response
        /// </summary>
        /// <param name="ropGetPerUserLongTermIdsResponse">The failure response of RopGetPerUserLongTermIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPerUserLongTermIds request</param>
        private void VerifyRopGetPerUserLongTermIdsFailureResponse(RopGetPerUserLongTermIdsResponse ropGetPerUserLongTermIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R414");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R414
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserLongTermIdsResponse.RopId.GetType(),
                414,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R416");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R416
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPerUserLongTermIds,
                ropGetPerUserLongTermIdsResponse.RopId,
                416,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetPerUserLongTermIds], this field[RopId (1 byte)] is set to 0x60.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R417");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R417
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserLongTermIdsResponse.InputHandleIndex.GetType(),
                417,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R418");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R418
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPerUserLongTermIdsResponse.InputHandleIndex,
                418,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R419");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R419
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPerUserLongTermIdsResponse.ReturnValue.GetType(),
                419,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R421");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R421
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPerUserLongTermIdsResponse.ReturnValue,
                421,
                @"[In RopGetPerUserLongTermIds ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetPerUserGuid
        /// <summary>
        /// Verify RopGetPerUserGuid Success Response
        /// </summary>
        /// <param name="ropGetPerUserGuidResponse">The success response of RopGetPerUserGuid request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPerUserGuid request</param>
        private void VerifyRopGetPerUserGuidSuccessResponse(RopGetPerUserGuidResponse ropGetPerUserGuidResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R435");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R435
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserGuidResponse.RopId.GetType(),
                435,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R437");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R437
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPerUserGuid,
                ropGetPerUserGuidResponse.RopId,
                437,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetPerUserGuid], this field[RopId (1 byte)] is set to 0x61.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R438");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R438
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserGuidResponse.InputHandleIndex.GetType(),
                438,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R439");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R439
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPerUserGuidResponse.InputHandleIndex,
                439,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R440");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R440
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPerUserGuidResponse.ReturnValue.GetType(),
                440,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R442");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R442
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPerUserGuidResponse.ReturnValue,
                442,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R443, the actual type of DatabaseGuid is :{0}", ropGetPerUserGuidResponse.DatabaseGuid.GetType().Name);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R443
            bool isVerifyR443 = this.IsGUID(ropGetPerUserGuidResponse.DatabaseGuid);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR443,
                443,
                @"[In RopGetPerUserGuid ROP Success Response Buffer] DatabaseGuid (16 bytes): A GUID.");
        }

        /// <summary>
        /// Verify RopGetPerUserGuid Failure Response
        /// </summary>
        /// <param name="ropGetPerUserGuidResponse">The failure response of RopGetPerUserGuid request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPerUserGuid request</param>
        private void VerifyRopGetPerUserGuidFailureResponse(RopGetPerUserGuidResponse ropGetPerUserGuidResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R445");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R445
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserGuidResponse.RopId.GetType(),
                445,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R447");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R447
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPerUserGuid,
                ropGetPerUserGuidResponse.RopId,
                447,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetPerUserGuid], this field[RopId (1 byte)] is set to 0x61.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R448");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R448
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPerUserGuidResponse.InputHandleIndex.GetType(),
                448,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R449");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R449
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPerUserGuidResponse.InputHandleIndex,
                449,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R450");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R450
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPerUserGuidResponse.ReturnValue.GetType(),
                450,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R452");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R452
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPerUserGuidResponse.ReturnValue,
                452,
                @"[In RopGetPerUserGuid ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopReadPerUserInformation
        /// <summary>
        /// Verify RopReadPerUserInformation Success Response
        /// </summary>
        /// <param name="ropReadPerUserInformationResponse">The success response of RopReadPerUserInformation request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReadPerUserInformation request</param>
        private void VerifyRopReadPerUserInformationSuccessResponse(RopReadPerUserInformationResponse ropReadPerUserInformationResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R476");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R476
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadPerUserInformationResponse.RopId.GetType(),
                476,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R478");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R478
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReadPerUserInformation,
                ropReadPerUserInformationResponse.RopId,
                478,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] RopId (1 byte): For this operation[RopReadPerUserInformation], this field[RopId (1 byte)] is set to 0x63.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R479");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R479
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadPerUserInformationResponse.InputHandleIndex.GetType(),
                479,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R480");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R480
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReadPerUserInformationResponse.InputHandleIndex,
                480,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R481");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R481
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReadPerUserInformationResponse.ReturnValue.GetType(),
                481,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R483");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R483
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropReadPerUserInformationResponse.ReturnValue,
                483,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R484,HasFinished:{0}", ropReadPerUserInformationResponse.HasFinished);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R484
            // If HasFinished is of the byte type and can be represented by the boolean type, then Verify R484
            bool isVerifyR484 = (typeof(byte) == ropReadPerUserInformationResponse.HasFinished.GetType()) &&
                                ((Convert.ToBoolean(ropReadPerUserInformationResponse.HasFinished) == true) ||
                                (Convert.ToBoolean(ropReadPerUserInformationResponse.HasFinished) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR484,
                484,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] HasFinished (1 byte): A Boolean.");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R486
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropReadPerUserInformationResponse.DataSize.GetType(),
                486,
                @"[In RopReadPerUserInformation ROP Success Response Buffer] DataSize (2 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopReadPerUserInformation Failure Response
        /// </summary>
        /// <param name="ropReadPerUserInformationResponse">The failure response of RopReadPerUserInformation request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReadPerUserInformation request</param>
        private void VerifyRopReadPerUserInformationFailureResponse(RopReadPerUserInformationResponse ropReadPerUserInformationResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R491");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R491
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadPerUserInformationResponse.RopId.GetType(),
                491,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R493");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R493
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReadPerUserInformation,
                ropReadPerUserInformationResponse.RopId,
                493,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] RopId (1 byte): For this operation[RopReadPerUserInformation], this field[RopId (1 byte)] is set to 0x63.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R494");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R494
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadPerUserInformationResponse.InputHandleIndex.GetType(),
                494,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R495");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R495
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReadPerUserInformationResponse.InputHandleIndex,
                495,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R496");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R496
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReadPerUserInformationResponse.ReturnValue.GetType(),
                496,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R498");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R498
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropReadPerUserInformationResponse.ReturnValue,
                498,
                @"[In RopReadPerUserInformation ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopWritePerUserInformation
        /// <summary>
        /// Verify RopWritePerUserInformation Response
        /// </summary>
        /// <param name="ropWritePerUserInformationResponse">The response of RopWritePerUserInformation request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopWritePerUserInformation request</param>
        private void VerifyRopWritePerUserInformationResponse(RopWritePerUserInformationResponse ropWritePerUserInformationResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R524");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R524
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropWritePerUserInformationResponse.RopId.GetType(),
                524,
                @"[In RopWritePerUserInformation ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R526");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R526
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopWritePerUserInformation,
                ropWritePerUserInformationResponse.RopId,
                526,
                @"[In RopWritePerUserInformation ROP Response Buffer] RopId (1 byte): For this operation[RopWritePerUserInformation], this field[RopId (1 byte)] is set to 0x64.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R527");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R527
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropWritePerUserInformationResponse.InputHandleIndex.GetType(),
                527,
                @"[In RopWritePerUserInformation ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R528");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R528
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropWritePerUserInformationResponse.InputHandleIndex,
                528,
                @"[In RopWritePerUserInformation ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R529");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R529
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropWritePerUserInformationResponse.ReturnValue.GetType(),
                529,
                @"[In RopWritePerUserInformation ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Folder ROPs

        #region RopOpenFolder
        /// <summary>
        /// Verify RopOpenFolder Success Response
        /// </summary>
        /// <param name="ropOpenFolderResponse">The success response of RopOpenFolder request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenFolder request</param>
        private void VerifyRopOpenFolderSuccessResponse(RopOpenFolderResponse ropOpenFolderResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R547");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R547
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenFolderResponse.RopId.GetType(),
                547,
                @"[In RopOpenFolder ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R549");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R549
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenFolder,
                ropOpenFolderResponse.RopId,
                549,
                @"[In RopOpenFolder ROP Success Response Buffer] RopId (1 byte): For this operation[RopOpenFolder], this field[RopId (1 byte)] is set to 0x02.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R550");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R550
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenFolderResponse.OutputHandleIndex.GetType(),
                550,
                @"[In RopOpenFolder ROP Success Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R551");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R551
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenFolderResponse.OutputHandleIndex,
                551,
                @"[In RopOpenFolder ROP Success Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R552");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R552
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenFolderResponse.ReturnValue.GetType(),
                552,
                @"[In RopOpenFolder ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer .");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R554");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R554
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropOpenFolderResponse.ReturnValue,
                554,
                @"[In RopOpenFolder ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R555,HasRules:{0}", ropOpenFolderResponse.HasRules);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R555
            // If HasRules is byte type and can be boolean type representation, then verify R555
            bool isVerifyR555 = (typeof(byte) == ropOpenFolderResponse.HasRules.GetType()) &&
                                ((Convert.ToBoolean(ropOpenFolderResponse.HasRules) == true) ||
                                (Convert.ToBoolean(ropOpenFolderResponse.HasRules) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR555,
                555,
                @"[In RopOpenFolder ROP Success Response Buffer] HasRules (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R557,IsGhosted:{0}", ropOpenFolderResponse.IsGhosted);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R557
            // If IsGhosted is byte type and can be boolean type representation, then verify R557
            bool isVerifyR557 = (typeof(byte) == ropOpenFolderResponse.IsGhosted.GetType()) &&
                                ((Convert.ToBoolean(ropOpenFolderResponse.IsGhosted) == true) ||
                                (Convert.ToBoolean(ropOpenFolderResponse.IsGhosted) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR557,
                557,
                @"[In RopOpenFolder ROP Success Response Buffer] IsGhosted (1 byte): A Boolean.");

            // The field ServerCount,CheapServerCount and Servers are present if IsGhosted is non-zero, and are not present if IsGhosted is zero. 
            if (ropOpenFolderResponse.IsGhosted > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R559");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R559
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropOpenFolderResponse.ServerCount.GetType(),
                    559,
                    @"[In RopOpenFolder ROP Success Response Buffer] ServerCount (2 bytes): An unsigned integer .");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R563");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R563
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropOpenFolderResponse.CheapServerCount.GetType(),
                    563,
                    @"[In RopOpenFolder ROP Success Response Buffer] CheapServerCount (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R567,the length of the Servers is {0}", ropOpenFolderResponse.Servers.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R567
                // Assume the result is true. If one of List is not null-terminated ASCII string, then false.
                bool isVerifyR567 = true;
                foreach (string s in ropOpenFolderResponse.Servers)
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(s);
                    isVerifyR567 = isVerifyR567 && this.IsNullTerminatedASCIIStr(buffer);

                    // If one of List is not null-terminated ASCII string then break iteration
                    if (!isVerifyR567)
                    {
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR567,
                    567,
                    @"[In RopOpenFolder ROP Success Response Buffer] Servers (variable): A List of null-terminated ASCII strings.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R570");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R570
                Site.CaptureRequirementIfAreEqual<ushort>(
                    (ushort)ropOpenFolderResponse.ServerCount,
                    (ushort)ropOpenFolderResponse.Servers.Length,
                    570,
                    @"[In RopOpenFolder ROP Success Response Buffer] Servers (variable): The number of strings contained in this field is specified by the ServerCount field.");
            }
        }

        /// <summary>
        /// Verify RopOpenFolder Failure Response
        /// </summary>
        /// <param name="ropOpenFolderResponse">The failure response of RopOpenFolder request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenFolder request</param>
        private void VerifyRopOpenFolderFailureResponse(RopOpenFolderResponse ropOpenFolderResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R572");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R572
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenFolderResponse.RopId.GetType(),
                572,
                @"[In RopOpenFolder ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R574");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R574
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenFolder,
                ropOpenFolderResponse.RopId,
                574,
                @"[In RopOpenFolder ROP Failure Response Buffer] RopId (1 byte): For this operation[RopOpenFolder], this field[RopId (1 byte)] is set to 0x02.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R575");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R575
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenFolderResponse.OutputHandleIndex.GetType(),
                575,
                @"[In RopOpenFolder ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R576");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R576
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenFolderResponse.OutputHandleIndex,
                576,
                @"[In RopOpenFolder ROP Failure Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R577");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R577
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenFolderResponse.ReturnValue.GetType(),
                577,
                @"[In RopOpenFolder ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R579");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R579
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropOpenFolderResponse.ReturnValue,
                579,
                @"[In RopOpenFolder ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopCreateFolder
        /// <summary>
        /// Verify RopCreateFolder Success Response
        /// </summary>
        /// <param name="ropCreateFolderResponse">The success response of RopCreateFolder request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateFolder request</param>
        private void VerifyRopCreateFolderSuccessResponse(RopCreateFolderResponse ropCreateFolderResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R609");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R609
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateFolderResponse.RopId.GetType(),
                609,
                @"[In RopCreateFolder ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R611");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R611
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateFolder,
                ropCreateFolderResponse.RopId,
                611,
                @"[In RopCreateFolder ROP Success Response Buffer] RopId (1 byte): For this operation[RopCreateFolder], this field[RopId (1 byte)] is set to 0x1C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R612");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R612
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateFolderResponse.OutputHandleIndex.GetType(),
                612,
                @"[In RopCreateFolder ROP Success Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R613");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R613
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateFolderResponse.OutputHandleIndex,
                613,
                @"[In RopCreateFolder ROP Success Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R614");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R614
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateFolderResponse.ReturnValue.GetType(),
                614,
                @"[In RopCreateFolder ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R616");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R616
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCreateFolderResponse.ReturnValue,
                616,
                @"[In RopCreateFolder ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R617");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R617
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropCreateFolderResponse.FolderId),
                617,
                @"[In RopCreateFolder ROP Success Response Buffer] FolderId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R619,IsExistingFolder:{0}", ropCreateFolderResponse.IsExistingFolder);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R619
            // If IsExistingFolder is byte type and can be boolean type representation, then verify R619
            bool isVerifyR619 = (typeof(byte) == ropCreateFolderResponse.IsExistingFolder.GetType()) &&
                                ((Convert.ToBoolean(ropCreateFolderResponse.IsExistingFolder) == true) ||
                                (Convert.ToBoolean(ropCreateFolderResponse.IsExistingFolder) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR619,
                619,
                @"[In RopCreateFolder ROP Success Response Buffer] IsExistingFolder (1 byte): A Boolean.");

            if (ropCreateFolderResponse.HasRules != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R621,HasRules:{0}", ropCreateFolderResponse.HasRules);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R621
                // If HasRules is byte type and can be boolean type representation, then verify R621
                bool isVerifyR621 = (typeof(byte) == ropCreateFolderResponse.HasRules.GetType()) &&
                                    ((Convert.ToBoolean(ropCreateFolderResponse.HasRules) == true) ||
                                    (Convert.ToBoolean(ropCreateFolderResponse.HasRules) == false));

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR621,
                    621,
                    @"[In RopCreateFolder ROP Success Response Buffer] HasRules (1 byte): A Boolean.");
            }

            if (ropCreateFolderResponse.IsGhosted != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R625,IsGhosted:{0}", ropCreateFolderResponse.IsGhosted);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R625
                // If IsGhosted is byte type and can be boolean type representation, then verify R625
                bool isVerifyR625 = (typeof(byte) == ropCreateFolderResponse.IsGhosted.GetType()) &&
                                    ((Convert.ToBoolean(ropCreateFolderResponse.IsGhosted) == true) ||
                                    (Convert.ToBoolean(ropCreateFolderResponse.IsGhosted) == false));

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR625,
                    625,
                    @"[In RopCreateFolder ROP Success Response Buffer] IsGhosted (1 byte): A Boolean.");
            }

            // The field ServerCount,CheapServerCount and Servers are present if IsGhosted is non-zero, and are not present if IsGhosted is zero. 
            if (ropCreateFolderResponse.IsGhosted != null && ropCreateFolderResponse.IsGhosted != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R629");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R629
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropCreateFolderResponse.ServerCount.GetType(),
                    629,
                    @"[In RopCreateFolder ROP Success Response Buffer] ServerCount (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R633");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R633
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropCreateFolderResponse.CheapServerCount.GetType(),
                    633,
                    @"[In RopCreateFolder ROP Success Response Buffer] CheapServerCount (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R637,the length of the Servers is {0}", ropCreateFolderResponse.Servers.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R637
                // Assume the result is true. If one of List is not null-terminated ASCII string, then false.
                bool isVerifyR637 = true;
                foreach (string s in ropCreateFolderResponse.Servers)
                {
                    byte[] buffer = Encoding.ASCII.GetBytes(s);
                    isVerifyR637 = isVerifyR637 && this.IsNullTerminatedASCIIStr(buffer);

                    // If one of List is not null-terminated ASCII string then break iteration
                    if (!isVerifyR637)
                    {
                        break;
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR637,
                    637,
                    @"[In RopCreateFolder ROP Success Response Buffer] Servers (variable):A list of null-terminated ASCII strings.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R640");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R640
                Site.CaptureRequirementIfAreEqual<ushort>(
                    (ushort)ropCreateFolderResponse.ServerCount,
                    (ushort)ropCreateFolderResponse.Servers.Length,
                    640,
                    @"[In RopCreateFolder ROP Success Response Buffer] Servers (variable): The number of strings contained in this field is specified by the ServersCount field.");
            }
        }

        /// <summary>
        /// Verify RopCreateFolder Failure Response
        /// </summary>
        /// <param name="ropCreateFolderResponse">The failure response of RopCreateFolder request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateFolder request</param>
        private void VerifyRopCreateFolderFailureResponse(RopCreateFolderResponse ropCreateFolderResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R645");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R645
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateFolderResponse.RopId.GetType(),
                645,
                @"[In RopCreateFolder ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R647");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R647
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateFolder,
                ropCreateFolderResponse.RopId,
                647,
                @"[In RopCreateFolder ROP Failure Response Buffer] RopId (1 byte): For this operation[RopCreateFolder], this field[RopId (1 byte)] is set to 0x1C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R648");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R648
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateFolderResponse.OutputHandleIndex.GetType(),
                648,
                @"[In RopCreateFolder ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R649");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R649
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateFolderResponse.OutputHandleIndex,
                649,
                @"[In RopCreateFolder ROP Failure Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R650");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R650
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateFolderResponse.ReturnValue.GetType(),
                650,
                @"[In RopCreateFolder ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R652");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R652
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropCreateFolderResponse.ReturnValue,
                652,
                @"[In RopCreateFolder ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopDeleteFolder
        /// <summary>
        /// Verify RopDeleteFolder Response
        /// </summary>
        /// <param name="ropDeleteFolderResponse">The response of RopDeleteFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeleteFolder request</param>
        private void VerifyRopDeleteFolderResponse(RopDeleteFolderResponse ropDeleteFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R667");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R667
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteFolderResponse.RopId.GetType(),
                667,
                @"[In RopDeleteFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R669");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R669
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeleteFolder,
                ropDeleteFolderResponse.RopId,
                669,
                @"[In RopDeleteFolder ROP Response Buffer] RopId (1 byte): For this operation[RopDeleteFolder], this field[RopId (1 byte)] is set to 0x1D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R670");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R670
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteFolderResponse.InputHandleIndex.GetType(),
                670,
                @"[In RopDeleteFolder ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R671");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R671
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeleteFolderResponse.InputHandleIndex,
                671,
                @"[In RopDeleteFolder ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R672");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R672
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeleteFolderResponse.ReturnValue.GetType(),
                672,
                @"[In RopDeleteFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R674,PartialCompletion:{0}", ropDeleteFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R674
            // If PartialCompletion is byte type and can be boolean type representation, then verify R674
            bool isVerifyR674 = (typeof(byte) == ropDeleteFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropDeleteFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropDeleteFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR674,
                674,
                @"[In Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopSetSearchCriteria
        /// <summary>
        /// Verify RopSetSearchCriteria Response
        /// </summary>
        /// <param name="ropSetSearchCriteriaResponse">The response of RopSetSearchCriteria request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetSearchCriteria request</param>
        private void VerifyRopSetSearchCriteriaResponse(RopSetSearchCriteriaResponse ropSetSearchCriteriaResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R698");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R698
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetSearchCriteriaResponse.RopId.GetType(),
                698,
                @"[In RopSetSearchCriteria ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R700");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R700
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetSearchCriteria,
                ropSetSearchCriteriaResponse.RopId,
                700,
                @"[In RopSetSearchCriteria ROP Response Buffer] RopId (1 byte): For this operation[RopSetSearchCriteria], this field[RopId (1 byte)] is set to 0x30.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R701");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R701
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetSearchCriteriaResponse.InputHandleIndex.GetType(),
                701,
                @"[In RopSetSearchCriteria ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R702");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R702
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetSearchCriteriaResponse.InputHandleIndex,
                702,
                @"[In RopSetSearchCriteria ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R703");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R703
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetSearchCriteriaResponse.ReturnValue.GetType(),
                703,
                @"[In RopSetSearchCriteria ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetSearchCriteria
        /// <summary>
        /// Verify RopGetSearchCriteria Success Response
        /// </summary>
        /// <param name="ropGetSearchCriteriaResponse">The success response of RopGetSearchCriteria request</param>
        /// <param name="logonId">The field of LogonId in RopGetSearchCriteria request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetSearchCriteria request</param>
        private void VerifyRopGetSearchCriteriaSuccessResponse(RopGetSearchCriteriaResponse ropGetSearchCriteriaResponse, byte logonId, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R720");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R720
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetSearchCriteriaResponse.RopId.GetType(),
                720,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R722");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R722
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetSearchCriteria,
                ropGetSearchCriteriaResponse.RopId,
                722,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetSearchCriteria], this field[RopId (1 byte)] is set to 0x31.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R723");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R723
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetSearchCriteriaResponse.InputHandleIndex.GetType(),
                723,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R724");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R724
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetSearchCriteriaResponse.InputHandleIndex,
                724,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R725");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R725
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetSearchCriteriaResponse.ReturnValue.GetType(),
                725,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R727");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R727
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetSearchCriteriaResponse.ReturnValue,
                727,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field[ReturnValue (4 bytes)] is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R728");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R728
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetSearchCriteriaResponse.RestrictionDataSize.GetType(),
                728,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] RestrictionDataSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R734");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R734
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetSearchCriteriaResponse.LogonId.GetType(),
                734,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] LogonId (1 byte): ): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R735");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R735
            Site.CaptureRequirementIfAreEqual<byte>(
                logonId,
                ropGetSearchCriteriaResponse.LogonId,
                735,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] LogonId (1 byte): This field MUST be set to the value of the LogonId field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R736");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R736
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetSearchCriteriaResponse.FolderIdCount.GetType(),
                736,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] FolderIdCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R738, the actual type of FolderIds is:{0}", ropGetSearchCriteriaResponse.FolderIds.GetType().Name);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R738
            // If FolderIds is long array or ulong array, can be array of 64-bit 
            bool isVerifyR738 = (typeof(long[]) == ropGetSearchCriteriaResponse.FolderIds.GetType()) ||
                                (typeof(ulong[]) == ropGetSearchCriteriaResponse.FolderIds.GetType());

            Site.CaptureRequirementIfIsTrue(
                isVerifyR738,
                738,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] FolderIds (variable): An array of 64-bit identifiers.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R740");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R740
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropGetSearchCriteriaResponse.FolderIdCount,
                (ushort)ropGetSearchCriteriaResponse.FolderIds.Length,
                740,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] FolderIds (variable): The number of identifiers contained in this field is specified by the FolderIdCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R741");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R741
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ropGetSearchCriteriaResponse.SearchFlags),
                741,
                @"[In RopGetSearchCriteria ROP Success Response Buffer] SearchFlags (4 bytes): ): A flags structure.");
        }

        /// <summary>
        /// Verify RopGetSearchCriteria Failure Response
        /// </summary>
        /// <param name="ropGetSearchCriteriaResponse">The failure response of RopGetSearchCriteria request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetSearchCriteria request</param>
        private void VerifyRopGetSearchCriteriaFailureResponse(RopGetSearchCriteriaResponse ropGetSearchCriteriaResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R743");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R743
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetSearchCriteriaResponse.RopId.GetType(),
                743,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R745");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R745
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetSearchCriteria,
                ropGetSearchCriteriaResponse.RopId,
                745,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetSearchCriteria], this field[RopId (1 byte)] is set to 0x31.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R746");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R746
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetSearchCriteriaResponse.InputHandleIndex.GetType(),
                746,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R747");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R747
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetSearchCriteriaResponse.InputHandleIndex,
                747,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R748");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R748
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetSearchCriteriaResponse.ReturnValue.GetType(),
                748,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R750");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R750
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetSearchCriteriaResponse.ReturnValue,
                750,
                @"[In RopGetSearchCriteria ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopMoveCopyMessages
        /// <summary>
        /// Verify RopMoveCopyMessages Response
        /// </summary>
        /// <param name="ropMoveCopyMessageResponse">The response of RopMoveCopyMessages request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopMoveCopyMessages request</param>
        private void VerifyRopMoveCopyMessagesResponse(RopMoveCopyMessagesResponse ropMoveCopyMessageResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R771");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R771
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveCopyMessageResponse.RopId.GetType(),
                771,
                @"[In RopMoveCopyMessages ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R773");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R773
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopMoveCopyMessages,
                ropMoveCopyMessageResponse.RopId,
                773,
                @"[In RopMoveCopyMessages ROP Response Buffer] RopId (1 byte): For this operation[RopMoveCopyMessages], this field[RopId (1 byte)] is set to 0x33.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R774");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R774
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveCopyMessageResponse.SourceHandleIndex.GetType(),
                774,
                @"[In RopMoveCopyMessages ROP Response Buffer] SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R775");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R775
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropMoveCopyMessageResponse.SourceHandleIndex,
                775,
                @"[In RopMoveCopyMessages ROP Response Buffer] SourceHandleIndex (1 byte): This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R776");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R776
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveCopyMessageResponse.ReturnValue.GetType(),
                776,
                @"[In RopMoveCopyMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R778");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R778
            Site.CaptureRequirementIfAreNotEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropMoveCopyMessageResponse.ReturnValue,
                778,
                @"[In RopMoveCopyMessages ROP Response Buffer] ReturnValue (4 bytes): For this response[Response Buffer], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R779,PartialCompletion:{0}", ropMoveCopyMessageResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R779
            // If PartialCompletion is byte type and can be boolean type representation, then verify R779
            bool isVerifyR779 = (typeof(byte) == ropMoveCopyMessageResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropMoveCopyMessageResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropMoveCopyMessageResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR779,
                779,
                @"[In RopMoveCopyMessages ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }

        /// <summary>
        /// Verify RopMoveCopyMessages Null Destination Failure Response
        /// </summary>
        /// <param name="ropMoveCopyMessageResponse">The null destination failure response of RopMoveCopyMessages request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopMoveCopyMessages request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopMoveCopyMessages request</param>
        private void VerifyRopMoveCopyMessagesNullDestinationFailureResponse(
            RopMoveCopyMessagesResponse ropMoveCopyMessageResponse,
            byte sourceHandleIndex,
            uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R781");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R781
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveCopyMessageResponse.RopId.GetType(),
                781,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R783");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R783
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopMoveCopyMessages,
                ropMoveCopyMessageResponse.RopId,
                783,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] RopId (1 byte): For this operation[RopMoveCopyMessages], this field[RopId (1 byte)] is set to 0x33.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R784");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R784
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveCopyMessageResponse.SourceHandleIndex.GetType(),
                784,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R785");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R785
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropMoveCopyMessageResponse.SourceHandleIndex,
                785,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R786");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R786
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveCopyMessageResponse.ReturnValue.GetType(),
                786,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R789");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R789
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropMoveCopyMessageResponse.ReturnValue,
                789,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): For this response[Null Destination Failure Response], this field[ReturnValue (4 bytes)] is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R790");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R790
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveCopyMessageResponse.DestHandleIndex.GetType(),
                790,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R791");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R791
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropMoveCopyMessageResponse.DestHandleIndex,
                791,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): This index MUST be set to the value specified in the DestHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R792,PartialCompletion:{0}", ropMoveCopyMessageResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R792
            // If PartialCompletion is byte type and can be boolean type representation, then verify R792
            bool isVerifyR792 = (typeof(byte) == ropMoveCopyMessageResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropMoveCopyMessageResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropMoveCopyMessageResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR792,
                792,
                @"[In RopMoveCopyMessages ROP Null Destination Failure Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }

        #endregion

        #region RopMoveFolder
        /// <summary>
        /// Verify RopMoveFolder Response
        /// </summary>
        /// <param name="ropMoveFolderResponse">The response of RopMoveFolder request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopMoveFolder request</param>
        private void VerifyRopMoveFolderResponse(RopMoveFolderResponse ropMoveFolderResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R815");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R815
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveFolderResponse.RopId.GetType(),
                815,
                @"[In RopMoveFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R817");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R817
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopMoveFolder,
                ropMoveFolderResponse.RopId,
                817,
                @"[In RopMoveFolder ROP Response Buffer] RopId (1 byte): For this operation[RopMoveFolder], this field is set to 0x35.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R818");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R818
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveFolderResponse.SourceHandleIndex.GetType(),
                818,
                @"[In RopMoveFolder ROP Response Buffer] SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R819");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R819
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropMoveFolderResponse.SourceHandleIndex,
                819,
                @"[In RopMoveFolder ROP Response Buffer] SourceHandleIndex (1 byte): This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R820");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R820
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveFolderResponse.ReturnValue.GetType(),
                820,
                @"[In RopMoveFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R822");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R822
            Site.CaptureRequirementIfAreNotEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropMoveFolderResponse.ReturnValue,
                822,
                @"[In RopMoveFolder ROP Response Buffer] ReturnValue (4 bytes): For this response[Response Buffer], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R823,PartialCompletion:{0}", ropMoveFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R823
            // If PartialCompletion is byte type and can be boolean type, then verify R823
            bool isVerifyR823 = (typeof(byte) == ropMoveFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropMoveFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropMoveFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR823,
                823,
                @"[In RopMoveFolder ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }

        /// <summary>
        /// Verify RopMoveFolder Null Destination Failure Response
        /// </summary>
        /// <param name="ropMoveFolderResponse">The response of RopMoveFolder request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopMoveFolder request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopMoveFolder request</param>
        private void VerifyRopMoveFolderNullDestinationFailureResponse(RopMoveFolderResponse ropMoveFolderResponse, byte sourceHandleIndex, uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R825");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R825
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveFolderResponse.RopId.GetType(),
                825,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] RopId (1 byte): ): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R827");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R827
            Site.CaptureRequirementIfAreEqual<byte>(
               (byte)RopId.RopMoveFolder,
                ropMoveFolderResponse.RopId,
                827,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] RopId (1 byte): For this operation[RopMoveFolder], this field[RopId (1 byte)] is set to 0x35.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R828");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R828
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropMoveFolderResponse.SourceHandleIndex.GetType(),
                828,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): ): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R829");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R829
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropMoveFolderResponse.SourceHandleIndex,
                829,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R830");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R830
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveFolderResponse.ReturnValue.GetType(),
                830,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R832");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R832
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropMoveFolderResponse.ReturnValue,
                832,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): For this response[Null Destination Failure Response], this field[ReturnValue (4 bytes)] is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R833");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R833
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropMoveFolderResponse.DestHandleIndex.GetType(),
                833,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R834");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R834
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropMoveFolderResponse.DestHandleIndex,
                834,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): This index MUST be set to the value specified in the DestHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R835,PartialCompletion:{0}", ropMoveFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R835
            // If PartialCompletion is byte type and can be boolean type representation, then verify R835
            bool isVerifyR835 = (typeof(byte) == ropMoveFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropMoveFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropMoveFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR835,
                835,
                @"[In RopMoveFolder ROP Null Destination Failure Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopCopyFolder
        /// <summary>
        /// Verify RopCopyFolder Response
        /// </summary>
        /// <param name="ropCopyFolderResponse">The response of RopCopyFolder request</param>
        /// <param name="sourceHandleIndex">The field SourceHandleIndex in RopCopyFolder request</param>
        private void VerifyRopCopyFolderResponse(RopCopyFolderResponse ropCopyFolderResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R860");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R860
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyFolderResponse.RopId.GetType(),
                860,
                @"[In RopCopyFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R862");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R862
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyFolder,
                ropCopyFolderResponse.RopId,
                862,
                @"[In RopCopyFolder ROP Response Buffer] RopId (1 byte): For this operation[RopCopyFolder], this field[RopId (1 byte)] is set to 0x36.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R863");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R863
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyFolderResponse.SourceHandleIndex.GetType(),
                863,
                @"[In RopCopyFolder ROP Response Buffer] SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R864");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R864
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyFolderResponse.SourceHandleIndex,
                864,
                @"[In RopCopyFolder ROP Response Buffer] SourceHandleIndex (1 byte): This index[SourceHandleIndex (1 byte)] MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R865");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R865
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyFolderResponse.ReturnValue.GetType(),
                865,
                @"[In RopCopyFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R867");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R867
            Site.CaptureRequirementIfAreNotEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyFolderResponse.ReturnValue,
                867,
                @"[In RopCopyFolder ROP Response Buffer] ReturnValue (4 bytes): For this response[Response Buffer], this field[ReturnValue (4 bytes)] is set to a value other than 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R868,PartialCompletion:{0}", ropCopyFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R868
            // If PartialCompletion is byte type and can be boolean type representation, then verify R868
            bool isVerifyR868 = (typeof(byte) == ropCopyFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropCopyFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropCopyFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR868,
                868,
                @"[In RopCopyFolder ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }

        /// <summary>
        /// Verify RopCopyFolder Null Destination Failure Response
        /// </summary>
        /// <param name="ropCopyFolderResponse">The response of RopCopyFolder request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyFolder request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopCopyFolder request</param>
        private void VerifyRopCopyFolderNullDestinationFailureResponse(RopCopyFolderResponse ropCopyFolderResponse, byte sourceHandleIndex, uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R870");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R870
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyFolderResponse.RopId.GetType(),
                870,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R872");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R872
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyFolder,
                ropCopyFolderResponse.RopId,
                872,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] RopId (1 byte): For this operation[RopCopyFolder], this field[RopId (1 byte)] is set to 0x36.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R873");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R873
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyFolderResponse.SourceHandleIndex.GetType(),
                873,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R874");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R874
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyFolderResponse.SourceHandleIndex,
                874,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] SourceHandleIndex (1 byte): This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R875");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R875
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyFolderResponse.ReturnValue.GetType(),
                875,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R877");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R877
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyFolderResponse.ReturnValue,
                877,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] ReturnValue (4 bytes): For this response[Null Destination Failure Response], this field is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R878");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R878
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyFolderResponse.DestHandleIndex.GetType(),
                878,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R879");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R879
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropCopyFolderResponse.DestHandleIndex,
                879,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] DestHandleIndex (4 bytes): This index[DestHandleIndex (4 bytes)] MUST be set to the value specified in the DestHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R880,PartialCompletion:{0}", ropCopyFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R880
            // If PartialCompletion is byte type and can be boolean type representation, then verify R880
            bool isVerifyR880 = (typeof(byte) == ropCopyFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropCopyFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropCopyFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR880,
                880,
                @"[In RopCopyFolder ROP Null Destination Failure Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopEmptyFolder
        /// <summary>
        /// Verify RopEmptyFolder Response
        /// </summary>
        /// <param name="ropEmptyFolderResponse">The response of RopEmptyFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopEmptyFolder request</param>
        private void VerifyRopEmptyFolderResponse(RopEmptyFolderResponse ropEmptyFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R895");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R895
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropEmptyFolderResponse.RopId.GetType(),
                895,
                @"[In RopEmptyFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R897");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R897
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopEmptyFolder,
                ropEmptyFolderResponse.RopId,
                897,
                @"[In RopEmptyFolder ROP Response Buffer] RopId (1 byte): For this operation[RopEmptyFolder], this field[RopId (1 byte)] is set to 0x58.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R898");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R898
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropEmptyFolderResponse.InputHandleIndex.GetType(),
                898,
                @"[In RopEmptyFolder ROP Response Buffer] InputHandleIndex (1 byte):An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R899");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R899
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropEmptyFolderResponse.InputHandleIndex,
                899,
                @"[In RopEmptyFolder ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R900");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R900
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropEmptyFolderResponse.ReturnValue.GetType(),
                900,
                @"[In RopEmptyFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R902,PartialCompletion:{0}", ropEmptyFolderResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R902
            // If PartialCompletion is byte type and can be boolean type representation, then verify R902
            bool isVerifyR902 = (typeof(byte) == ropEmptyFolderResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropEmptyFolderResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropEmptyFolderResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR902,
                902,
                @"[In RopEmptyFolder ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopHardDeleteMessagesAndSubfolders
        /// <summary>
        /// Verify RopHardDeleteMessagesAndSubfolders Response
        /// </summary>
        /// <param name="ropHardDeleteMessagesAndSubfoldersResponse">The response of RopHardDeleteMessagesAndSubfolders request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopHardDeleteMessagesAndSubfolders request</param>
        private void VerifyRopHardDeleteMessagesAndSubfoldersResponse(RopHardDeleteMessagesAndSubfoldersResponse ropHardDeleteMessagesAndSubfoldersResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R917");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R917
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropHardDeleteMessagesAndSubfoldersResponse.RopId.GetType(),
                917,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R919");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R919
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopHardDeleteMessagesAndSubfolders,
                ropHardDeleteMessagesAndSubfoldersResponse.RopId,
                919,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] RopId (1 byte): For this operation[RopHardDeleteMessagesAndSubfolders], this field[RopId (1 byte)] is set to 0x92.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R920");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R920
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropHardDeleteMessagesAndSubfoldersResponse.InputHandleIndex.GetType(),
                920,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R921");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R921
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropHardDeleteMessagesAndSubfoldersResponse.InputHandleIndex,
                921,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] InputHandleIndex (1 byte): This indexMUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R922");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R922
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropHardDeleteMessagesAndSubfoldersResponse.ReturnValue.GetType(),
                922,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R924,PartialCompletion:{0}", ropHardDeleteMessagesAndSubfoldersResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R924
            // If PartialCompletion is byte type and can be boolean type representation, then verify R924
            bool isVerifyR924 = (typeof(byte) == ropHardDeleteMessagesAndSubfoldersResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropHardDeleteMessagesAndSubfoldersResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropHardDeleteMessagesAndSubfoldersResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR924,
                924,
                @"[In RopHardDeleteMessagesAndSubfolders ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopDeleteMessages
        /// <summary>
        /// Verify RopDeleteMessages Response
        /// </summary>
        /// <param name="ropDeleteMessageResponse">The response of RopDeleteMessages request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeleteMessages request</param>
        private void VerifyRopDeleteMessagesResponse(RopDeleteMessagesResponse ropDeleteMessageResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R944");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R944
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteMessageResponse.RopId.GetType(),
                944,
                @"[In RopDeleteMessages ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R946");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R946
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeleteMessages,
                ropDeleteMessageResponse.RopId,
                946,
                @"[In RopDeleteMessages ROP Response Buffer] RopId (1 byte): For this operation[RopDeleteMessages], this field[RopId (1 byte)] is set to 0x1E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R947");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R947
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteMessageResponse.InputHandleIndex.GetType(),
                947,
                @"[In RopDeleteMessages ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R948");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R948
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeleteMessageResponse.InputHandleIndex,
                948,
                @"[In RopDeleteMessages ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R949");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R949
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeleteMessageResponse.ReturnValue.GetType(),
                949,
                @"[In RopDeleteMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R951,PartialCompletion:{0}", ropDeleteMessageResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R951
            // If PartialCompletion is byte type and can be boolean type representation, then verify R951
            bool isVerifyR951 = (typeof(byte) == ropDeleteMessageResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropDeleteMessageResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropDeleteMessageResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR951,
                951,
                @"[In RopDeleteMessages ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopHardDeleteMessages
        /// <summary>
        /// Verify RopHardDeleteMessages Response
        /// </summary>
        /// <param name="ropHardDeleteMessagesResponse">The response of RopHardDeleteMessages request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopHardDeleteMessages request</param>
        private void VerifyRopHardDeleteMessagesResponse(RopHardDeleteMessagesResponse ropHardDeleteMessagesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R971");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R971
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropHardDeleteMessagesResponse.RopId.GetType(),
                971,
                @"[In RopHardDeleteMessages ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R973");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R973
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopHardDeleteMessages,
                ropHardDeleteMessagesResponse.RopId,
                973,
                @"[In RopHardDeleteMessages ROP Response Buffer] RopId (1 byte): For this operation[RopHardDeleteMessages], this field[RopId (1 byte)] is set to 0x91.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R974");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R974
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropHardDeleteMessagesResponse.InputHandleIndex.GetType(),
                974,
                @"[In RopHardDeleteMessages ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R975");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R975
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropHardDeleteMessagesResponse.InputHandleIndex,
                975,
                @"[In RopHardDeleteMessages ROP Response Buffer] InputHandleIndex (1 byte): This index that MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R976");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R976
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropHardDeleteMessagesResponse.ReturnValue.GetType(),
                976,
                @"[In RopHardDeleteMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R978,PartialCompletion:{0}", ropHardDeleteMessagesResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R978
            // If PartialCompletion is byte type and can be boolean type representation, then verify R978
            bool isVerifyR978 = (typeof(byte) == ropHardDeleteMessagesResponse.PartialCompletion.GetType()) &&
                                ((Convert.ToBoolean(ropHardDeleteMessagesResponse.PartialCompletion) == true) ||
                                (Convert.ToBoolean(ropHardDeleteMessagesResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR978,
                978,
                @"[In RopHardDeleteMessages ROP Response Buffer] PartialCompletion (1 byte): A Boolean.");
        }
        #endregion

        #region RopGetHierarchyTable
        /// <summary>
        /// Verify RopGetHierarchyTable Success Response
        /// </summary>
        /// <param name="ropGetHierarchyTableResponse">The success response of RopGetHierarchyTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetHierarchyTable request</param>
        private void VerifyRopGetHierarchyTableSuccessResponse(RopGetHierarchyTableResponse ropGetHierarchyTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R993");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R993
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetHierarchyTableResponse.RopId.GetType(),
                993,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R995");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R995
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetHierarchyTable,
                ropGetHierarchyTableResponse.RopId,
                995,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetHierarchyTable], this field[RopId (1 byte)] is set to 0x04.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R996");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R996
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetHierarchyTableResponse.OutputHandleIndex.GetType(),
                996,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R997");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R997
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetHierarchyTableResponse.OutputHandleIndex,
                997,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R998");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R998
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetHierarchyTableResponse.ReturnValue.GetType(),
                998,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1000");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1000
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetHierarchyTableResponse.ReturnValue,
                1000,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1001");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1001
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetHierarchyTableResponse.RowCount.GetType(),
                1001,
                @"[In RopGetHierarchyTable ROP Success Response Buffer] RowCount (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopGetHierarchyTable Failure Response
        /// </summary>
        /// <param name="ropGetHierarchyTableResponse">The failure response of RopGetHierarchyTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetHierarchyTable request</param>
        private void VerifyRopGetHierarchyTableFailureResponse(RopGetHierarchyTableResponse ropGetHierarchyTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1003");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1003
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetHierarchyTableResponse.RopId.GetType(),
                1003,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1005");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1005
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetHierarchyTable,
                ropGetHierarchyTableResponse.RopId,
                1005,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetHierarchyTable], this field[RopId (1 byte)] is set to 0x04.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1006");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1006
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetHierarchyTableResponse.OutputHandleIndex.GetType(),
                1006,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1007");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1007
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetHierarchyTableResponse.OutputHandleIndex,
                1007,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1008");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1008
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetHierarchyTableResponse.ReturnValue.GetType(),
                1008,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1010");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1010
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetHierarchyTableResponse.ReturnValue,
                1010,
                @"[In RopGetHierarchyTable ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetContentsTable
        /// <summary>
        /// Verify RopGetContentsTable Success Response
        /// </summary>
        /// <param name="ropGetContentsTableResponse">The success response of RopGetContentsTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetContentsTable request</param>
        private void VerifyRopGetContentsTableSuccessResponse(RopGetContentsTableResponse ropGetContentsTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1026");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1026
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetContentsTableResponse.RopId.GetType(),
                1026,
                @"[In RopGetContentsTable ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1028");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1028
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetContentsTable,
                ropGetContentsTableResponse.RopId,
                1028,
                @"[In RopGetContentsTable ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetContentsTable], this field[RopId (1 byte)] is set to 0x05.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1029");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1029
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetContentsTableResponse.OutputHandleIndex.GetType(),
                1029,
                @"[In RopGetContentsTable ROP Success Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1030");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1030
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetContentsTableResponse.OutputHandleIndex,
                1030,
                @"[In RopGetContentsTable ROP Success Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1031");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1031
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetContentsTableResponse.ReturnValue.GetType(),
                1031,
                @"[In RopGetContentsTable ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1033");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1033
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetContentsTableResponse.ReturnValue,
                1033,
                @"[In RopGetContentsTable ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1034");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1034
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetContentsTableResponse.RowCount.GetType(),
                1034,
                @"[In RopGetContentsTable ROP Success Response Buffer] RowCount (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopGetContentsTable Failure Response
        /// </summary>
        /// <param name="ropGetContentsTableResponse">The failure response of RopGetContentsTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetContentsTable request</param>
        private void VerifyRopGetContentsTableFailureResponse(RopGetContentsTableResponse ropGetContentsTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1036");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1036
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetContentsTableResponse.RopId.GetType(),
                1036,
                @"[In RopGetContentsTable ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1038");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1038
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetContentsTable,
                ropGetContentsTableResponse.RopId,
                1038,
                @"[In RopGetContentsTable ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetContentsTable], this field[RopId (1 byte)] is set to 0x05.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1039");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1039
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetContentsTableResponse.OutputHandleIndex.GetType(),
                1039,
                @"[In RopGetContentsTable ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1040");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1040
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetContentsTableResponse.OutputHandleIndex,
                1040,
                @"[In RopGetContentsTable ROP Failure Response Buffer] OutputHandleIndex (1 byte): This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1041");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1041
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetContentsTableResponse.ReturnValue.GetType(),
                1041,
                @"[In RopGetContentsTable ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1043");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1043
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetContentsTableResponse.ReturnValue,
                1043,
                @"[In RopGetContentsTable ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #endregion

        #region Table ROPs

        #region RopSetColumns
        /// <summary>
        /// Verify RopSetColumns Success Response
        /// </summary>
        /// <param name="ropSetColumnsResponse">The success response of RopSetColumns request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetColumns request</param>
        private void VerifyRopSetColumnsSuccessResponse(RopSetColumnsResponse ropSetColumnsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1063");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1063
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetColumnsResponse.RopId.GetType(),
                1063,
                @"[In RopSetColumns ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1065");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1065
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetColumns,
                ropSetColumnsResponse.RopId,
                1065,
                @"[In RopSetColumns ROP Success Response Buffer] RopId (1 byte): For this operation[RopSetColumns], this field[RopId (1 byte)] is set to 0x12.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1066");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1066
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetColumnsResponse.InputHandleIndex.GetType(),
                1066,
                @"[In RopSetColumns ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1067");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1067
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetColumnsResponse.InputHandleIndex,
                1067,
                @"[In RopSetColumns ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1068");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1068
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetColumnsResponse.ReturnValue.GetType(),
                1068,
                @"[In RopSetColumns ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1070");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1070
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetColumnsResponse.ReturnValue,
                1070,
                @"[In RopSetColumns ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1071");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1071
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetColumnsResponse.TableStatus.GetType(),
                1071,
                @"[In RopSetColumns ROP Success Response Buffer] TableStatus (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1072,TableStatus:{0}", ropSetColumnsResponse.TableStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1072
            bool isVerifyR1072 = ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatComplete)) == ((byte)TableStatus.TblstatComplete))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatSorting)) == ((byte)TableStatus.TblstatSorting))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatSortError)) == ((byte)TableStatus.TblstatSortError))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatSettingCols)) == ((byte)TableStatus.TblstatSettingCols))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatSetColError)) == ((byte)TableStatus.TblstatSetColError))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatRestricting)) == ((byte)TableStatus.TblstatRestricting))
                                 || ((ropSetColumnsResponse.TableStatus & ((byte)TableStatus.TblstatRestrictError)) == ((byte)TableStatus.TblstatRestrictError));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1072,
                1072,
                @"[In RopSetColumns ROP Success Response Buffer] TableStatus (1 byte): The possible values[the value of TBLSTAT_COMPLETE is 0x00, the value of TBLSTAT_SORTING is 0x09, the value of TBLSTAT_SORT_ERROR is 0x0A, the value of TBLSTAT_SETTING_COLS is 0x0B, the value of TBLSTAT_SETCOL_ERROR is 0x0D, the value of TBLSTAT_RESTRICTING is 0x0E and the value of TBLSTAT_RESTRICT_ERROR is 0x0F] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.1.3.");
        }

        /// <summary>
        /// Verify RopSetColumns Failure Response
        /// </summary>
        /// <param name="ropSetColumnsResponse">The failure response of RopSetColumns request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetColumns request</param>
        private void VerifyRopSetColumnsFailureResponse(RopSetColumnsResponse ropSetColumnsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1074");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1074
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetColumnsResponse.RopId.GetType(),
                1074,
                @"[In RopSetColumns ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1076");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1076
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetColumns,
                ropSetColumnsResponse.RopId,
                1076,
                @"[In RopSetColumns ROP Failure Response Buffer] RopId (1 byte): For this operation[RopSetColumns], this field[RopId (1 byte)] is set to 0x12.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1077");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1077
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetColumnsResponse.InputHandleIndex.GetType(),
                1077,
                @"[In RopSetColumns ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1078");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1078
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetColumnsResponse.InputHandleIndex,
                1078,
                @"[In RopSetColumns ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1079");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1079
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetColumnsResponse.ReturnValue.GetType(),
                1079,
                @"[In RopSetColumns ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1081");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1081
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetColumnsResponse.ReturnValue,
                1081,
                @"[In RopSetColumns ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSortTable
        /// <summary>
        /// Verify RopSortTable Success Response
        /// </summary>
        /// <param name="ropSortTableResponse">The response of RopSortTable request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSortTable request</param>
        private void VerifyRopSortTableSuccessResponse(RopSortTableResponse ropSortTableResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1104");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1104
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSortTableResponse.RopId.GetType(),
                1104,
                @"[In RopSortTable ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1106");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1106
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSortTable,
                ropSortTableResponse.RopId,
                1106,
                @"[In RopSortTable ROP Success Response Buffer] RopId (1 byte): For this operation[RopSortTable], this field[RopId (1 byte)] is set to 0x13.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1107");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1107
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSortTableResponse.InputHandleIndex.GetType(),
                1107,
                @"[In RopSortTable ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1108");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1108
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSortTableResponse.InputHandleIndex,
                1108,
                @"[In RopSortTable ROP Success Response Buffer] InputHandleIndex (1 byte): This index that MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1109");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1109
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSortTableResponse.ReturnValue.GetType(),
                1109,
                @"[In RopSortTable ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1111");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1111
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSortTableResponse.ReturnValue,
                1111,
                @"[In RopSortTable ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1112");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1112
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSortTableResponse.TableStatus.GetType(),
                1112,
                @"[In RopSortTable ROP Success Response Buffer] TableStatus (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1113,TableStatus:{0}", ropSortTableResponse.TableStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1113
            bool isVerifyR1113 = ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatComplete)) == ((byte)TableStatus.TblstatComplete))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatSorting)) == ((byte)TableStatus.TblstatSorting))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatSortError)) == ((byte)TableStatus.TblstatSortError))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatSettingCols)) == ((byte)TableStatus.TblstatSettingCols))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatSetColError)) == ((byte)TableStatus.TblstatSetColError))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatRestricting)) == ((byte)TableStatus.TblstatRestricting))
                                 || ((ropSortTableResponse.TableStatus & ((byte)TableStatus.TblstatRestrictError)) == ((byte)TableStatus.TblstatRestrictError));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1113,
                1113,
                @"[In RopSortTable ROP Success Response Buffer] TableStatus (1 byte): The possible values[the value of TBLSTAT_COMPLETE is 0x00, the value of TBLSTAT_SORTING is 0x09, the value of TBLSTAT_SORT_ERROR is 0x0A, the value of TBLSTAT_SETTING_COLS is 0x0B, the value of TBLSTAT_SETCOL_ERROR is 0x0D, the value of TBLSTAT_RESTRICTING is 0x0E and the value of TBLSTAT_RESTRICT_ERROR is 0x0F] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.1.3.");
        }

        /// <summary>
        /// Verify RopSortTable Failure Response
        /// </summary>
        /// <param name="ropSortTableResponse">The response of RopSortTable request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSortTable request</param>
        private void VerifyRopSortTableFailureResponse(RopSortTableResponse ropSortTableResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1115");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1115
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSortTableResponse.RopId.GetType(),
                1115,
                @"[In RopSortTable ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1117");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1117
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSortTable,
                ropSortTableResponse.RopId,
                1117,
                @"[In RopSortTable ROP Failure Response Buffer] RopId (1 byte): For this operation[RopSortTable], this field is set to 0x13.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1118");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1118
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSortTableResponse.InputHandleIndex.GetType(),
                1118,
                @"[In RopSortTable ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1119");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1119
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSortTableResponse.InputHandleIndex,
                1119,
                @"[In RopSortTable ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1120");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1120
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSortTableResponse.ReturnValue.GetType(),
                1120,
                @"[In RopSortTable ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1122");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1122
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSortTableResponse.ReturnValue,
                1122,
                @"[In RopSortTable ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopRestrict
        /// <summary>
        /// Verify RopRestrict Success Response
        /// </summary>
        /// <param name="ropRestrictResponse">The response of RopRestrict request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopRestrict request</param>
        private void VerifyRopRestrictSuccessResponse(RopRestrictResponse ropRestrictResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1141");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1141
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRestrictResponse.RopId.GetType(),
                1141,
                @"[In RopRestrict ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1143");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1143
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopRestrict,
                ropRestrictResponse.RopId,
                1143,
                @"[In RopRestrict ROP Success Response Buffer] RopId (1 byte): For this operation[RopRestrict], this field[RopId (1 byte)] is set to 0x14.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1144");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1144
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRestrictResponse.InputHandleIndex.GetType(),
                1144,
                @"[In RopRestrict ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1145");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1145
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropRestrictResponse.InputHandleIndex,
                1145,
                @"[In RopRestrict ROP Success Response Buffer] InputHandleIndex (1 byte): This index that MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1146");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1146
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropRestrictResponse.ReturnValue.GetType(),
                1146,
                @"[In RopRestrict ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1148");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1148
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropRestrictResponse.ReturnValue,
                1148,
                @"[In RopRestrict ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1149");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1149
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRestrictResponse.TableStatus.GetType(),
                1149,
                @"[In RopRestrict ROP Success Response Buffer] TableStatus (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1150,TableStatus:{0}", ropRestrictResponse.TableStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1150
            bool isVerifyR1150 = ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatComplete)) == ((byte)TableStatus.TblstatComplete))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatSorting)) == ((byte)TableStatus.TblstatSorting))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatSortError)) == ((byte)TableStatus.TblstatSortError))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatSettingCols)) == ((byte)TableStatus.TblstatSettingCols))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatSetColError)) == ((byte)TableStatus.TblstatSetColError))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatRestricting)) == ((byte)TableStatus.TblstatRestricting))
                                 || ((ropRestrictResponse.TableStatus & ((byte)TableStatus.TblstatRestrictError)) == ((byte)TableStatus.TblstatRestrictError));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1150,
                1150,
                @"[In RopRestrict ROP Success Response Buffer] TableStatus (1 byte): The possible values[the value of TBLSTAT_COMPLETE is 0x00, the value of TBLSTAT_SORTING is 0x09, the value of TBLSTAT_SORT_ERROR is 0x0A, the value of TBLSTAT_SETTING_COLS is 0x0B, the value of TBLSTAT_SETCOL_ERROR is 0x0D, the value of TBLSTAT_RESTRICTING is 0x0E and the value of TBLSTAT_RESTRICT_ERROR is 0x0F] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.1.3.");
        }

        /// <summary>
        /// Verify RopRestrict Failure Response
        /// </summary>
        /// <param name="ropRestrictResponse">The response of RopRestrict request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopRestrict request</param>
        private void VerifyRopRestrictFailureResponse(RopRestrictResponse ropRestrictResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1152");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1152
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRestrictResponse.RopId.GetType(),
                1152,
                @"[In RopRestrict ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1154");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1154
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopRestrict,
                ropRestrictResponse.RopId,
                1154,
                @"[In RopRestrict ROP Failure Response Buffer] RopId (1 byte): For this operation[RopRestrict], this field[RopId (1 byte)] is set to 0x14.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1155");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1155
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRestrictResponse.InputHandleIndex.GetType(),
                1155,
                @"[In RopRestrict ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1156");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1156
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropRestrictResponse.InputHandleIndex,
                1156,
                @"[In RopRestrict ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1157");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1157
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropRestrictResponse.ReturnValue.GetType(),
                1157,
                @"[In RopRestrict ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1159");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1159
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropRestrictResponse.ReturnValue,
                1159,
                @"[In RopRestrict ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopQueryRows
        /// <summary>
        /// Verify RopQueryRows Success Response
        /// </summary>
        /// <param name="ropQueryRowsResponse">The response of RopQueryRows request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryRows request</param>
        /// <param name="propertyTags">Array of PropertyTag structures specifies the property values that are visible in table rows, set by a RopSetColumns request</param>
        private void VerifyRopQueryRowsSuccessResponse(RopQueryRowsResponse ropQueryRowsResponse, byte inputHandleIndex, PropertyTag[] propertyTags)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1176");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1176
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryRowsResponse.RopId.GetType(),
                1176,
                @"[In RopQueryRows ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1178");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1178
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryRows,
                ropQueryRowsResponse.RopId,
                1178,
                @"[In RopQueryRows ROP Success Response Buffer] RopId (1 byte): For this operation[RopQueryRows], this field[RopId (1 byte)] is set to 0x15.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1179");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1179
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryRowsResponse.InputHandleIndex.GetType(),
                1179,
                @"[In RopQueryRows ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1180");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1180
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryRowsResponse.InputHandleIndex,
                1180,
                @"[In RopQueryRows ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1181");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryRowsResponse.ReturnValue.GetType(),
                1181,
                @"[In RopQueryRows ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1183");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1183
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropQueryRowsResponse.ReturnValue,
                1183,
                @"[In RopQueryRows ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1184");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1184
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryRowsResponse.Origin.GetType(),
                1184,
                @"[In RopQueryRows ROP Success Response Buffer] Origin (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1185,Origin:{0}", ropQueryRowsResponse.Origin);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1185
            bool isVerifyR1185 = ((ropQueryRowsResponse.Origin & (byte)Origin.Beginning) == (byte)Origin.Beginning)
                                 || ((ropQueryRowsResponse.Origin & (byte)Origin.Current) == (byte)Origin.Current)
                                 || ((ropQueryRowsResponse.Origin & (byte)Origin.End) == (byte)Origin.End);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1185,
                1185,
                @"[In RopQueryRows ROP Success Response Buffer] Origin (1 byte): The possible values[the value of BOOKMARK_BEGINNING is 0x00, the value of BOOKMARK_CURRENT is 0x01 and the value of BOOKMARK_END is 0x02] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.5.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1187");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1187
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropQueryRowsResponse.RowCount.GetType(),
                1187,
                @"[In RopQueryRows ROP Success Response Buffer] RowCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1189");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1189
            // The class PropertyRowSet is List of PropertRow, if the type of Rows is PropertyRowSet can verify R1189 
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyRowSet),
                ropQueryRowsResponse.RowData.GetType(),
                1189,
                @"[In RopQueryRows ROP Success Response Buffer] RowData (variable): A list of PropertyRow structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1190");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1190
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropQueryRowsResponse.RowCount,
                (ushort)ropQueryRowsResponse.RowData.Count,
                1190,
                @"[In RopQueryRows ROP Success Response Buffer] RowData (variable): The number of structures contained in this field is specified by the RowCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1192,RowData:{0}", ropQueryRowsResponse.RowData);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1192
            // Assume the result is true. If one of rows element's count is not equal to the columns count, then false.
            bool isVerifyR1192 = true;
            foreach (PropertyRow propertyRow in ropQueryRowsResponse.RowData.PropertyRows)
            {
                isVerifyR1192 = isVerifyR1192 && (propertyRow.PropertyValues.Count == propertyTags.Length);

                // If one row element's count is not equal to the columns count then break the iteration
                if (!isVerifyR1192)
                {
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1192,
                1192,
                @"[In RopQueryRows ROP Success Response Buffer] RowData (variable): The columns used for these rows were those previously set on this table by a RopSetColumns ROP request (section 2.2.5.1).");
        }

        /// <summary>
        /// Verify RopQueryRows Failure Response
        /// </summary>
        /// <param name="ropQueryRowsResponse">The response of RopQueryRows request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryRows request</param>
        private void VerifyRopQueryRowsFailureResponse(RopQueryRowsResponse ropQueryRowsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1193");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1193
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryRowsResponse.RopId.GetType(),
                1193,
                @"[In RopQueryRows ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1195");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1195
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryRows,
                ropQueryRowsResponse.RopId,
                1195,
                @"[In RopQueryRows ROP Failure Response Buffer] RopId (1 byte): For this operation[RopQueryRows], this field[RopId (1 byte)] is set to 0x15.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1196");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1196
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryRowsResponse.InputHandleIndex.GetType(),
                1196,
                @"[In RopQueryRows ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1197");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1197
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryRowsResponse.InputHandleIndex,
                1197,
                @"[In RopQueryRows ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1198");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1198
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryRowsResponse.ReturnValue.GetType(),
                1198,
                @"[In RopQueryRows ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1200");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1200
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropQueryRowsResponse.ReturnValue,
                1200,
                @"[In RopQueryRows ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopAbort
        /// <summary>
        /// Verify RopAbort Success Response
        /// </summary>
        /// <param name="ropAbortResponse">The response of RopAbort request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopAbort request</param>
        private void VerifyRopAbortSuccessResponse(RopAbortResponse ropAbortResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1210");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1210
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortResponse.RopId.GetType(),
                1210,
                @"[In RopAbort ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1212");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1212
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopAbort,
                ropAbortResponse.RopId,
                1212,
                @"[In RopAbort ROP Success Response Buffer] RopId (1 byte): For this operation[RopAbort], this field[RopId (1 byte)] is set to 0x38.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1213");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1213
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortResponse.InputHandleIndex.GetType(),
                1213,
                @"[In RopAbort ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1214");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1214
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropAbortResponse.InputHandleIndex,
                1214,
                @"[In RopAbort ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1215");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1215
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropAbortResponse.ReturnValue.GetType(),
                1215,
                @"[In RopAbort ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1217");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1217
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropAbortResponse.ReturnValue,
                1217,
                @"[In RopAbort ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1218");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1218
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortResponse.TableStatus.GetType(),
                1218,
                @"[In RopAbort ROP Success Response Buffer] TableStatus (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1219,TableStatus:{0}", ropAbortResponse.TableStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1219
            bool isVerifyR1219 = ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatComplete)) == ((byte)TableStatus.TblstatComplete))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatSorting)) == ((byte)TableStatus.TblstatSorting))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatSortError)) == ((byte)TableStatus.TblstatSortError))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatSettingCols)) == ((byte)TableStatus.TblstatSettingCols))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatSetColError)) == ((byte)TableStatus.TblstatSetColError))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatRestricting)) == ((byte)TableStatus.TblstatRestricting))
                                 || ((ropAbortResponse.TableStatus & ((byte)TableStatus.TblstatRestrictError)) == ((byte)TableStatus.TblstatRestrictError));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1219,
                1219,
                @"[In RopAbort ROP Success Response Buffer] TableStatus (1 byte): The possible values[the value of TBLSTAT_COMPLETE is 0x00, the value of TBLSTAT_SORTING is 0x09, the value of TBLSTAT_SORT_ERROR is 0x0A, the value of TBLSTAT_SETTING_COLS is 0x0B, the value of TBLSTAT_SETCOL_ERROR is 0x0D, the value of TBLSTAT_RESTRICTING is 0x0E and the value of TBLSTAT_RESTRICT_ERROR is 0x0F] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.1.3.");
        }

        /// <summary>
        /// Verify RopAbort Failure Response
        /// </summary>
        /// <param name="ropAbortResponse">The response of RopAbort request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopAbort request</param>
        private void VerifyRopAbortFailureResponse(RopAbortResponse ropAbortResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1221");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1221
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortResponse.RopId.GetType(),
                1221,
                @"[In RopAbort ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1223");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1223
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopAbort,
                ropAbortResponse.RopId,
                1223,
                @"[In RopAbort ROP Failure Response Buffer] RopId (1 byte): For this operation[RopAbort], this field[RopId (1 byte)] is set to 0x38.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1224");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1224
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortResponse.InputHandleIndex.GetType(),
                1224,
                @"[In RopAbort ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1225");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1225
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropAbortResponse.InputHandleIndex,
                1225,
                @"[In RopAbort ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1227");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1227
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropAbortResponse.ReturnValue.GetType(),
                1227,
                @"[In RopAbort ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1229");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1229
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropAbortResponse.ReturnValue,
                1229,
                @"[In RopAbort ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetStatus
        /// <summary>
        /// Verify RopGetStatus Success Response
        /// </summary>
        /// <param name="ropGetStatusResponse">The response of RopGetStatus request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStatus request</param>
        private void VerifyRopGetStatusSuccessResponse(RopGetStatusResponse ropGetStatusResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1239");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1239
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStatusResponse.RopId.GetType(),
                1239,
                @"[In RopGetStatus ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1241");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1241
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStatus,
                ropGetStatusResponse.RopId,
                1241,
                @"[In RopGetStatus ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetStatus], this field[RopId (1 byte)] is set to 0x16.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1242");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1242
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStatusResponse.InputHandleIndex.GetType(),
                1242,
                @"[In RopGetStatus ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1243");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1243
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStatusResponse.InputHandleIndex,
                1243,
                @"[In RopGetStatus ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1244");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1244
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStatusResponse.ReturnValue.GetType(),
                1244,
                @"[In RopGetStatus ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1246");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1246
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetStatusResponse.ReturnValue,
                1246,
                @"[In RopGetStatus ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1247");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1247
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStatusResponse.TableStatus.GetType(),
                1247,
                @"[In RopGetStatus ROP Success Response Buffer] TableStatus (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1248,TableStatus:{0}", ropGetStatusResponse.TableStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1248
            bool isVerifyR1248 = ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatComplete)) == ((byte)TableStatus.TblstatComplete))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatSorting)) == ((byte)TableStatus.TblstatSorting))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatSortError)) == ((byte)TableStatus.TblstatSortError))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatSettingCols)) == ((byte)TableStatus.TblstatSettingCols))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatSetColError)) == ((byte)TableStatus.TblstatSetColError))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatRestricting)) == ((byte)TableStatus.TblstatRestricting))
                                 || ((ropGetStatusResponse.TableStatus & ((byte)TableStatus.TblstatRestrictError)) == ((byte)TableStatus.TblstatRestrictError));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1248,
                1248,
                @"[In RopGetStatus ROP Success Response Buffer] TableStatus (1 byte): The possible values[the value of TBLSTAT_COMPLETE is 0x00, the value of TBLSTAT_SORTING is 0x09, the value of TBLSTAT_SORT_ERROR is 0x0A, the value of TBLSTAT_SETTING_COLS is 0x0B, the value of TBLSTAT_SETCOL_ERROR is 0x0D, the value of TBLSTAT_RESTRICTING is 0x0E and the value of TBLSTAT_RESTRICT_ERROR is 0x0F] for this enumeration are specified in [MS-OXCTABL] section 2.2.2.1.3.");
        }

        /// <summary>
        /// Verify RopGetStatus Failure Response
        /// </summary>
        /// <param name="ropGetStatusResponse">The response of RopGetStatus request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStatus request</param>
        private void VerifyRopGetStatusFailureResponse(RopGetStatusResponse ropGetStatusResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1250");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1250
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStatusResponse.RopId.GetType(),
                1250,
                @"[In RopGetStatus ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1252");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1252
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStatus,
                ropGetStatusResponse.RopId,
                1252,
                @"[In RopGetStatus ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetStatus], this field is set to 0x16.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1253");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1253
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStatusResponse.InputHandleIndex.GetType(),
                1253,
                @"[In RopGetStatus ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1254");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1254
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStatusResponse.InputHandleIndex,
                1254,
                @"[In RopGetStatus ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1255");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1255
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStatusResponse.ReturnValue.GetType(),
                1255,
                @"[In RopGetStatus ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1257");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1257
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetStatusResponse.ReturnValue,
                1257,
                @"[In RopGetStatus ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopQueryPosition
        /// <summary>
        /// Verify RopQueryPosition Success Response
        /// </summary>
        /// <param name="ropQueryPositionResponse">The response of RopQueryPosition request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryPosition request</param>
        private void VerifyRopQueryPositionSuccessResponse(RopQueryPositionResponse ropQueryPositionResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1267");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1267
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryPositionResponse.RopId.GetType(),
                1267,
                @"[In RopQueryPosition ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1269");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1269
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryPosition,
                ropQueryPositionResponse.RopId,
                1269,
                @"[In RopQueryPosition ROP Success Response Buffer] RopId (1 byte): For this operation[RopQueryPosition], this field[RopId (1 byte)] is set to 0x17.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1270");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1270
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryPositionResponse.InputHandleIndex.GetType(),
                1270,
                @"[In RopQueryPosition ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1271");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1271
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryPositionResponse.InputHandleIndex,
                1271,
                @"[In RopQueryPosition ROP Success Response Buffer] InputHandleIndex (1 byte): This index[InputHandleIndex] MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1272");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1272
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryPositionResponse.ReturnValue.GetType(),
                1272,
                @"[In RopQueryPosition ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1274");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1274
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropQueryPositionResponse.ReturnValue,
                1274,
                @"[In RopQueryPosition ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1275");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1275
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryPositionResponse.Numerator.GetType(),
                1275,
                @"[In RopQueryPosition ROP Success Response Buffer] Numerator (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1277");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1277
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryPositionResponse.Denominator.GetType(),
                1277,
                @"[In RopQueryPosition ROP Success Response Buffer] Denominator (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopQueryPosition Failure Response
        /// </summary>
        /// <param name="ropQueryPositionResponse">The response of RopQueryPosition request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryPosition request</param>
        private void VerifyRopQueryPositionFailureResponse(RopQueryPositionResponse ropQueryPositionResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1279");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1279
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryPositionResponse.RopId.GetType(),
                1279,
                @"[In RopQueryPosition ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1281");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1281
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryPosition,
                ropQueryPositionResponse.RopId,
                1281,
                @"[In RopQueryPosition ROP Failure Response Buffer] RopId (1 byte): For this operation[RopQueryPosition], this field[RopId (1 byte)] is set to 0x17.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1282");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1282
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryPositionResponse.InputHandleIndex.GetType(),
                1282,
                @"[In RopQueryPosition ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1283");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1283
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryPositionResponse.InputHandleIndex,
                1283,
                @"[In RopQueryPosition ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1284");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1284
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryPositionResponse.ReturnValue.GetType(),
                1284,
                @"[In RopQueryPosition ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1286");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1286
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropQueryPositionResponse.ReturnValue,
                1286,
                @"[In RopQueryPosition ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSeekRow
        /// <summary>
        /// Verify RopSeekRow Success Response
        /// </summary>
        /// <param name="ropSeekRowResponse">The response of RopSeekRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekRow request</param>
        private void VerifyRopSeekRowSuccessResponse(RopSeekRowResponse ropSeekRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1303");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1303
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowResponse.RopId.GetType(),
                1303,
                @"[In RopSeekRow ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1305");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1305
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekRow,
                ropSeekRowResponse.RopId,
                1305,
                @"[In RopSeekRow ROP Success Response Buffer] RopId (1 byte): For this operation[RopSeekRow], this field[RopId (1 byte)] is set to 0x18.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1306");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1306
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowResponse.InputHandleIndex.GetType(),
                1306,
                @"[In RopSeekRow ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1307");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1307
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekRowResponse.InputHandleIndex,
                1307,
                @"[In RopSeekRow ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1308");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1308
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowResponse.ReturnValue.GetType(),
                1308,
                @"[In RopSeekRow ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1310");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1310
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSeekRowResponse.ReturnValue,
                1310,
                @"[In RopSeekRow ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1311,HasSoughtLess:{0}", ropSeekRowResponse.HasSoughtLess);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1311
            // If HasSoughtLess is byte type and can be boolean type representation, then verify R1311
            bool isVerifyR1311 = (typeof(byte) == ropSeekRowResponse.HasSoughtLess.GetType()) &&
                                 ((Convert.ToBoolean(ropSeekRowResponse.HasSoughtLess) == true) ||
                                 (Convert.ToBoolean(ropSeekRowResponse.HasSoughtLess) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1311,
                1311,
                @"[In RopSeekRow ROP Success Response Buffer] HasSoughtLess (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1313");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1313
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(int),
                ropSeekRowResponse.RowsSought.GetType(),
                1313,
                @"[In RopSeekRow ROP Success Response Buffer] RowsSought (4 bytes): A signed integer.");
        }

        /// <summary>
        /// Verify RopSeekRow Failure Response
        /// </summary>
        /// <param name="ropSeekRowResponse">The response of RopSeekRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekRow request</param>
        private void VerifyRopSeekRowFailureResponse(RopSeekRowResponse ropSeekRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1315");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1315
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowResponse.RopId.GetType(),
                1315,
                @"[In RopSeekRow ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1317");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1317
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekRow,
                ropSeekRowResponse.RopId,
                1317,
                @"[In RopSeekRow ROP Failure Response Buffer] RopId (1 byte): For this operation[RopSeekRow], this field[RopId (1 byte)] is set to 0x18.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1318");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1318
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowResponse.InputHandleIndex.GetType(),
                1318,
                @"[In RopSeekRow ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1319");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1319
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekRowResponse.InputHandleIndex,
                1319,
                @"[In RopSeekRow ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1320");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1320
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowResponse.ReturnValue.GetType(),
                1320,
                @"[In RopSeekRow ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1322");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1322
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSeekRowResponse.ReturnValue,
                1322,
                @"[In RopSeekRow ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSeekRowBookmark
        /// <summary>
        /// Verify RopSeekRowBookmark Success Response
        /// </summary>
        /// <param name="ropSeekRowBookmarkResponse">The response of RopSeekRowBookmark request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekRowBookmark request</param>
        private void VerifyRopSeekRowBookmarkSuccessResponse(RopSeekRowBookmarkResponse ropSeekRowBookmarkResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1341");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1341
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowBookmarkResponse.RopId.GetType(),
                1341,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1343");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1343
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekRowBookmark,
                ropSeekRowBookmarkResponse.RopId,
                1343,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] RopId (1 byte): For this operation[RopSeekRowBookmark], this field[RopId (1 byte)] is set to 0x19.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1344");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1344
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowBookmarkResponse.InputHandleIndex.GetType(),
                1344,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1345");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1345
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekRowBookmarkResponse.InputHandleIndex,
                1345,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1346");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1346
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowBookmarkResponse.ReturnValue.GetType(),
                1346,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1348");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1348
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSeekRowBookmarkResponse.ReturnValue,
                1348,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1349,RowNoLongerVisible:{0}", ropSeekRowBookmarkResponse.RowNoLongerVisible);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1349
            // If RowNoLongerVisible is byte type and can be boolean type representation, then verify R1349
            bool isVerifyR1349 = (typeof(byte) == ropSeekRowBookmarkResponse.RowNoLongerVisible.GetType()) &&
                                 ((Convert.ToBoolean(ropSeekRowBookmarkResponse.RowNoLongerVisible) == true) ||
                                 (Convert.ToBoolean(ropSeekRowBookmarkResponse.RowNoLongerVisible) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1349,
                1349,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] RowNoLongerVisible (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1351,HasSoughtLess:{0}", ropSeekRowBookmarkResponse.HasSoughtLess);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1351
            bool isVerifyR1351 = (typeof(byte) == ropSeekRowBookmarkResponse.HasSoughtLess.GetType()) &&
                                 ((Convert.ToBoolean(ropSeekRowBookmarkResponse.HasSoughtLess) == true) ||
                                 (Convert.ToBoolean(ropSeekRowBookmarkResponse.HasSoughtLess) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1351,
                1351,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] HasSoughtLess (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1353");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1353
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowBookmarkResponse.RowsSought.GetType(),
                1353,
                @"[In RopSeekRowBookmark ROP Success Response Buffer] RowsSought (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopSeekRowBookmark Failure Response
        /// </summary>
        /// <param name="ropSeekRowBookmarkResponse">The response of RopSeekRowBookmark request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekRowBookmark request</param>
        private void VerifyRopSeekRowBookmarkFailureResponse(RopSeekRowBookmarkResponse ropSeekRowBookmarkResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1355");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1355
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowBookmarkResponse.RopId.GetType(),
                1355,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1357");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1357
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekRowBookmark,
                ropSeekRowBookmarkResponse.RopId,
                1357,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] RopId (1 byte): For this operation[RopSeekRowBookmark], this field[RopId (1 byte)] is set to 0x19.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1358");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1358
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowBookmarkResponse.InputHandleIndex.GetType(),
                1358,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1359");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1359
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekRowBookmarkResponse.InputHandleIndex,
                1359,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1360");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1360
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowBookmarkResponse.ReturnValue.GetType(),
                1360,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1362");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1362
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSeekRowBookmarkResponse.ReturnValue,
                1362,
                @"[In RopSeekRowBookmark ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSeekRowFractional
        /// <summary>
        /// Verify RopSeekRowFractional Response
        /// </summary>
        /// <param name="ropSeekRowFractionalResponse">The response of RopSeekRowFractional request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekRowFractional request</param>
        private void VerifyRopSeekRowFractionalResponse(RopSeekRowFractionalResponse ropSeekRowFractionalResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1376");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1376
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowFractionalResponse.RopId.GetType(),
                1376,
                @"[In RopSeekRowFractional ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1378");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1378
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekRowFractional,
                ropSeekRowFractionalResponse.RopId,
                1378,
                @"[In RopSeekRowFractional ROP Response Buffer] RopId (1 byte): For this operation[RopSeekRowFractional], this field[RopId (1 byte)] is set to 0x1A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1379");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1379
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekRowFractionalResponse.InputHandleIndex.GetType(),
                1379,
                @"[In RopSeekRowFractional ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1380");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1380
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekRowFractionalResponse.InputHandleIndex,
                1380,
                @"[In RopSeekRowFractional ROP Response Buffer] InputHandleIndex (1 byte): This index[InputHandleIndex] MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1381");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1381
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekRowFractionalResponse.ReturnValue.GetType(),
                1381,
                @"[In RopSeekRowFractional ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopCreateBookmark
        /// <summary>
        /// Verify RopCreateBookmark Success Response
        /// </summary>
        /// <param name="ropCreateBookmarkResponse">The response of RopCreateBookmark request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopCreateBookmark request</param>
        private void VerifyRopCreateBookmarkSuccessResponse(RopCreateBookmarkResponse ropCreateBookmarkResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1392");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1392
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateBookmarkResponse.RopId.GetType(),
                1392,
                @"[In RopCreateBookmark ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1394");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1394
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateBookmark,
                ropCreateBookmarkResponse.RopId,
                1394,
                @"[In RopCreateBookmark ROP Success Response Buffer] RopId (1 byte): For this operation[RopCreateBookmark], this field[RopId (1 byte)] is set to 0x1B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1395");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1395
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateBookmarkResponse.InputHandleIndex.GetType(),
                1395,
                @"[In RopCreateBookmark ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1396");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1396
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropCreateBookmarkResponse.InputHandleIndex,
                1396,
                @"[In RopCreateBookmark ROP Success Response Buffer] InputHandleIndex (1 byte): This index[InputHandleIndex] MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1397");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1397
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateBookmarkResponse.ReturnValue.GetType(),
                1397,
                @"[In RopCreateBookmark ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1399");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1399
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCreateBookmarkResponse.ReturnValue,
                1399,
                @"[In RopCreateBookmark ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1400");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1400
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropCreateBookmarkResponse.BookmarkSize.GetType(),
                1400,
                @"[In RopCreateBookmark ROP Success Response Buffer] BookmarkSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1402");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1402
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropCreateBookmarkResponse.Bookmark.GetType(),
                1402,
                @"[In RopCreateBookmark ROP Success Response Buffer] Bookmark (variable): An array of bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1403");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1403
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropCreateBookmarkResponse.BookmarkSize,
                (ushort)ropCreateBookmarkResponse.Bookmark.Length,
                1403,
                @"[In RopCreateBookmark ROP Success Response Buffer] Bookmark (variable): The size of this field, in bytes, is specified by the BookmarkSize field.");
        }

        /// <summary>
        /// Verify RopCreateBookmark Failure Response
        /// </summary>
        /// <param name="ropCreateBookmarkResponse">The response of RopCreateBookmark request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopCreateBookmark request</param>
        private void VerifyRopCreateBookmarkFailureResponse(RopCreateBookmarkResponse ropCreateBookmarkResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1405");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1405
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateBookmarkResponse.RopId.GetType(),
                1405,
                @"[In RopCreateBookmark ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1407");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1407
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateBookmark,
                ropCreateBookmarkResponse.RopId,
                1407,
                @"[In RopCreateBookmark ROP Failure Response Buffer] RopId (1 byte): For this operation[RopCreateBookmark], this field[RopId (1 byte)] is set to 0x1B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1408");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1408
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateBookmarkResponse.InputHandleIndex.GetType(),
                1408,
                @"[In RopCreateBookmark ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1409");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1409
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropCreateBookmarkResponse.InputHandleIndex,
                1409,
                @"[In RopCreateBookmark ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1410");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1410
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateBookmarkResponse.ReturnValue.GetType(),
                1410,
                @"[In RopCreateBookmark ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1412");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1412
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropCreateBookmarkResponse.ReturnValue,
                1412,
                @"[In RopCreateBookmark ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopQueryColumnsAll
        /// <summary>
        /// Verify RopQueryColumnsAll Success Response
        /// </summary>
        /// <param name="ropQueryColumnsAllResponse">The response of RopQueryColumnsAll request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryColumnsAll request</param>
        private void VerifyRopQueryColumnsAllSuccessResponse(RopQueryColumnsAllResponse ropQueryColumnsAllResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1423");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1423
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryColumnsAllResponse.RopId.GetType(),
                1423,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1425");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1425
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryColumnsAll,
                ropQueryColumnsAllResponse.RopId,
                1425,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] RopId (1 byte): For this operation[RopQueryColumnsAll], this field[RopId (1 byte)] is set to 0x37.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1426");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1426
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryColumnsAllResponse.InputHandleIndex.GetType(),
                1426,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1427");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1427
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryColumnsAllResponse.InputHandleIndex,
                1427,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1428");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1428
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryColumnsAllResponse.ReturnValue.GetType(),
                1428,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1430");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1430
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropQueryColumnsAllResponse.ReturnValue,
                1430,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1431");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1431
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropQueryColumnsAllResponse.PropertyTagCount.GetType(),
                1431,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] PropertyTagCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1433");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1433
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                ropQueryColumnsAllResponse.PropertyTags.GetType(),
                1433,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] PropertyTags (variable): An array of PropertyTag structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1434");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1434
            Site.CaptureRequirementIfAreEqual<int>(
                ropQueryColumnsAllResponse.PropertyTags.Length,
                ropQueryColumnsAllResponse.PropertyTagCount,
                1434,
                @"[In RopQueryColumnsAll ROP Success Response Buffer] PropertyTags (variable): The number of structures contained in this field is specified by the PropertyTagCount field.");
        }

        /// <summary>
        /// Verify RopQueryColumnsAll Failure Response
        /// </summary>
        /// <param name="ropQueryColumnsAllResponse">The response of RopQueryColumnsAll request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryColumnsAll request</param>
        private void VerifyRopQueryColumnsAllFailureResponse(RopQueryColumnsAllResponse ropQueryColumnsAllResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1437");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1437
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryColumnsAllResponse.RopId.GetType(),
                1437,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1439");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1439
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryColumnsAll,
                ropQueryColumnsAllResponse.RopId,
                1439,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] RopId (1 byte): For this operation[RopQueryColumnsAll], this field[RopId (1 byte)] is set to 0x37.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1440");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1440
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryColumnsAllResponse.InputHandleIndex.GetType(),
                1440,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1441");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1441
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryColumnsAllResponse.InputHandleIndex,
                1441,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1442");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1442
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryColumnsAllResponse.ReturnValue.GetType(),
                1442,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1444");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1444
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropQueryColumnsAllResponse.ReturnValue,
                1444,
                @"[In RopQueryColumnsAll ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopFindRow
        /// <summary>
        /// Verify RopFindRow Success Response
        /// </summary>
        /// <param name="ropFindRowResponse">The response of RopFindRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopFindRow request</param>
        private void VerifyRopFindRowSuccessResponse(RopFindRowResponse ropFindRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1471");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1471
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFindRowResponse.RopId.GetType(),
                1471,
                @"[In RopFindRow ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1473");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1473
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFindRow,
                ropFindRowResponse.RopId,
                1473,
                @"[In RopFindRow ROP Success Response Buffer] RopId (1 byte): For this operation[RopFindRow], this field[RopId (1 byte)] is set to 0x4F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1474");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1474
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFindRowResponse.InputHandleIndex.GetType(),
                1474,
                @"[In RopFindRow ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1475");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1475
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropFindRowResponse.InputHandleIndex,
                1475,
                @"[In RopFindRow ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the InputHandleIndex specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1476");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1476
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFindRowResponse.ReturnValue.GetType(),
                1476,
                @"[In RopFindRow ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1478");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1478
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropFindRowResponse.ReturnValue,
                1478,
                @"[In RopFindRow ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1479,RowNoLongerVisible:{0}", ropFindRowResponse.RowNoLongerVisible);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1479
            // If RowNoLongerVisible is byte type and can be boolean type representation, then verify R1479
            bool isVerifyR1479 = (typeof(byte) == ropFindRowResponse.RowNoLongerVisible.GetType()) &&
                                 ((Convert.ToBoolean(ropFindRowResponse.RowNoLongerVisible) == true) ||
                                 (Convert.ToBoolean(ropFindRowResponse.RowNoLongerVisible) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1479,
                1479,
                @"[In RopFindRow ROP Success Response Buffer] RowNoLongerVisible (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1481,HasRowData:{0}", ropFindRowResponse.HasRowData);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1481
            // If HasRowData is byte type and can be boolean type representation, then verify R1481
            bool isVerifyR1481 = (typeof(byte) == ropFindRowResponse.HasRowData.GetType()) &&
                                 ((Convert.ToBoolean(ropFindRowResponse.HasRowData) == true) ||
                                 (Convert.ToBoolean(ropFindRowResponse.HasRowData) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1481,
                1481,
                @"[In RopFindRow ROP Success Response Buffer] HasRowData (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1483");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1483
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyRow),
                ropFindRowResponse.RowData.GetType(),
                1483,
                @"[In RopFindRow ROP Success Response Buffer] RowData (variable): A PropertyRow structure");
        }

        /// <summary>
        /// Verify RopFindRow Failure Response
        /// </summary>
        /// <param name="ropFindRowResponse">The response of RopFindRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopFindRow request</param>
        private void VerifyRopFindRowFailureResponse(RopFindRowResponse ropFindRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1487");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1487
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFindRowResponse.RopId.GetType(),
                1487,
                @"[In RopFindRow ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1489");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1489
            Site.CaptureRequirementIfAreEqual<byte>(
               (byte)RopId.RopFindRow,
                ropFindRowResponse.RopId,
                1489,
                @"[In RopFindRow ROP Failure Response Buffer] RopId (1 byte): For this operation[RopFindRow], this field[RopId (1 byte)] is set to 0x4F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1490");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1490
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFindRowResponse.InputHandleIndex.GetType(),
                1490,
                @"[In RopFindRow ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1491");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1491
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropFindRowResponse.InputHandleIndex,
                1491,
                @"[In RopFindRow ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1492");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1492
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFindRowResponse.ReturnValue.GetType(),
                1492,
                @"[In RopFindRow ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1494");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1494
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropFindRowResponse.ReturnValue,
                1494,
                @"[In RopFindRow ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopFreeBookmark
        /// <summary>
        /// Verify RopFreeBookmark Response
        /// </summary>
        /// <param name="ropFreeBookmarkResponse">The response of RopFreeBookmark request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopFreeBookmark request</param>
        private void VerifyRopFreeBookmarkResponse(RopFreeBookmarkResponse ropFreeBookmarkResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1509");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1509
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFreeBookmarkResponse.RopId.GetType(),
                1509,
                @"[In RopFreeBookmark ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1511");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1511
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFreeBookmark,
                ropFreeBookmarkResponse.RopId,
                1511,
                @"[In RopFreeBookmark ROP Response Buffer] RopId (1 byte): For this operation[RopFreeBookmark], this field[RopId (1 byte)] is set to 0x89.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1512");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1512
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFreeBookmarkResponse.InputHandleIndex.GetType(),
                1512,
                @"[In RopFreeBookmark ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1513");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1513
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropFreeBookmarkResponse.InputHandleIndex,
                1513,
                @"[In RopFreeBookmark ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1514");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1514
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFreeBookmarkResponse.ReturnValue.GetType(),
                1514,
                @"[In RopFreeBookmark ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopResetTable
        /// <summary>
        /// Verify RopResetTable Response
        /// </summary>
        /// <param name="ropResetTableResponse">The response of RopResetTable request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopResetTable request</param>
        private void VerifyRopResetTableResponse(RopResetTableResponse ropResetTableResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1525");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1525
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropResetTableResponse.RopId.GetType(),
                1525,
                @"[In RopResetTable ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1527");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1527
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopResetTable,
                ropResetTableResponse.RopId,
                1527,
                @"[In RopResetTable ROP Response Buffer] RopId (1 byte): For this operation[RopResetTable], this field is set to 0x81.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1528");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1528
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropResetTableResponse.InputHandleIndex.GetType(),
                1528,
                @"[In RopResetTable ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1529");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1529
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropResetTableResponse.InputHandleIndex,
                1529,
                @"[In RopResetTable ROP Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1530");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1530
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropResetTableResponse.ReturnValue.GetType(),
                1530,
                @"[In RopResetTable ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopExpandRow
        /// <summary>
        /// Verify RopExpandRow Success Response
        /// </summary>
        /// <param name="ropExpandRowResponse">The response of RopExpandRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopExpandRow request</param>
        /// <param name="propertyTags">Array of PropertyTag structures specifies the property values that are visible in table rows, set by a RopSetColumns request</param>
        private void VerifyRopExpandRowSuccessResponse(RopExpandRowResponse ropExpandRowResponse, byte inputHandleIndex, PropertyTag[] propertyTags)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1545");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1545
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropExpandRowResponse.RopId.GetType(),
                1545,
                @"[In RopExpandRow ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1547");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1547
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopExpandRow,
                ropExpandRowResponse.RopId,
                1547,
                @"[In RopExpandRow ROP Success Response Buffer] RopId (1 byte): For this operation[RopExpandRow], this field[RopId (1 byte)] is set to 0x59.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1548");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1548
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropExpandRowResponse.InputHandleIndex.GetType(),
                1548,
                @"[In RopExpandRow ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1549");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1549
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropExpandRowResponse.InputHandleIndex,
                1549,
                @"[In RopExpandRow ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1550");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1550
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropExpandRowResponse.ReturnValue.GetType(),
                1550,
                @"[In RopExpandRow ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1552");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1552
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropExpandRowResponse.ReturnValue,
                1552,
                @"[In RopExpandRow ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1553");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1553
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropExpandRowResponse.ExpandedRowCount.GetType(),
                1553,
                @"[In RopExpandRow ROP Success Response Buffer] ExpandedRowCount (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1555");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1555
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropExpandRowResponse.RowCount.GetType(),
                1555,
                @"[In RopExpandRow ROP Success Response Buffer] RowCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1558");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1558
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropExpandRowResponse.RowCount,
                (ushort)ropExpandRowResponse.RowData.Count,
                1558,
                @"[In RopExpandRow ROP Success Response Buffer] RowData (variable): The number of structures contained in this field is specified by the RowCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1560,the actual value of the RowData is {0}", ropExpandRowResponse.RowData);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1560
            // Assume the result is true. If one of rows element's count not equal columns count, then false.
            bool isVerifyR1560 = true;
            foreach (PropertyRow propertyRow in ropExpandRowResponse.RowData.PropertyRows)
            {
                isVerifyR1560 = isVerifyR1560 && (propertyRow.PropertyValues.Count == propertyTags.Length);

                // If one row element's count not equal columns count then break the iteration
                if (!isVerifyR1560)
                {
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1560,
                1560,
                @"[In RopExpandRow ROP Success Response Buffer] RowData (variable): The columns used for these rows were those previously set on this table by a RopSetColumns ROP request (section 2.2.5.1).");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1557");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1557
            Site.CaptureRequirementIfAreEqual<Type>(
                ropExpandRowResponse.RowData.PropertyRows.GetType(),
                typeof(List<PropertyRow>),
                1557,
                @"[In RopExpandRow ROP Success Response Buffer] RowData (variable): A list of PropertyRow structures.");

            foreach (PropertyRow propertyRow in ropExpandRowResponse.RowData.PropertyRows)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R76");

                // Verify MS-OXCROPS requirement: MS-OXCDATA_R76
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    propertyRow.Flag.GetType(),
                    "MS-OXCDATA",
                    76,
                    @"[In FlaggedPropertyRow Structure] Flag (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R80");

                // Verify MS-OXCROPS requirement: MS-OXCDATA_R80
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(List<PropertyValue>),
                    propertyRow.PropertyValues.GetType(),
                    "MS-OXCDATA",
                    80,
                    @"[In FlaggedPropertyRow Structure] ValueArray (variable):  An array of variable-sized structures.");
            }
        }

        /// <summary>
        /// Verify RopExpandRow Failure Response
        /// </summary>
        /// <param name="ropExpandRowResponse">The response of RopExpandRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopExpandRow request</param>
        private void VerifyRopExpandRowFailureResponse(RopExpandRowResponse ropExpandRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1561");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1561
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropExpandRowResponse.RopId.GetType(),
                1561,
                @"[In RopExpandRow ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1563");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1563
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopExpandRow,
                ropExpandRowResponse.RopId,
                1563,
                @"[In RopExpandRow ROP Failure Response Buffer] RopId (1 byte): For this operation[RopExpandRow], this field[RopId (1 byte)] is set to 0x59.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1564");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1564
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropExpandRowResponse.InputHandleIndex.GetType(),
                1564,
                @"[In RopExpandRow ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1565");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1565
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropExpandRowResponse.InputHandleIndex,
                1565,
                @"[In RopExpandRow ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1566");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1566
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropExpandRowResponse.ReturnValue.GetType(),
                1566,
                @"[In RopExpandRow ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1568");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1568
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropExpandRowResponse.ReturnValue,
                1568,
                @"[In RopExpandRow ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopCollapseRow
        /// <summary>
        /// Verify RopCollapseRow Success Response
        /// </summary>
        /// <param name="ropCollapseRowResponse">The response of RopCollapseRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopCollapseRow request</param>
        private void VerifyRopCollapseRowSuccessResponse(RopCollapseRowResponse ropCollapseRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1580");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1580
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCollapseRowResponse.RopId.GetType(),
                1580,
                @"[In RopCollapseRow ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1582");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1582
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCollapseRow,
                ropCollapseRowResponse.RopId,
                1582,
                @"[In RopCollapseRow ROP Success Response Buffer] RopId (1 byte): For this operation[RopCollapseRow], this field[RopId (1 byte)] is set to 0x5A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1583");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1583
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCollapseRowResponse.InputHandleIndex.GetType(),
                1583,
                @"[In RopCollapseRow ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1584");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1584
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropCollapseRowResponse.InputHandleIndex,
                1584,
                @"[In RopCollapseRow ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1585");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1585
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCollapseRowResponse.ReturnValue.GetType(),
                1585,
                @"[In RopCollapseRow ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1587");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1587
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCollapseRowResponse.ReturnValue,
                1587,
                @"[In RopCollapseRow ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1588");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1588
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCollapseRowResponse.CollapsedRowCount.GetType(),
                1588,
                @"[In RopCollapseRow ROP Success Response Buffer] CollapsedRowCount (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopCollapseRow Failure Response
        /// </summary>
        /// <param name="ropCollapseRowResponse">The response of RopCollapseRow request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopCollapseRow request</param>
        private void VerifyRopCollapseRowFailureResponse(RopCollapseRowResponse ropCollapseRowResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1590");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1590
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCollapseRowResponse.RopId.GetType(),
                1590,
                @"[In RopCollapseRow ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1592");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1592
            Site.CaptureRequirementIfAreEqual<byte>(
               (byte)RopId.RopCollapseRow,
                ropCollapseRowResponse.RopId,
                1592,
                @"[In RopCollapseRow ROP Failure Response Buffer] RopId (1 byte): For this operation[RopCollapseRow], this field is set to 0x5A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1593");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1593
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCollapseRowResponse.InputHandleIndex.GetType(),
                1593,
                @"[In RopCollapseRow ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1594");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1594
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropCollapseRowResponse.InputHandleIndex,
                1594,
                @"[In RopCollapseRow ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1595");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1595
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCollapseRowResponse.ReturnValue.GetType(),
                1595,
                @"[In RopCollapseRow ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1597");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1597
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropCollapseRowResponse.ReturnValue,
                1597,
                @"[In RopCollapseRow ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetCollapseState
        /// <summary>
        /// Verify RopGetCollapseState Success Response
        /// </summary>
        /// <param name="ropGetCollapseStateResponse">The response of RopGetCollapseState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetCollapseState request</param>
        private void VerifyRopGetCollapseStateSuccessResponse(RopGetCollapseStateResponse ropGetCollapseStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1612");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1612
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetCollapseStateResponse.RopId.GetType(),
                1612,
                @"[In RopGetCollapseState ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1614");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1614
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetCollapseState,
                ropGetCollapseStateResponse.RopId,
                1614,
                @"[In RopGetCollapseState ROP Success Response Buffer] RopId (1 byte): For this operation[RopGetCollapseState], this field[RopId (1 byte)] is set to 0x6B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1615");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1615
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetCollapseStateResponse.InputHandleIndex.GetType(),
                1615,
                @"[In RopGetCollapseState ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1616");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1616
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetCollapseStateResponse.InputHandleIndex,
                1616,
                @"[In RopGetCollapseState ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1617");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1617
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetCollapseStateResponse.ReturnValue.GetType(),
                1617,
                @"[In RopGetCollapseState ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1619");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1619
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetCollapseStateResponse.ReturnValue,
                1619,
                @"[In RopGetCollapseState ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1620");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1620
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetCollapseStateResponse.CollapseStateSize.GetType(),
                1620,
                @"[In RopGetCollapseState ROP Success Response Buffer] CollapseStateSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1622");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1622
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropGetCollapseStateResponse.CollapseState.GetType(),
                1622,
                @"[In RopGetCollapseState ROP Success Response Buffer] CollapseState (variable): An array of bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1623");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1623
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropGetCollapseStateResponse.CollapseStateSize,
                (ushort)ropGetCollapseStateResponse.CollapseState.Length,
                1623,
                @"[In RopGetCollapseState ROP Success Response Buffer] CollapseState (variable): The size of this field, in bytes, is specified by the CollapseStateSize field.");
        }

        /// <summary>
        /// Verify RopGetCollapseState Failure Response
        /// </summary>
        /// <param name="ropGetCollapseStateResponse">The response of RopGetCollapseState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetCollapseState request</param>
        private void VerifyRopGetCollapseStateFailureResponse(RopGetCollapseStateResponse ropGetCollapseStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1625");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1625
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetCollapseStateResponse.RopId.GetType(),
                1625,
                @"[In RopGetCollapseState ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1627");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1627
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetCollapseState,
                ropGetCollapseStateResponse.RopId,
                1627,
                @"[In RopGetCollapseState ROP Failure Response Buffer] RopId (1 byte): For this operation[RopGetCollapseState], this field[RopId (1 byte)] is set to 0x6B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1628");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1628
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetCollapseStateResponse.InputHandleIndex.GetType(),
                1628,
                @"[In RopGetCollapseState ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1629");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1629
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetCollapseStateResponse.InputHandleIndex,
                1629,
                @"[In RopGetCollapseState ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1630");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1630
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetCollapseStateResponse.ReturnValue.GetType(),
                1630,
                @"[In RopGetCollapseState ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1632");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1632
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetCollapseStateResponse.ReturnValue,
                1632,
                @"[In RopGetCollapseState ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSetCollapseState
        /// <summary>
        /// Verify RopSetCollapseState Success Response
        /// </summary>
        /// <param name="ropSetCollapseStateResponse">The response of RopSetCollapseState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetCollapseState request</param>
        private void VerifyRopSetCollapseStateSuccessResponse(RopSetCollapseStateResponse ropSetCollapseStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1647");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1647
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetCollapseStateResponse.RopId.GetType(),
                1647,
                @"[In RopSetCollapseState ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1649");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1649
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetCollapseState,
                ropSetCollapseStateResponse.RopId,
                1649,
                @"[In RopSetCollapseState ROP Success Response Buffer] RopId (1 byte): For this operation[RopSetCollapseState], this field[RopId (1 byte)] is set to 0x6C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1650");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1650
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetCollapseStateResponse.InputHandleIndex.GetType(),
                1650,
                @"[In RopSetCollapseState ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1651");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1651
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetCollapseStateResponse.InputHandleIndex,
                1651,
                @"[In RopSetCollapseState ROP Success Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1652");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1652
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetCollapseStateResponse.ReturnValue.GetType(),
                1652,
                @"[In RopSetCollapseState ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1654");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1654
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetCollapseStateResponse.ReturnValue,
                1654,
                @"[In RopSetCollapseState ROP Success Response Buffer] ReturnValue (4 bytes): For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1655");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1655
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropSetCollapseStateResponse.BookmarkSize.GetType(),
                1655,
                @"[In RopSetCollapseState ROP Success Response Buffer] BookmarkSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1657");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1657
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropSetCollapseStateResponse.Bookmark.GetType(),
                1657,
                @"[In RopSetCollapseState ROP Success Response Buffer] Bookmark (variable): An array of bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1658");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1658
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropSetCollapseStateResponse.BookmarkSize,
                (ushort)ropSetCollapseStateResponse.Bookmark.Length,
                1658,
                @"[In RopSetCollapseState ROP Success Response Buffer] Bookmark (variable): The size of this field, in bytes, is specified by the BookmarkSize field.");
        }

        /// <summary>
        /// Verify RopSetCollapseState Failure Response
        /// </summary>
        /// <param name="ropSetCollapseStateResponse">The response of RopSetCollapseState request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetCollapseState request</param>
        private void VerifyRopSetCollapseStateFailureResponse(RopSetCollapseStateResponse ropSetCollapseStateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1660");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1660
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetCollapseStateResponse.RopId.GetType(),
                1660,
                @"[In RopSetCollapseState ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1662");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1662
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetCollapseState,
                ropSetCollapseStateResponse.RopId,
                1662,
                @"[In RopSetCollapseState ROP Failure Response Buffer] RopId (1 byte): For this operation[RopSetCollapseState], this field[RopId (1 byte)] is set to 0x6C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1663");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1663
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetCollapseStateResponse.InputHandleIndex.GetType(),
                1663,
                @"[In RopSetCollapseState ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1664");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1664
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetCollapseStateResponse.InputHandleIndex,
                1664,
                @"[In RopSetCollapseState ROP Failure Response Buffer] InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1665");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1665
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetCollapseStateResponse.ReturnValue.GetType(),
                1665,
                @"[In RopSetCollapseState ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1667");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1667
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetCollapseStateResponse.ReturnValue,
                1667,
                @"[In RopSetCollapseState ROP Failure Response Buffer] ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #endregion

        #region Message ROPs

        #region RopOpenMessage
        /// <summary>
        /// Verify RopOpenMessage Success Response
        /// </summary>
        /// <param name="ropOpenMessageResponse">The response of RopOpenMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenMessage request</param>
        private void VerifyRopOpenMessageSuccessResponse(RopOpenMessageResponse ropOpenMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1692");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1692
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenMessageResponse.RopId.GetType(),
                1692,
                @"[In RopOpenMessage ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1694");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1694
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenMessage,
                ropOpenMessageResponse.RopId,
                1694,
                @"[In RopOpenMessage ROP Success Response Buffer,RopId (1 byte)]For this operation[RopOpenMessage], this field[RopId (1 byte)] is set to 0x03.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1695");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1695
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenMessageResponse.OutputHandleIndex.GetType(),
                1695,
                @"[In RopOpenMessage ROP Success Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1696");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1696
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenMessageResponse.OutputHandleIndex,
                1696,
                @"[In RopOpenMessage ROP Success Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1698");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1698
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenMessageResponse.ReturnValue.GetType(),
                1698,
                @"[In RopOpenMessage ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1702");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1702
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropOpenMessageResponse.ReturnValue,
                1702,
                @"[In RopOpenMessage ROP Success Response Buffer,ReturnValue (4 bytes)]For this[Success Response Buffer] response, this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1703,HasNamedProperties:{0}", ropOpenMessageResponse.HasNamedProperties);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1703
            // If HasNamedProperties is byte type and can be boolean type representation, then verify R1703
            bool isVerifyR1703 = (typeof(byte) == ropOpenMessageResponse.HasNamedProperties.GetType()) &&
                                 ((Convert.ToBoolean(ropOpenMessageResponse.HasNamedProperties) == true) ||
                                 (Convert.ToBoolean(ropOpenMessageResponse.HasNamedProperties) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1703,
                1703,
                @"[In RopOpenMessage ROP Success Response Buffer]HasNamedProperties (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1705");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1705
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropOpenMessageResponse.SubjectPrefix.GetType(),
                1705,
                @"[In RopOpenMessage ROP Success Response Buffer]SubjectPrefix (variable): A TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1708");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1708
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropOpenMessageResponse.NormalizedSubject.GetType(),
                1708,
                @"[In RopOpenMessage ROP Success Response Buffer]NormalizedSubject (variable): TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1711");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1711
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOpenMessageResponse.RecipientCount.GetType(),
                1711,
                @"[In RopOpenMessage ROP Success Response Buffer]RecipientCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1713");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1713
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOpenMessageResponse.ColumnCount.GetType(),
                1713,
                @"[In RopOpenMessage ROP Success Response Buffer]ColumnCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1715");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1715
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                ropOpenMessageResponse.RecipientColumns.GetType(),
                1715,
                @"[In RopOpenMessage ROP Success Response Buffer]RecipientColumns (variable): An array of PropertyTag structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1716");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1716
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropOpenMessageResponse.ColumnCount,
                (ushort)ropOpenMessageResponse.RecipientColumns.Length,
                1716,
                @"[In RopOpenMessage ROP Success Response Buffer,RecipientColumns (variable)]The number of structures contained in this field[RecipientColumns (variable)] is specified by the ColumnCount field.");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1718
            // Assume the result is true. If RecipientColumns specifies property values don't be included in each row that is specified in the RecipientRows field, then false.
            // Check the number of property values in each row that is specified in the RecipientRows whether equal to the RecipientColumns count.
            bool isVerifyR1718 = true;
            foreach (OpenRecipientRow openRecipientRow in ropOpenMessageResponse.RecipientRows)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1718, the actual count of PropertyValues is {0} and length of RecipientColumns is {1}", openRecipientRow.RecipientRow.RecipientProperties.PropertyValues.Count, ropOpenMessageResponse.RecipientColumns.Length);
                isVerifyR1718 = isVerifyR1703 &&
                                (openRecipientRow.RecipientRow.RecipientProperties.PropertyValues.Count == ropOpenMessageResponse.RecipientColumns.Length);

                // If one row of RecipientRows don't include property values that is specified in RecipientColumns then break iteration
                if (!isVerifyR1718)
                {
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1718,
                1718,
                @"[In RopOpenMessage ROP Success Response Buffer,RecipientColumns (variable)]This field specifies the property values that can be included in each row that is specified in the RecipientRows field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1719");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1719
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenMessageResponse.RowCount.GetType(),
                1719,
                @"[In RopOpenMessage ROP Success Response Buffer]RowCount (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1722");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1722
            Site.CaptureRequirementIfAreEqual<byte>(
                ropOpenMessageResponse.RowCount,
                (byte)ropOpenMessageResponse.RecipientRows.Length,
                1722,
                @"[In RopOpenMessage ROP Success Response Buffer,RecipientRows (variable)]The number of structures contained in this field is specified by the RowCount field.");

            // RecipientRows is array of OpenRecipientRow if it has element can get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
            if (ropOpenMessageResponse.RecipientRows.Length > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1724");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1724
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropOpenMessageResponse.RecipientRows[0].RecipientType.GetType(),
                    1724,
                    @"[In OpenRecipientRow structure]RecipientType (1 byte): An enumeration.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1725,RecipientType:{0}", ropOpenMessageResponse.RecipientRows[0].RecipientType);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1725
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                bool isVerifyR1725 = ((ropOpenMessageResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.PrimaryRecipient)) == ((byte)RecipientType.PrimaryRecipient))
                                     || ((ropOpenMessageResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.CcRecipient)) == ((byte)RecipientType.CcRecipient))
                                     || ((ropOpenMessageResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.BccRecipient)) == ((byte)RecipientType.BccRecipient));

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1725,
                    1725,
                    @"[In OpenRecipientRow structure,RecipientType (1 byte)]The possible values[0x01, 0x02 and 0x03] for this enumeration are specified in [MS-OXCMSG] section 2.2.3.1.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1727");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1727
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropOpenMessageResponse.RecipientRows[0].CodePageId),
                    1727,
                    @"[In OpenRecipientRow structure]CodePageId (2 bytes): An identifier.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1729");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1729
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropOpenMessageResponse.RecipientRows[0].Reserved),
                    1729,
                    @"[In OpenRecipientRow structure]Reserved (2 bytes): Reserved.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1730");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1730
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<ushort>(
                    ReservedTwoBytes,
                    ropOpenMessageResponse.RecipientRows[0].Reserved,
                    1730,
                    @"[In OpenRecipientRow structure,Reserved (2 bytes)]The server MUST set this field[Reserved (2 bytes)] to 0x0000.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1731");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1731
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropOpenMessageResponse.RecipientRows[0].RecipientRowSize.GetType(),
                    1731,
                    @"[In OpenRecipientRow structure]RecipientRowSize (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1733");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1733
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(RecipientRow),
                    ropOpenMessageResponse.RecipientRows[0].RecipientRow.GetType(),
                    1733,
                    @"[In OpenRecipientRow structure]RecipientRow (variable): A RecipientRow structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1735");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1735
                // RecipientRows is array of OpenRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of OpenRecipientRow to verify OpenRecipientRow structure
                Site.CaptureRequirementIfAreEqual<ushort>(
                    ropOpenMessageResponse.RecipientRows[0].RecipientRowSize,
                    (ushort)ropOpenMessageResponse.RecipientRows[0].RecipientRow.Size(),
                    1735,
                    @"[In OpenRecipientRow structure,RecipientRow (variable)]The size of this field, in bytes, is specified by the RecipientRowSize field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1721");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1721
                Site.CaptureRequirementIfAreEqual<Type>(
                    ropOpenMessageResponse.RecipientRows.GetType(),
                    typeof(OpenRecipientRow[]),
                    1721,
                    @"[In RopOpenMessage ROP Success Response Buffer]RecipientRows (variable): A List of OpenRecipientRow structures.");

                foreach (OpenRecipientRow openRecipientRow in ropOpenMessageResponse.RecipientRows)
                {
                    this.VerifyRecipientRowStructure(openRecipientRow.RecipientRow);
                }
            }
        }

        /// <summary>
        /// Verify RecipientRow Structure
        /// </summary>
        /// <param name="recipientRow">The value of RecipientRow</param>
        private void VerifyRecipientRowStructure(RecipientRow recipientRow)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R131");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R131
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                recipientRow.AddressPrfixUsed.GetType(),
                "MS-OXCDATA",
                131,
                @"[In RecipientRow Structure]AddressPrefixUsed (optional) (1 byte): Unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R134");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R134
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                recipientRow.DisplayType.GetType(),
                "MS-OXCDATA",
                134,
                @"[In RecipientRow Structure] DisplayType (optional) (1 byte): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R2764");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R2764
            bool isVerifyR2764 = recipientRow.DisplayType == 0 ||
                            recipientRow.DisplayType == 1 ||
                            recipientRow.DisplayType == 2 ||
                            recipientRow.DisplayType == 3 ||
                            recipientRow.DisplayType == 4 ||
                            recipientRow.DisplayType == 5 ||
                            recipientRow.DisplayType == 6;
            Site.CaptureRequirementIfIsTrue(
                isVerifyR2764,
                "MS-OXCDATA",
                2764,
                @"[In RecipientRow Structure] Valid values [0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06] for this field [DisplayType (optional) (1 byte)] are specified in the following table.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R140");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R140
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                recipientRow.EntryIdSize.GetType(),
                "MS-OXCDATA",
                140,
                @"[In RecipientRow Structure] EntryIdSize (optional) (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R148");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R148
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                recipientRow.SearchKeySize.GetType(),
                "MS-OXCDATA",
                148,
                @"[In RecipientRow Structure] SearchKeySize (optional) (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R176");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R176
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                recipientRow.RecipientColumnCount.GetType(),
                "MS-OXCDATA",
                176,
                @"[In RecipientRow Structure] RecipientColumnCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R177");

            // Verify MS-OXCROPS requirement: MS-OXCDATA_R177
            Site.CaptureRequirementIfAreEqual<int>(
                recipientRow.RecipientProperties.PropertyValues.Count,
                recipientRow.RecipientColumnCount,
                "MS-OXCDATA",
                177,
                @"[In RecipientRow] RecipientColumnCount (2 bytes): This value [RecipientColumnCount] specifies the number of columns from the RecipientColumns field that are included in the RecipientProperties field.");
        }

        /// <summary>
        /// Verify RopOpenMessage Failure Response
        /// </summary>
        /// <param name="ropOpenMessageResponse">The response of RopOpenMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenMessage request</param>
        private void VerifyRopOpenMessageFailureResponse(RopOpenMessageResponse ropOpenMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1736");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1736
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenMessageResponse.RopId.GetType(),
                1736,
                @"[In RopOpenMessage ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1738");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1738
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenMessage,
                ropOpenMessageResponse.RopId,
                1738,
                @"[In RopOpenMessage ROP Failure Response Buffer,RopId (1 byte)] For this[RopOpenMessage] operation, this field is set to 0x03.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1739");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1739
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenMessageResponse.OutputHandleIndex.GetType(),
                1739,
                @"[In RopOpenMessage ROP Failure Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1740");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1740
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenMessageResponse.OutputHandleIndex,
                1740,
                @"[In RopOpenMessage ROP Failure Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1742");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1742
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenMessageResponse.ReturnValue.GetType(),
                1742,
                @"[In RopOpenMessage ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1744");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1744
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropOpenMessageResponse.ReturnValue,
                1744,
                @"[In RopOpenMessage ROP Failure Response Buffer,ReturnValue (4 bytes)]For this[Failure Response] response, this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopCreateMessage
        /// <summary>
        /// Verify RopCreateMessage Success Response
        /// </summary>
        /// <param name="ropCreateMessageResponse">The response of RopCreateMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateMessage request</param>
        private void VerifyRopCreateMessageSuccessResponse(RopCreateMessageResponse ropCreateMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1764");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1764
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateMessageResponse.RopId.GetType(),
                1764,
                @"[In RopCreateMessage ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1766");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1766
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateMessage,
                ropCreateMessageResponse.RopId,
                1766,
                @"[In RopCreateMessage ROP Success Response Buffer,RopId (1 byte)]For this[RopCreateMessage] operation, this field is set to 0x06.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1767");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1767
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateMessageResponse.OutputHandleIndex.GetType(),
                1767,
                @"[In RopCreateMessage ROP Success Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1768");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1768
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateMessageResponse.OutputHandleIndex,
                1768,
                @"[In RopCreateMessage ROP Success Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex specified in field the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1770");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1770
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateMessageResponse.ReturnValue.GetType(),
                1770,
                @"[In RopCreateMessage ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1772");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1772
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCreateMessageResponse.ReturnValue,
                1772,
                @"[In RopCreateMessage ROP Success Response Buffer,ReturnValue (4 bytes)]For this[Success Response] response, this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1773,HasMessageId:{0}", ropCreateMessageResponse.HasMessageId);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1773
            // If HasMessageId is byte type and can be boolean type representation, then verify R1773
            bool isVerifyR1773 = (typeof(byte) == ropCreateMessageResponse.HasMessageId.GetType()) &&
                                 ((Convert.ToBoolean(ropCreateMessageResponse.HasMessageId) == true) ||
                                 (Convert.ToBoolean(ropCreateMessageResponse.HasMessageId) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1773,
                1773,
                @"[In RopCreateMessage ROP Success Response Buffer]HasMessageId (1 byte): A Boolean.");

            if (ropCreateMessageResponse.HasMessageId != 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1775");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1775
                Site.CaptureRequirementIfAreEqual<int>(
                    8,
                    Marshal.SizeOf(ropCreateMessageResponse.MessageId),
                    1775,
                    @"[In RopCreateMessage ROP Success Response Buffer]MessageId (8 bytes): An identifier.");
            }
        }

        /// <summary>
        /// Verify RopCreateMessage Failure Response
        /// </summary>
        /// <param name="ropCreateMessageResponse">The response of RopCreateMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateMessage request</param>
        private void VerifyRopCreateMessageFailureResponse(RopCreateMessageResponse ropCreateMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1781");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1781
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateMessageResponse.RopId.GetType(),
                1781,
                @"[In RopCreateMessage ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1783");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1783
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateMessage,
                ropCreateMessageResponse.RopId,
                1783,
                @"[In RopCreateMessage ROP Failure Response Buffer,RopId (1 byte)] For this[RopCreateMessage] operation, this field is set to 0x06.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1784");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1784
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateMessageResponse.OutputHandleIndex.GetType(),
                1784,
                @"[In RopCreateMessage ROP Failure Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1785");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1785
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateMessageResponse.OutputHandleIndex,
                1785,
                @"[In RopCreateMessage ROP Failure Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1787");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1787
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateMessageResponse.ReturnValue.GetType(),
                1787,
                @"[In RopCreateMessage ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1789");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1789
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropCreateMessageResponse.ReturnValue,
                1789,
                @"[In RopCreateMessage ROP Failure Response Buffer,ReturnValue (4 bytes)]For this[Failure Response] response, this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSaveChangesMessage
        /// <summary>
        /// Verify RopSaveChangesMessage Success Response
        /// </summary>
        /// <param name="ropSaveChangesMessageResponse">The response of RopSaveChangesMessage request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSaveChangesMessage request</param>
        /// <param name="responseHandleIndex">The field of ResponseHandleIndex in RopSaveChangesMessage request</param>
        private void VerifyRopSaveChangesMessageSuccessResponse(RopSaveChangesMessageResponse ropSaveChangesMessageResponse, byte inputHandleIndex, byte responseHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1806");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1806
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesMessageResponse.RopId.GetType(),
                1806,
                @"[In RopSaveChangesMessage ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1808");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1808
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSaveChangesMessage,
                ropSaveChangesMessageResponse.RopId,
                1808,
                @"[In RopSaveChangesMessage ROP Success Response Buffer,RopId (1 byte)]For this[RopSaveChangesMessage] operation, this field is set to 0x0C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1809");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1809
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesMessageResponse.ResponseHandleIndex.GetType(),
                1809,
                @"[In RopSaveChangesMessage ROP Success Response Buffer]ResponseHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1810");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1810
            Site.CaptureRequirementIfAreEqual<byte>(
                responseHandleIndex,
                ropSaveChangesMessageResponse.ResponseHandleIndex,
                1810,
                @"[In RopSaveChangesMessage ROP Success Response Buffer,ResponseHandleIndex (1 byte)]This index MUST be set to the value specified in the ResponseHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1812");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1812
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSaveChangesMessageResponse.ReturnValue.GetType(),
                1812,
                @"[In RopSaveChangesMessage ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1814");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1814
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSaveChangesMessageResponse.ReturnValue,
                1814,
                @"[In RopSaveChangesMessage ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1815");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1815
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesMessageResponse.InputHandleIndex.GetType(),
                1815,
                @"[In RopSaveChangesMessage ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1816");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1816
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSaveChangesMessageResponse.InputHandleIndex,
                1816,
                @"[In RopSaveChangesMessage ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1818");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1818
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropSaveChangesMessageResponse.MessageId),
                1818,
                @"[In RopSaveChangesMessage ROP Success Response Buffer]MessageId (8 bytes): An identifier.");
        }

        /// <summary>
        /// Verify RopSaveChangesMessage Failure Response
        /// </summary>
        /// <param name="ropSaveChangesMessageResponse">The response of RopSaveChangesMessage request</param>
        /// <param name="responseHandleIndex">The field of ResponseHandleIndex in RopSaveChangesMessage request</param>
        private void VerifyRopSaveChangesMessageFailureResponse(RopSaveChangesMessageResponse ropSaveChangesMessageResponse, byte responseHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1820");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1820
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesMessageResponse.RopId.GetType(),
                1820,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1822");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1822
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSaveChangesMessage,
                ropSaveChangesMessageResponse.RopId,
                1822,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopSaveChangesMessage], this field is set to 0x0C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1823");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1823
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesMessageResponse.ResponseHandleIndex.GetType(),
                1823,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer]ResponseHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1824");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1824
            Site.CaptureRequirementIfAreEqual<byte>(
                responseHandleIndex,
                ropSaveChangesMessageResponse.ResponseHandleIndex,
                1824,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer,ResponseHandleIndex (1 byte)]This index MUST be set to the value specified in the ResponseHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1826");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1826
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSaveChangesMessageResponse.ReturnValue.GetType(),
                1826,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1828");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1828
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSaveChangesMessageResponse.ReturnValue,
                1828,
                @"[In RopSaveChangesMessage ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopRemoveAllRecipients
        /// <summary>
        /// Verify RopRemoveAllRecipients Response
        /// </summary>
        /// <param name="ropRemoveAllRecipientsResponse">The response of RopRemoveAllRecipients request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopRemoveAllRecipients request</param>
        private void VerifyRopRemoveAllRecipientsResponse(RopRemoveAllRecipientsResponse ropRemoveAllRecipientsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1842");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1842
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRemoveAllRecipientsResponse.RopId.GetType(),
                1842,
                @"[In RopRemoveAllRecipients ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1844");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1844
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopRemoveAllRecipients,
                ropRemoveAllRecipientsResponse.RopId,
                1844,
                @"[In RopRemoveAllRecipients ROP Response Buffer,RopId (1 byte)]For this operation[RopRemoveAllRecipients], this field is set to 0x0D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1845");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1845
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRemoveAllRecipientsResponse.InputHandleIndex.GetType(),
                1845,
                @"[In RopRemoveAllRecipients ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1846");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1846
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropRemoveAllRecipientsResponse.InputHandleIndex,
                1846,
                @"[In RopRemoveAllRecipients ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request..");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1848");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1848
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropRemoveAllRecipientsResponse.ReturnValue.GetType(),
                1848,
                @"[In RopRemoveAllRecipients ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopModifyRecipients
        /// <summary>
        /// Verify RopModifyRecipients Response
        /// </summary>
        /// <param name="ropModifyRecipientsResponse">The response of RopModifyRecipients request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopModifyRecipients request</param>
        private void VerifyRopModifyRecipientsResponse(RopModifyRecipientsResponse ropModifyRecipientsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1883");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1883
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyRecipientsResponse.RopId.GetType(),
                1883,
                @"[In RopModifyRecipients ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1885");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1885
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopModifyRecipients,
                ropModifyRecipientsResponse.RopId,
                1885,
                @"[In RopModifyRecipients ROP Response Buffer,RopId (1 byte)]For this operation[RopModifyRecipients], this field is set to 0x0E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1886");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1886
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyRecipientsResponse.InputHandleIndex.GetType(),
                1886,
                @"[In RopModifyRecipients ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1887");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1887
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropModifyRecipientsResponse.InputHandleIndex,
                1887,
                @"[In RopModifyRecipients ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1889");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1889
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropModifyRecipientsResponse.ReturnValue.GetType(),
                1889,
                @"[In RopModifyRecipients ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopReadRecipients
        /// <summary>
        /// Verify RopReadRecipients Success Response
        /// </summary>
        /// <param name="ropReadRecipientsResponse">The response of RopReadRecipients request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReadRecipients request</param>
        private void VerifyRopReadRecipientsSuccessResponse(RopReadRecipientsResponse ropReadRecipientsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1906");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1906
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadRecipientsResponse.RopId.GetType(),
                1906,
                @"[In RopReadRecipients ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1908");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1908
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReadRecipients,
                ropReadRecipientsResponse.RopId,
                1908,
                @"[In RopReadRecipients ROP Success Response Buffer,RopId (1 byte)]For this operation[RopReadRecipients], this field is set to 0x0F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1909");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1909
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadRecipientsResponse.InputHandleIndex.GetType(),
                1909,
                @"[In RopReadRecipients ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1910");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1910
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReadRecipientsResponse.InputHandleIndex,
                1910,
                @"[In RopReadRecipients ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1912");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1912
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReadRecipientsResponse.ReturnValue.GetType(),
                1912,
                @"[In RopReadRecipients ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1914");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1914
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropReadRecipientsResponse.ReturnValue,
                1914,
                @"[In RopReadRecipients ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1915");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1915
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadRecipientsResponse.RowCount.GetType(),
                1915,
                @"[In RopReadRecipients ROP Success Response Buffer]RowCount (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1918");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1918
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropReadRecipientsResponse.RowCount,
                (ushort)ropReadRecipientsResponse.RecipientRows.Length,
                1918,
                @"[In RopReadRecipients ROP Success Response Buffer,RecipientRows (variable)]The number of structures contained in this field is specified by the RowCount field.");

            // RecipientRows is array of ReadRecipientRow if it has element can get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
            if (ropReadRecipientsResponse.RecipientRows.Length > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1920");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1920
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(uint),
                    ropReadRecipientsResponse.RecipientRows[0].RowId.GetType(),
                    1920,
                    @"[In ReadRecipientRow structure]RowId (4 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1922");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1922
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<int>(
                    1,
                    Marshal.SizeOf(ropReadRecipientsResponse.RecipientRows[0].RecipientType),
                    1922,
                    @"[In ReadRecipientRow structure]RecipientType (1 byte): An enumeration.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1923,RecipientType:{0}", ropReadRecipientsResponse.RecipientRows[0].RecipientType);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1923
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                bool isVerifyR1923 = ((ropReadRecipientsResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.PrimaryRecipient)) == ((byte)RecipientType.PrimaryRecipient))
                                     || ((ropReadRecipientsResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.CcRecipient)) == ((byte)RecipientType.CcRecipient))
                                     || ((ropReadRecipientsResponse.RecipientRows[0].RecipientType & ((byte)RecipientType.BccRecipient)) == ((byte)RecipientType.BccRecipient));

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1923,
                    1923,
                    @"[In ReadRecipientRow structure,RecipientType (1 byte)]The possible values[0x01, 0x02 and 0x03] for this enumeration are specified in [MS-OXCMSG] section 2.2.3.1.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1925");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1925
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropReadRecipientsResponse.RecipientRows[0].CodePageId),
                    1925,
                    @"[In ReadRecipientRow structure]CodePageId (2 bytes): An identifier.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1927");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1927
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<int>(
                    2,
                    Marshal.SizeOf(ropReadRecipientsResponse.RecipientRows[0].Reserved),
                    1927,
                    @"[In ReadRecipientRow structure]Reserved (2 bytes): Reserved.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1928");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1928
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    ropReadRecipientsResponse.RecipientRows[0].Reserved,
                    1928,
                    @"[In ReadRecipientRow structure,Reserved (2 bytes)]The server MUST set this field to 0x0000.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1929");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1929
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    ropReadRecipientsResponse.RecipientRows[0].RecipientRowSize.GetType(),
                    1929,
                    @"[In ReadRecipientRow structure]RecipientRowSize (2 bytes): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1931");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1931
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(RecipientRow),
                    ropReadRecipientsResponse.RecipientRows[0].RecipientRow.GetType(),
                    1931,
                    @"[In ReadRecipientRow structure]RecipientRow (variable): A RecipientRow structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1933");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1933
                // RecipientRows is array of ReadRecipientRow, in the precondition statement RecipientRows has element.
                // Get the first element as instance of ReadRecipientRow to verify ReadRecipientRow structure
                Site.CaptureRequirementIfAreEqual<ushort>(
                    ropReadRecipientsResponse.RecipientRows[0].RecipientRowSize,
                    (ushort)ropReadRecipientsResponse.RecipientRows[0].RecipientRow.Size(),
                    1933,
                    @"[In ReadRecipientRow structure,RecipientRow (variable)]The size of this field, in bytes, is specified by the RecipientRowSize field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1917");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R1917
                Site.CaptureRequirementIfAreEqual<Type>(
                    ropReadRecipientsResponse.RecipientRows.GetType(),
                    typeof(ReadRecipientRow[]),
                    1917,
                    @"[In RopReadRecipients ROP Success Response Buffer]RecipientRows (variable): A list of ReadRecipientRow structures.");

                foreach (ReadRecipientRow readRecipientRow in ropReadRecipientsResponse.RecipientRows)
                {
                    this.VerifyRecipientRowStructure(readRecipientRow.RecipientRow);
                }
            }
        }

        /// <summary>
        /// Verify RopReadRecipients Failure Response
        /// </summary>
        /// <param name="ropReadRecipientsResponse">The response of RopReadRecipients request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReadRecipients request</param>
        private void VerifyRopReadRecipientsFailureResponse(RopReadRecipientsResponse ropReadRecipientsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1934");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1934
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadRecipientsResponse.RopId.GetType(),
                1934,
                @"[In RopReadRecipients ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1936");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1936
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReadRecipients,
                ropReadRecipientsResponse.RopId,
                1936,
                @"[In RopReadRecipients ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopReadRecipients], this field is set to 0x0F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1937");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1937
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadRecipientsResponse.InputHandleIndex.GetType(),
                1937,
                @"[In RopReadRecipients ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1938");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1938
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReadRecipientsResponse.InputHandleIndex,
                1938,
                @"[In RopReadRecipients ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1940");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1940
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReadRecipientsResponse.ReturnValue.GetType(),
                1940,
                @"[In RopReadRecipients ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1942");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1942
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropReadRecipientsResponse.ReturnValue,
                1942,
                @"[In RopReadRecipients ROP Failure Response Buffer,ReturnValue (4 bytes)] For this response, this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopReloadCachedInformation
        /// <summary>
        /// Verify RopReloadCachedInformation Success Response
        /// </summary>
        /// <param name="ropReloadCachedInformationResponse">The response of RopReloadCachedInformation request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReloadCachedInformation request</param>
        private void VerifyRopReloadCachedInformationSuccessResponse(RopReloadCachedInformationResponse ropReloadCachedInformationResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1956");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1956
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReloadCachedInformationResponse.RopId.GetType(),
                1956,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1958");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1958
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReloadCachedInformation,
                ropReloadCachedInformationResponse.RopId,
                1958,
                @"[In RopReloadCachedInformation ROP Success Response Buffer,RopId (1 byte)]For this operation[RopReloadCachedInformation], this field is set to 0x10.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1959");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1959
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReloadCachedInformationResponse.InputHandleIndex.GetType(),
                1959,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1960");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1960
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReloadCachedInformationResponse.InputHandleIndex,
                1960,
                @"[In RopReloadCachedInformation ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex specified field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1962");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1962
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReloadCachedInformationResponse.ReturnValue.GetType(),
                1962,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1964");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1964
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropReloadCachedInformationResponse.ReturnValue,
                1964,
                @"[In RopReloadCachedInformation ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1965,HasNamedProperties:{0}", ropReloadCachedInformationResponse.HasNamedProperties);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1965
            // If HasNamedProperties is byte type and can be boolean type representation, then verify R1965
            bool isVerifyR1965 = (typeof(byte) == ropReloadCachedInformationResponse.HasNamedProperties.GetType()) &&
                                 ((Convert.ToBoolean(ropReloadCachedInformationResponse.HasNamedProperties) == true) ||
                                 (Convert.ToBoolean(ropReloadCachedInformationResponse.HasNamedProperties) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR1965,
                1965,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]HasNamedProperties (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1967");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1967
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropReloadCachedInformationResponse.SubjectPrefix.GetType(),
                1967,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]SubjectPrefix (variable): A TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1970");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1970
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropReloadCachedInformationResponse.NormalizedSubject.GetType(),
                1970,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]NormalizedSubject (variable): A TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1973");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1973
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropReloadCachedInformationResponse.RecipientCount.GetType(),
                1973,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]RecipientCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1975");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1975
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropReloadCachedInformationResponse.ColumnCount.GetType(),
                1975,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]ColumnCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1977");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1977
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                ropReloadCachedInformationResponse.RecipientColumns.GetType(),
                1977,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]RecipientColumns (variable): An array of PropertyTag structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1978");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1978
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropReloadCachedInformationResponse.ColumnCount,
                (ushort)ropReloadCachedInformationResponse.RecipientColumns.Length,
                1978,
                @"[In RopReloadCachedInformation ROP Success Response Buffer,RecipientColumns (variable)]The number of structures contained in this field[RecipientColumns (variable)] is specified by the ColumnCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1981");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1981
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReloadCachedInformationResponse.RowCount.GetType(),
                1981,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]RowCount (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1983");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1983
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(OpenRecipientRow[]),
                ropReloadCachedInformationResponse.RecipientRows.GetType(),
                1983,
                @"[In RopReloadCachedInformation ROP Success Response Buffer]RecipientRows (variable): A list of OpenRecipientRow structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1984");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1984
            Site.CaptureRequirementIfAreEqual<byte>(
                ropReloadCachedInformationResponse.RowCount,
                (byte)ropReloadCachedInformationResponse.RecipientRows.Length,
                1984,
                @"[In RopReloadCachedInformation ROP Success Response Buffer,RecipientRows (variable)]The number of structures contained in this field[RecipientRows (variable)] is specified by the RowCount field.");
        }

        /// <summary>
        /// Verify RopReloadCachedInformation Failure Response
        /// </summary>
        /// <param name="ropReloadCachedInformationResponse">The response of RopReloadCachedInformation request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReloadCachedInformation request</param>
        private void VerifyRopReloadCachedInformationFailureResponse(RopReloadCachedInformationResponse ropReloadCachedInformationResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1986");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1986
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReloadCachedInformationResponse.RopId.GetType(),
                1986,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1988");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1988
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReloadCachedInformation,
                ropReloadCachedInformationResponse.RopId,
                1988,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopReloadCachedInformation], this field is set to 0x10.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1989");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1989
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReloadCachedInformationResponse.InputHandleIndex.GetType(),
                1989,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1990");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1990
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReloadCachedInformationResponse.InputHandleIndex,
                1990,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1992");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1992
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReloadCachedInformationResponse.ReturnValue.GetType(),
                1992,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R1994");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R1994
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropReloadCachedInformationResponse.ReturnValue,
                1994,
                @"[In RopReloadCachedInformation ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSetMessageStatus
        /// <summary>
        /// Verify RopSetMessageStatus Success Response
        /// </summary>
        /// <param name="ropSetMessageStatusResponse">The response of RopSetMessageStatus request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetMessageStatus request</param>
        private void VerifyRopSetMessageStatusSuccessResponse(RopSetMessageStatusResponse ropSetMessageStatusResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2013");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2013
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageStatusResponse.RopId.GetType(),
                2013,
                @"[In RopSetMessageStatus ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2015");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2015
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetMessageStatus,
                ropSetMessageStatusResponse.RopId,
                2015,
                @"[In RopSetMessageStatus ROP Success Response Buffer,RopId (1 byte)]For this operation[RopSetMessageStatus], this field is set to 0x20.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2016");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2016
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageStatusResponse.InputHandleIndex.GetType(),
                2016,
                @"[In RopSetMessageStatus ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2017");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2017
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetMessageStatusResponse.InputHandleIndex,
                2017,
                @"[In RopSetMessageStatus ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2019");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2019
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetMessageStatusResponse.ReturnValue.GetType(),
                2019,
                @"[In RopSetMessageStatus ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2021");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2021
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetMessageStatusResponse.ReturnValue,
                2021,
                @"[In RopSetMessageStatus ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2022");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2022
            Site.CaptureRequirementIfAreEqual<int>(
                4,
                Marshal.SizeOf(ropSetMessageStatusResponse.MessageStatusFlags),
                2022,
                @"[In RopSetMessageStatus ROP Success Response Buffer]MessageStatusFlags (4 bytes): A flags structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2024,MessageStatusFlags:{0}", ropSetMessageStatusResponse.MessageStatusFlags);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2024
            bool isVerifyR2024 = (ropSetMessageStatusResponse.MessageStatusFlags == (uint)MessageStatusFlags.None)
                                 || ((ropSetMessageStatusResponse.MessageStatusFlags & (uint)MessageStatusFlags.MsRemoteDownload) == (uint)MessageStatusFlags.MsRemoteDownload)
                                 || ((ropSetMessageStatusResponse.MessageStatusFlags & (uint)MessageStatusFlags.MsInConflict) == (uint)MessageStatusFlags.MsInConflict)
                                 || ((ropSetMessageStatusResponse.MessageStatusFlags & (uint)MessageStatusFlags.MsRemoteDelete) == (uint)MessageStatusFlags.MsRemoteDelete);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2024,
                2024,
                @"[In RopSetMessageStatus ROP Success Response Buffer,MessageStatusFlags (4 bytes)]The possible values[zero, the value of msRemoteDownload is 0x00001000, the value of msInConflict is 0x00000800 and the value of msRemoteDelete is 0x00002000] are specified in [MS-OXCMSG] section 2.2.3.8.2.");
           
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R20004");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R20004
            this.Site.CaptureRequirementIfAreEqual<byte>(
                0x20,
                ropSetMessageStatusResponse.RopId,
                20004,
                @"[In RopGetMessageStatus ROP Response Buffers] The value of the RopId field for RopGetMessageStatus responses MUST be 0x20, which is the same as that for the RopSetMessageStatus responses.");

        }

        /// <summary>
        /// Verify RopSetMessageStatus Failure Response
        /// </summary>
        /// <param name="ropSetMessageStatusResponse">The response of RopSetMessageStatus request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetMessageStatus request</param>
        private void VerifyRopSetMessageStatusFailureResponse(RopSetMessageStatusResponse ropSetMessageStatusResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2025");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2025
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageStatusResponse.RopId.GetType(),
                2025,
                @"[In RopSetMessageStatus ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2027");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2027
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetMessageStatus,
                ropSetMessageStatusResponse.RopId,
                2027,
                @"[In RopSetMessageStatus ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopSetMessageStatus], this field is set to 0x20.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2028");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2028
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageStatusResponse.InputHandleIndex.GetType(),
                2028,
                @"[In RopSetMessageStatus ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2029");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2029
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetMessageStatusResponse.InputHandleIndex,
                2029,
                @"[In RopSetMessageStatus ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2031");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2031
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetMessageStatusResponse.ReturnValue.GetType(),
                2031,
                @"[In RopSetMessageStatus ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2033");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2033
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetMessageStatusResponse.ReturnValue,
                2033,
                @"[In RopSetMessageStatus ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSetReadFlags
        /// <summary>
        /// Verify RopSetReadFlags Response
        /// </summary>
        /// <param name="ropSetReadFlagsResponse">The response of RopSetReadFlags request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetReadFlags request</param>
        private void VerifyRopSetReadFlagsResponse(RopSetReadFlagsResponse ropSetReadFlagsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2067");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2067
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetReadFlagsResponse.RopId.GetType(),
                2067,
                @"[In RopSetReadFlags ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2069");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2069
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetReadFlags,
                ropSetReadFlagsResponse.RopId,
                2069,
                @"[In RopSetReadFlags ROP Response Buffer,RopId (1 byte)]For this operation[RopSetReadFlags], this field is set to 0x66.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2070");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2070
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetReadFlagsResponse.InputHandleIndex.GetType(),
                2070,
                @"[In RopSetReadFlags ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2071");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2071
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetReadFlagsResponse.InputHandleIndex,
                2071,
                @"[In RopSetReadFlags ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2073");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2073
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetReadFlagsResponse.ReturnValue.GetType(),
                2073,
                @"[In RopSetReadFlags ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2075,PartialCompletion:{0}", ropSetReadFlagsResponse.PartialCompletion);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2075
            // If PartialCompletion is byte type and can be boolean type representation, then verify R2075
            bool isVerifyR2075 = (typeof(byte) == ropSetReadFlagsResponse.PartialCompletion.GetType()) &&
                                 ((Convert.ToBoolean(ropSetReadFlagsResponse.PartialCompletion) == true) ||
                                 (Convert.ToBoolean(ropSetReadFlagsResponse.PartialCompletion) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2075,
                2075,
                @"[In RopSetReadFlags ROP Response Buffer]PartialCompletion (1 byte): An Boolean.");
        }
        #endregion

        #region RopSetMessageReadFlag
        /// <summary>
        /// Verify RopSetMessageReadFlag Success Response
        /// </summary>
        /// <param name="ropSetMessageReadFlagResponse">The response of RopSetMessageReadFlag request</param>
        /// <param name="responseHandleIndex">The field of ResponseHandleIndex in RopSetMessageReadFlag request</param>
        /// <param name="logonId">The field of LogonId in RopSetMessageReadFlag request</param>
        /// <param name="clientData">The field of ClientData in RopSetMessageReadFlag request</param>
        private void VerifyRopSetMessageReadFlagSuccessResponse(
            RopSetMessageReadFlagResponse ropSetMessageReadFlagResponse,
            byte responseHandleIndex,
            byte logonId,
            byte[] clientData)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2097");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2097
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageReadFlagResponse.RopId.GetType(),
                2097,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer]RopId (1 byte): An unsigned integer");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2099");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2099
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetMessageReadFlag,
                ropSetMessageReadFlagResponse.RopId,
                2099,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer,RopId (1 byte)]For this operation[RopSetMessageReadFlag], this field is set to 0x11.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2100");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2100
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageReadFlagResponse.ResponseHandleIndex.GetType(),
                2100,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer]ResponseHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2101");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2101
            Site.CaptureRequirementIfAreEqual<byte>(
                responseHandleIndex,
                ropSetMessageReadFlagResponse.ResponseHandleIndex,
                2101,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer,ResponseHandleIndex (1 byte)]This index MUST be set to the value specified in the ResponseHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2103");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2103
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetMessageReadFlagResponse.ReturnValue.GetType(),
                2103,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2105");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2105
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetMessageReadFlagResponse.ReturnValue,
                2105,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2106,ReadStatusChanged:{0}", ropSetMessageReadFlagResponse.ReadStatusChanged);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2106
            // If ReadStatusChanged is byte type and can be boolean type representation, then verify R2106
            bool isVerifyR2106 = (typeof(byte) == ropSetMessageReadFlagResponse.ReadStatusChanged.GetType()) &&
                                 ((Convert.ToBoolean(ropSetMessageReadFlagResponse.ReadStatusChanged) == true) ||
                                 (Convert.ToBoolean(ropSetMessageReadFlagResponse.ReadStatusChanged) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2106,
                2106,
                @"[In RopSetMessageReadFlag ROP Success Response Buffer]ReadStatusChanged (1 byte): A Boolean.");

            if (ropSetMessageReadFlagResponse.LogonId != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2108");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2108
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte),
                    ropSetMessageReadFlagResponse.LogonId.GetType(),
                    2108,
                    @"[In RopSetMessageReadFlag ROP Success Response Buffer]LogonId (1 byte): An unsigned integer.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2111");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2111
                Site.CaptureRequirementIfAreEqual<byte>(
                    logonId,
                    (byte)ropSetMessageReadFlagResponse.LogonId,
                    2111,
                    @"[In RopSetMessageReadFlag ROP Success Response Buffer,LogonId (1 byte)]This field MUST be set to the value of the LogonId field in the request.");
            }

            if (ropSetMessageReadFlagResponse.ClientData != null)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2112, the actual type of ClientData is {0} and length of it is {1}", ropSetMessageReadFlagResponse.ClientData.GetType().Name, ropSetMessageReadFlagResponse.ClientData.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2112
                // If ClientData is array of byte and the length is 24, then verify R2112
                bool isVerifyR2112 = (typeof(byte[]) == ropSetMessageReadFlagResponse.ClientData.GetType()) &&
                                     (ropSetMessageReadFlagResponse.ClientData.Length == 24);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2112,
                    2112,
                    @"[In RopSetMessageReadFlag ROP Success Response Buffer]ClientData (24 bytes):An array of bytes.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2115, the ClientData value in request is {0} and in response is {1}", BitConverter.ToString(clientData), BitConverter.ToString(ropSetMessageReadFlagResponse.ClientData));

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2115
                // Assume the result is true. If the parameter clientData not equal response clientData, then false.
                bool isVerifyR2115 = true;
                if (clientData.Length != ropSetMessageReadFlagResponse.ClientData.Length)
                {
                    isVerifyR2115 = false;
                }
                else
                {
                    for (int i = 0; i < ropSetMessageReadFlagResponse.ClientData.Length; i++)
                    {
                        if (ropSetMessageReadFlagResponse.ClientData[i] != clientData[i])
                        {
                            isVerifyR2115 = false;
                            break;
                        }
                    }
                }

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2115,
                    2115,
                    @"[In RopSetMessageReadFlag ROP Success Response Buffer,ClientData (24 bytes)]This field MUST be set to the value of the ClientData field in the request.");
            }
        }

        /// <summary>
        /// Verify RopSetMessageReadFlag Failure Response
        /// </summary>
        /// <param name="ropSetMessageReadFlagResponse">The response of RopSetMessageReadFlag request</param>
        /// <param name="responseHandleIndex">The field of ResponseHandleIndex in RopSetMessageReadFlag request</param>        
        private void VerifyRopSetMessageReadFlagFailureResponse(RopSetMessageReadFlagResponse ropSetMessageReadFlagResponse, byte responseHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2116");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2116
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageReadFlagResponse.RopId.GetType(),
                2116,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2118");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2118
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetMessageReadFlag,
                ropSetMessageReadFlagResponse.RopId,
                2118,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer,RopId (1 byte)] For this operation[RopSetMessageReadFlag], this field is set to 0x11.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2119");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2119
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetMessageReadFlagResponse.ResponseHandleIndex.GetType(),
                2119,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer]ResponseHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2120");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2120
            Site.CaptureRequirementIfAreEqual<byte>(
                responseHandleIndex,
                ropSetMessageReadFlagResponse.ResponseHandleIndex,
                2120,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer,ResponseHandleIndex (1 byte)]This index MUST be set to the value specified in the ResponseHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2122");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2122
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetMessageReadFlagResponse.ReturnValue.GetType(),
                2122,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2124");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2124
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetMessageReadFlagResponse.ReturnValue,
                2124,
                @"[In RopSetMessageReadFlag ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopOpenAttachment
        /// <summary>
        /// Verify RopOpenAttachment Response
        /// </summary>
        /// <param name="ropOpenAttachmentResponse">The response of RopOpenAttachment request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenAttachment request</param>
        private void VerifyRopOpenAttachmentResponse(RopOpenAttachmentResponse ropOpenAttachmentResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2143");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2143
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenAttachmentResponse.RopId.GetType(),
                2143,
                @"[In RopOpenAttachment ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2145");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2145
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenAttachment,
                ropOpenAttachmentResponse.RopId,
                2145,
                @"[In RopOpenAttachment ROP Response Buffer,RopId (1 byte)]For this operation[RopOpenAttachment], this field is set to 0x22.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2146");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2146
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenAttachmentResponse.OutputHandleIndex.GetType(),
                2146,
                @"[In RopOpenAttachment ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2147");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2147
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenAttachmentResponse.OutputHandleIndex,
                2147,
                @"[In RopOpenAttachment ROP Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2149");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2149
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenAttachmentResponse.ReturnValue.GetType(),
                2149,
                @"[In RopOpenAttachment ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopCreateAttachment
        /// <summary>
        /// Verify RopCreateAttachment Success Response
        /// </summary>
        /// <param name="ropCreateAttachmentResponse">The response of RopCreateAttachment request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateAttachment request</param>
        private void VerifyRopCreateAttachmentSuccessResponse(RopCreateAttachmentResponse ropCreateAttachmentResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2164");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2164
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateAttachmentResponse.RopId.GetType(),
                2164,
                @"[In RopCreateAttachment ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2166");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2166
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateAttachment,
                ropCreateAttachmentResponse.RopId,
                2166,
                @"[In RopCreateAttachment ROP Success Response Buffer,RopId (1 byte)]For this operation[ RopCreateAttachment], this field is set to 0x23.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2167");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2167
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateAttachmentResponse.OutputHandleIndex.GetType(),
                2167,
                @"[In RopCreateAttachment ROP Success Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2168");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2168
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateAttachmentResponse.OutputHandleIndex,
                2168,
                @"[In RopCreateAttachment ROP Success Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2170");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2170
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateAttachmentResponse.ReturnValue.GetType(),
                2170,
                @"[In RopCreateAttachment ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2172");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2172
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCreateAttachmentResponse.ReturnValue,
                2172,
                @"[In RopCreateAttachment ROP Success Response Buffer,ReturnValue (4 bytes)] For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2173");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2173
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateAttachmentResponse.AttachmentID.GetType(),
                2173,
                @"[In RopCreateAttachment ROP Success Response Buffer]AttachmentID (4 bytes): An unsigned integer identifier.");
        }

        /// <summary>
        /// Verify RopCreateAttachment Failure Response
        /// </summary>
        /// <param name="ropCreateAttachmentResponse">The response of RopCreateAttachment request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCreateAttachment request</param>
        private void VerifyRopCreateAttachmentFailureResponse(RopCreateAttachmentResponse ropCreateAttachmentResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2176");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2176
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateAttachmentResponse.RopId.GetType(),
                2176,
                @"[In RopCreateAttachment ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2178");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2178
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCreateAttachment,
                ropCreateAttachmentResponse.RopId,
                2178,
                @"[In RopCreateAttachment ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopCreateAttachment], this field is set to 0x23.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2179");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2179
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCreateAttachmentResponse.OutputHandleIndex.GetType(),
                2179,
                @"[In RopCreateAttachment ROP Failure Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2180");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2180
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCreateAttachmentResponse.OutputHandleIndex,
                2180,
                @"[In RopCreateAttachment ROP Failure Response Buffer,OutputHandleIndex (1 byte)]This index that MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2182");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2182
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCreateAttachmentResponse.ReturnValue.GetType(),
                2182,
                @"[In RopCreateAttachment ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2184");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2184
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropCreateAttachmentResponse.ReturnValue,
                2184,
                @"[In RopCreateAttachment ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopDeleteAttachment
        /// <summary>
        /// Verify RopDeleteAttachment Response
        /// </summary>
        /// <param name="ropDeleteAttachmentResponse">The response of RopDeleteAttachment request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeleteAttachment request</param>
        private void VerifyRopDeleteAttachmentResponse(RopDeleteAttachmentResponse ropDeleteAttachmentResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2198");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2198
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteAttachmentResponse.RopId.GetType(),
                2198,
                @"[In RopDeleteAttachment ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2200");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2200
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeleteAttachment,
                ropDeleteAttachmentResponse.RopId,
                2200,
                @"[In RopDeleteAttachment ROP Response Buffer,RopId (1 byte)]For this operation[RopDeleteAttachment], this field is set to 0x24.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2201");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2201
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeleteAttachmentResponse.InputHandleIndex.GetType(),
                2201,
                @"[In RopDeleteAttachment ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2202");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2202
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeleteAttachmentResponse.InputHandleIndex,
                2202,
                @"[In RopDeleteAttachment ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2204");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2204
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeleteAttachmentResponse.ReturnValue.GetType(),
                2204,
                @"[In RopDeleteAttachment ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSaveChangesAttachment
        /// <summary>
        /// Verify RopSaveChangesAttachment Response
        /// </summary>
        /// <param name="ropSaveChangesAttachmentResponse">The response of RopSaveChangesAttachment request</param>
        /// <param name="responseHandleIndex">The field of ResponseHandleIndex in RopSaveChangesAttachment request</param>
        private void VerifyRopSaveChangesAttachmentResponse(RopSaveChangesAttachmentResponse ropSaveChangesAttachmentResponse, byte responseHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2222");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2222
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesAttachmentResponse.RopId.GetType(),
                2222,
                @"[In RopSaveChangesAttachment ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2224");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2224
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSaveChangesAttachment,
                ropSaveChangesAttachmentResponse.RopId,
                2224,
                @"[In RopSaveChangesAttachment ROP Response Buffer,RopId (1 byte)]For this operation[RopSaveChangesAttachment], this field is set to 0x25.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2225");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2225
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSaveChangesAttachmentResponse.ResponseHandleIndex.GetType(),
                2225,
                @"[In RopSaveChangesAttachment ROP Response Buffer]ResponseHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2226");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2226
            Site.CaptureRequirementIfAreEqual<byte>(
                responseHandleIndex,
                ropSaveChangesAttachmentResponse.ResponseHandleIndex,
                2226,
                @"[In RopSaveChangesAttachment ROP Response Buffer,ResponseHandleIndex (1 byte)]This index MUST be set to the value specified in the ResponseHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2228");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2228
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSaveChangesAttachmentResponse.ReturnValue.GetType(),
                2228,
                @"[In RopSaveChangesAttachment ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopOpenEmbeddedMessage
        /// <summary>
        /// Verify RopOpenEmbeddedMessage Success Response
        /// </summary>
        /// <param name="ropOpenEmbeddedMessageResponse">The response of RopOpenEmbeddedMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenEmbeddedMessage request</param>
        private void VerifyRopOpenEmbeddedMessageSuccessResponse(RopOpenEmbeddedMessageResponse ropOpenEmbeddedMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2248");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2248
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.RopId.GetType(),
                2248,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2250");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2250
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenEmbeddedMessage,
                ropOpenEmbeddedMessageResponse.RopId,
                2250,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,RopId (1 byte)]For this operation[RopOpenEmbeddedMessage], this field is set to 0x46.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2251");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2251
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.OutputHandleIndex.GetType(),
                2251,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2252");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2252
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenEmbeddedMessageResponse.OutputHandleIndex,
                2252,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2254");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2254
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenEmbeddedMessageResponse.ReturnValue.GetType(),
                2254,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2256");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2256
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropOpenEmbeddedMessageResponse.ReturnValue,
                2256,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2257");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2257
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.Reserved.GetType(),
                2257,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]Reserved (1 byte): Reserved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2258");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2258
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.None,
                ropOpenEmbeddedMessageResponse.Reserved,
                2258,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,Reserved (1 byte)]This field MUST be set to 0x00.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2259");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2259
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropOpenEmbeddedMessageResponse.MessageId),
                2259,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]MessageId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2261,HasNamedProperties:{0}", ropOpenEmbeddedMessageResponse.HasNamedProperties);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2261
            // If HasNamedProperties is byte type and can be boolean type representation, then verify R2261
            bool isVerifyR2261 = (typeof(byte) == ropOpenEmbeddedMessageResponse.HasNamedProperties.GetType()) &&
                                 ((Convert.ToBoolean(ropOpenEmbeddedMessageResponse.HasNamedProperties) == true) ||
                                 (Convert.ToBoolean(ropOpenEmbeddedMessageResponse.HasNamedProperties) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2261,
                2261,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]HasNamedProperties (1 byte): A Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2263");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2263
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropOpenEmbeddedMessageResponse.SubjectPrefix.GetType(),
                2263,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]SubjectPrefix (variable): A TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2266");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2266
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TypedString),
                ropOpenEmbeddedMessageResponse.NormalizedSubject.GetType(),
                2266,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]NormalizedSubject (variable): A TypedString structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2269");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2269
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOpenEmbeddedMessageResponse.RecipientCount.GetType(),
                2269,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]RecipientCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2271");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2271
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOpenEmbeddedMessageResponse.ColumnCount.GetType(),
                2271,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]ColumnCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2273");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2273
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                ropOpenEmbeddedMessageResponse.RecipientColumns.GetType(),
                2273,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]RecipientColumns (variable): An array of PropertyTag structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2274");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2274
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropOpenEmbeddedMessageResponse.ColumnCount,
                (ushort)ropOpenEmbeddedMessageResponse.RecipientColumns.Length,
                2274,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,RecipientColumns (variable)]The number of structures contained in this field is specified by the ColumnCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2277");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2277
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.RowCount.GetType(),
                2277,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]RowCount (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2280");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2280
            Site.CaptureRequirementIfAreEqual<byte>(
                ropOpenEmbeddedMessageResponse.RowCount,
                (byte)ropOpenEmbeddedMessageResponse.RecipientRows.Length,
                2280,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer,RecipientRows (variable)]The number of structures contained in this field[RecipientRows (variable)] is specified by the RowCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2279");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2279
            Site.CaptureRequirementIfAreEqual<Type>(
                ropOpenEmbeddedMessageResponse.RecipientRows.GetType(),
                typeof(OpenRecipientRow[]),
                2279,
                @"[In RopOpenEmbeddedMessage ROP Success Response Buffer]RecipientRows (variable): A list of OpenRecipientRow structures.");
        }

        /// <summary>
        /// Verify RopOpenEmbeddedMessage Failure Response
        /// </summary>
        /// <param name="ropOpenEmbeddedMessageResponse">The response of RopOpenEmbeddedMessage request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenEmbeddedMessage request</param>
        private void VerifyRopOpenEmbeddedMessageFailureResponse(RopOpenEmbeddedMessageResponse ropOpenEmbeddedMessageResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2282");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2282
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.RopId.GetType(),
                2282,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2285");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2285
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenEmbeddedMessage,
                ropOpenEmbeddedMessageResponse.RopId,
                2285,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopOpenEmbeddedMessage], this field is set to 0x46.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2286");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2286
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenEmbeddedMessageResponse.OutputHandleIndex.GetType(),
                2286,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2287");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2287
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenEmbeddedMessageResponse.OutputHandleIndex,
                2287,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2289");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2289
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenEmbeddedMessageResponse.ReturnValue.GetType(),
                2289,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2291");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2291
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropOpenEmbeddedMessageResponse.ReturnValue,
                2291,
                @"[In RopOpenEmbeddedMessage ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopGetAttachmentTable
        /// <summary>
        /// Verify RopGetAttachmentTable Response
        /// </summary>
        /// <param name="ropGetAttachmentTableResponse">The response of RopGetAttachmentTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetAttachmentTable request</param>
        private void VerifyRopGetAttachmentTableResponse(RopGetAttachmentTableResponse ropGetAttachmentTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2308");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2308
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAttachmentTableResponse.RopId.GetType(),
                2308,
                @"[In RopGetAttachmentTable ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2310");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2310
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetAttachmentTable,
                ropGetAttachmentTableResponse.RopId,
                2310,
                @"[In RopGetAttachmentTable ROP Response Buffer,RopId (1 byte)]For this operation[RopGetAttachmentTable], this field is set to 0x21.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2311");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2311
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAttachmentTableResponse.OutputHandleIndex.GetType(),
                2311,
                @"[In RopGetAttachmentTable ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2312");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2312
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetAttachmentTableResponse.OutputHandleIndex,
                2312,
                @"[In RopGetAttachmentTable ROP Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2314");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2314
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetAttachmentTableResponse.ReturnValue.GetType(),
                2314,
                @"[In RopGetAttachmentTable ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetValidAttachments
        /// <summary>
        /// Verify RopGetValidAttachments Success Response
        /// </summary>
        /// <param name="ropGetValidAttachmentsResponse">The response of RopGetValidAttachments request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopGetAttachmentTable request</param>
        private void VerifyRopGetValidAttachmentsSuccessResponse(RopGetValidAttachmentsResponse ropGetValidAttachmentsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2324");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2324
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetValidAttachmentsResponse.RopId.GetType(),
                2324,
                @"[in RopGetValidAttachments ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2327");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2327
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetValidAttachments,
                ropGetValidAttachmentsResponse.RopId,
                2327,
                @"[in RopGetValidAttachments ROP Success Response Buffer,Ropid(1 byte)]. For this operation[RopGetValidAttachments], this field is set to 0x52.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2328");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2328
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetValidAttachmentsResponse.InputHandleIndex.GetType(),
                2328,
                @"[in RopGetValidAttachments ROP Success Response Buffer,InputHandleIndex(1 byte)].An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2329");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2329
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetValidAttachmentsResponse.InputHandleIndex,
                2329,
                @"[in RopGetValidAttachments ROP Success Response Buffer,InputHandleIndex (1 byte)].This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2331");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2331
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetValidAttachmentsResponse.ReturnValue.GetType(),
                2331,
                @"[in RopGetValidAttachments ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2333");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2333
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetValidAttachmentsResponse.ReturnValue,
                2333,
                @"[in RopGetValidAttachments ROP Success Response Buffer, ReturnValue (4 bytes)]For this response, this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2334");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2334
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetValidAttachmentsResponse.AttachmentIdCount.GetType(),
                2334,
                @"[in RopGetValidAttachments ROP Success Response Buffer]AttachmentIdCount (4 bytes):An unsigned integer. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2335");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2335
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetValidAttachmentsResponse.AttachmentIdArray.Length,
                ropGetValidAttachmentsResponse.AttachmentIdCount,
                2335,
                @"[in RopGetValidAttachments ROP Success Response Buffer,AttachmentIdCount (4 bytes)]This value specifies the number of integers in the AttachmentIdArray field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2336");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2336
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint[]),
                ropGetValidAttachmentsResponse.AttachmentIdArray.GetType(),
                2336,
                @"[in RopGetValidAttachments ROP Success Response Buffer,AttachmentIdArray (variable)]An array of 32-bit integers. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2337");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2337
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetValidAttachmentsResponse.AttachmentIdCount,
                ropGetValidAttachmentsResponse.AttachmentIdArray.Length,
                2337,
                @"[in RopGetValidAttachments ROP Success Response Buffer,AttachmentIdArray (variable)]The number of integer values contained in this field is specified by the AttachmentIdCount field. ");
        }

        /// <summary>
        /// Verify RopGetValidAttachments Failure Response
        /// </summary>
        /// <param name="ropGetValidAttachmentsResponse">The response of RopGetValidAttachments request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetValidAttachments request</param>
        private void VerifyRopGetValidAttachmentsFailureResponse(RopGetValidAttachmentsResponse ropGetValidAttachmentsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2339");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2339
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetValidAttachmentsResponse.RopId.GetType(),
                2339,
                @"[In RopGetValidAttachments ROP Failure Response Buffer]RopId(1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2347");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2347
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetValidAttachments,
                ropGetValidAttachmentsResponse.RopId,
                2347,
                @"[In RopGetValidAttachments ROP Failure Response Buffer,RopId(1 byte)]: For this operation[RopGetValidAttachments], this field is set to 0x52.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2341");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2341
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetValidAttachmentsResponse.InputHandleIndex.GetType(),
                2341,
                @"[In RopGetValidAttachments ROP Failure Response Buffer]InputHandleIndex(1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2342");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2342
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetValidAttachmentsResponse.InputHandleIndex,
                2342,
                @"[In RopGetValidAttachments ROP Failure Response Buffer,InputHandleIndex(1 byte)].This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2344");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2344
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetValidAttachmentsResponse.ReturnValue.GetType(),
                2344,
                @"[In RopGetValidAttachments ROP Failure Response Buffer]ReturnValue(4 bytes):An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2346");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2346
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetValidAttachmentsResponse.ReturnValue,
                2346,
                @"[In RopGetValidAttachments ROP Failure Response Buffer,ReturnValue(4 bytes)]For this response, this field is set to a value other than 0x00000000.");
        }
        #endregion

        #endregion

        #region Transport ROPs

        #region RopSubmitMessage
        /// <summary>
        /// Verify RopSubmitMessage Response
        /// </summary>
        /// <param name="ropSubmitMessageResponse">The response of RopSubmitMessage request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSubmitMessage request</param>
        private void VerifyRopSubmitMessageResponse(RopSubmitMessageResponse ropSubmitMessageResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2361");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2361
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSubmitMessageResponse.RopId.GetType(),
                2361,
                @"[In RopSubmitMessage ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2363");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2363
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSubmitMessage,
                ropSubmitMessageResponse.RopId,
                2363,
                @"[In RopSubmitMessage ROP Response Buffer,RopId (1 byte)]For this operation[RopSubmitMessage], this field is set to 0x32.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2364");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2364
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSubmitMessageResponse.InputHandleIndex.GetType(),
                2364,
                @"[In RopSubmitMessage ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2365");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2365
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSubmitMessageResponse.InputHandleIndex,
                2365,
                @"[In RopSubmitMessage ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2367");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2367
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSubmitMessageResponse.ReturnValue.GetType(),
                2367,
                @"[In RopSubmitMessage ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopAbortSubmit
        /// <summary>
        /// Verify RopAbortSubmit Response
        /// </summary>
        /// <param name="ropAbortSubmitResponse">The response of RopAbortSubmit request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopAbortSubmit request</param>
        private void VerifyRopAbortSubmitResponse(RopAbortSubmitResponse ropAbortSubmitResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2383");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2383
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortSubmitResponse.RopId.GetType(),
                2383,
                @"[In RopAbortSubmit ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2385");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2385
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopAbortSubmit,
                ropAbortSubmitResponse.RopId,
                2385,
                @"[In RopAbortSubmit ROP Response Buffer,RopId (1 byte)]For this operation[RopAbortSubmit], this field is set to 0x34.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2386");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2386
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropAbortSubmitResponse.InputHandleIndex.GetType(),
                2386,
                @"[In RopAbortSubmit ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2387");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2387
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropAbortSubmitResponse.InputHandleIndex,
                2387,
                @"[In RopAbortSubmit ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2389");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2389
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropAbortSubmitResponse.ReturnValue.GetType(),
                2389,
                @"[In RopAbortSubmit ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetAddressTypes
        /// <summary>
        /// Verify RopGetAddressTypes Success Response
        /// </summary>
        /// <param name="ropGetAddressTypeResponse">The response of RopGetAddressTypes request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetAddressTypes request</param>
        private void VerifyRopGetAddressTypesSuccessResponse(RopGetAddressTypesResponse ropGetAddressTypeResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2401");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2401
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAddressTypeResponse.RopId.GetType(),
                2401,
                @"[In RopGetAddressTypes ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2403");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2403
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetAddressTypes,
                ropGetAddressTypeResponse.RopId,
                2403,
                @"[In RopGetAddressTypes ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetAddressTypes], this field is set to 0x49.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2404");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2404
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAddressTypeResponse.InputHandleIndex.GetType(),
                2404,
                @"[In RopGetAddressTypes ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2405");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2405
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetAddressTypeResponse.InputHandleIndex,
                2405,
                @"[In RopGetAddressTypes ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2407");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2407
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetAddressTypeResponse.ReturnValue.GetType(),
                2407,
                @"[In RopGetAddressTypes ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2409");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2409
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetAddressTypeResponse.ReturnValue,
                2409,
                @"[In RopGetAddressTypes ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2410");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2410
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetAddressTypeResponse.AddressTypeCount.GetType(),
                2410,
                @"[In RopGetAddressTypes ROP Success Response Buffer]AddressTypeCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2412");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2412
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetAddressTypeResponse.AddressTypeSize.GetType(),
                2412,
                @"[In RopGetAddressTypes ROP Success Response Buffer]AddressTypeSize (2 bytes): An unsigned integer.");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2414
            // Assume the result is true. If one of List is not null-terminated ASCII string, then false.
            bool isVerifyR2414 = true;
            foreach (string s in ropGetAddressTypeResponse.AddressTypes)
            {
                byte[] buffer = Encoding.ASCII.GetBytes(s);
                isVerifyR2414 = isVerifyR2414 && this.IsNullTerminatedASCIIStr(buffer);

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2414, the actual value of AddressType is {0}", s.Trim('\0'));
                
                // If one of List is not null-terminated ASCII string then break iteration
                if (!isVerifyR2414)
                {
                    break;
                }
            }

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2414,
                2414,
                @"[In RopGetAddressTypes ROP Success Response Buffer]AddressTypes (variable): A list of null-terminated ASCII strings.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2415");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2415
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropGetAddressTypeResponse.AddressTypeCount,
                (ushort)ropGetAddressTypeResponse.AddressTypes.Count,
                2415,
                @"[In RopGetAddressTypes ROP Success Response Buffer,AddressTypes (variable)]The number of strings contained in this field[AddressTypes (variable)] is specified by the AddressTypeCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2416");

            // Assume the AddressTypeSize to zero, count the actual AddressType size in bytes
            int addressTypeSize = 0;
            foreach (string s in ropGetAddressTypeResponse.AddressTypes)
            {
                byte[] buffer = Encoding.ASCII.GetBytes(s);

                // Add each null-terminated ASCII string's count in bytes
                addressTypeSize += buffer.Length;
            }

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2416
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropGetAddressTypeResponse.AddressTypeSize,
                (ushort)addressTypeSize,
                2416,
                @"[In RopGetAddressTypes ROP Success Response Buffer,AddressTypes (variable)]The size of this field, in bytes, is specified by the AddressTypeSize field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2411");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2411
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetAddressTypeResponse.AddressTypes.Count,
                ropGetAddressTypeResponse.AddressTypeCount,
                2411,
                @"[In RopGetAddressTypes ROP Success Response Buffer,AddressTypeCount (2 bytes)]This value specifies the number of strings in the AddressTypes field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2413");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2413
            Site.CaptureRequirementIfAreEqual<ushort>(
                (ushort)addressTypeSize,
                ropGetAddressTypeResponse.AddressTypeSize,
                2413,
                @"[In RopGetAddressTypes ROP Success Response Buffer,AddressTypeSize (2 bytes)]This value specifies the length of the AddressTypes field.");
        }

        /// <summary>
        /// Verify RopGetAddressTypes Failure Response
        /// </summary>
        /// <param name="ropGetAddressTypeResponse">The response of RopGetAddressTypes request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetAddressTypes request</param>
        private void VerifyRopGetAddressTypesFailureResponse(RopGetAddressTypesResponse ropGetAddressTypeResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2417");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2417
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAddressTypeResponse.RopId.GetType(),
                2417,
                @"[In RopGetAddressTypes ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2419");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2419
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetAddressTypes,
                ropGetAddressTypeResponse.RopId,
                2419,
                @"[In RopGetAddressTypes ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetAddressTypes], this field is set to 0x49.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2420");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2420
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetAddressTypeResponse.InputHandleIndex.GetType(),
                2420,
                @"[In RopGetAddressTypes ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2421");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2421
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetAddressTypeResponse.InputHandleIndex,
                2421,
                @"[In RopGetAddressTypes ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2423");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2423
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetAddressTypeResponse.ReturnValue.GetType(),
                2423,
                @"[In RopGetAddressTypes ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Refer to MS-OXCROPS: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve 
            // the Server object and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4713");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4713
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropGetAddressTypeResponse.ReturnValue,
                    4713,
                    @"[In Appendix B: Product Behavior] For some ROPs, Implementation does use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid. (<23> Section 3.2.5.1: For some ROPs, Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.)");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2425");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2425
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropGetAddressTypeResponse.ReturnValue,
                    2425,
                    @"[In RopGetAddressTypes ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
            }
        }
        #endregion

        #region RopSetSpooler
        /// <summary>
        /// Verify RopSetSpooler Response
        /// </summary>
        /// <param name="ropSetSpoolerResponse">The response of RopSetSpooler request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetSpooler request</param>
        private void VerifyRopSetSpoolerResponse(RopSetSpoolerResponse ropSetSpoolerResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2436");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2436
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetSpoolerResponse.RopId.GetType(),
                2436,
                @"[In RopSetSpooler ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2438");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2438
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetSpooler,
                ropSetSpoolerResponse.RopId,
                2438,
                @"[In RopSetSpooler ROP Response Buffer,RopId (1 byte)]For this operation[RopSetSpooler], this field is set to 0x47.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2439");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2439
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetSpoolerResponse.InputHandleIndex.GetType(),
                2439,
                @"[In RopSetSpooler ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2440");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2440
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetSpoolerResponse.InputHandleIndex,
                2440,
                @"[In RopSetSpooler ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2442");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2442
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetSpoolerResponse.ReturnValue.GetType(),
                2442,
                @"[In RopSetSpooler ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSpoolerLockMessage
        /// <summary>
        /// Verify RopSpoolerLockMessage Response
        /// </summary>
        /// <param name="ropSpoolerLockMessageResponse">The response of RopSpoolerLockMessage request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSpoolerLockMessage request</param>
        private void VerifyRopSpoolerLockMessageResponse(RopSpoolerLockMessageResponse ropSpoolerLockMessageResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2458");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2458
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSpoolerLockMessageResponse.RopId.GetType(),
                2458,
                @"[In RopSpoolerLockMessage ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2460");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2460
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSpoolerLockMessage,
                ropSpoolerLockMessageResponse.RopId,
                2460,
                @"[In RopSpoolerLockMessage ROP Response Buffer,RopId (1 byte)]For this operation[RopSpoolerLockMessage], this field is set to 0x48.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2461");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2461
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSpoolerLockMessageResponse.InputHandleIndex.GetType(),
                2461,
                @"[In RopSpoolerLockMessage ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2462");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2462
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSpoolerLockMessageResponse.InputHandleIndex,
                2462,
                @"[In RopSpoolerLockMessage ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2464");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2464
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSpoolerLockMessageResponse.ReturnValue.GetType(),
                2464,
                @"[In RopSpoolerLockMessage ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopTransportSend
        /// <summary>
        /// Verify RopTransportSend Success Response
        /// </summary>
        /// <param name="ropTransportSendResponse">The success response of RopTransportSend request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopTransportSend request</param>
        private void VerifyRopTransportSendSuccessResponse(RopTransportSendResponse ropTransportSendResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2476");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2476
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportSendResponse.RopId.GetType(),
                2476,
                @"[In RopTransportSend ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2478");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2478
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopTransportSend,
                ropTransportSendResponse.RopId,
                2478,
                @"[In RopTransportSend ROP Success Response Buffer] RopId (1 byte):For this operation[RopTransportSend], this field is set to 0x4A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2479");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2479
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportSendResponse.InputHandleIndex.GetType(),
                2479,
                @"[In RopTransportSend ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2480");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2480
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropTransportSendResponse.InputHandleIndex,
                2480,
                @"[In RopTransportSend ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2482");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2482
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropTransportSendResponse.ReturnValue.GetType(),
                2482,
                @"[In RopTransportSend ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4686");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4686
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                 ropTransportSendResponse.ReturnValue,
                4686,
                @"[In RopTransportSend ROP Success Response Buffer,ReturnValue (4 bytes)] For this response, this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2484,NoPropertiesReturned:{0}", ropTransportSendResponse.NoPropertiesReturned);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2484
            // If NoPropertiesReturned is byte type and can be boolean type representation, then verify R2484
            bool isVerifyR2484 = (typeof(byte) == ropTransportSendResponse.NoPropertiesReturned.GetType()) &&
                                 ((Convert.ToBoolean(ropTransportSendResponse.NoPropertiesReturned) == true) ||
                                 (Convert.ToBoolean(ropTransportSendResponse.NoPropertiesReturned) == false));

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2484,
                2484,
                @"[In RopTransportSend ROP Success Response Buffer]NoPropertiesReturned (1 byte): An Boolean.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2486");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2486
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropTransportSendResponse.PropertyValueCount.GetType(),
                2486,
                @"[In  RopTransportSend ROP Success Response Buffer]PropertyValueCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2488");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2488
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TaggedPropertyValue[]),
                ropTransportSendResponse.PropertyValues.GetType(),
                2488,
                @"[In RopTransportSend ROP Success Response Buffer]PropertyValues (variable): An array of TaggedPropertyValue  structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2490");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2490
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropTransportSendResponse.PropertyValueCount,
                (ushort)ropTransportSendResponse.PropertyValues.Length,
                2490,
                @"[In RopTransportSend ROP Success Response Buffer,PropertyValues (variable)]The number of structures contained in this field is specified by the PropertyValueCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2485");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2485
            Site.CaptureRequirementIfIsTrue(
                (ropTransportSendResponse.NoPropertiesReturned == 0 && ropTransportSendResponse.PropertyValueCount > 0) || (ropTransportSendResponse.NoPropertiesReturned == 1 && ropTransportSendResponse.PropertyValueCount == 0),
                2485,
                @"[In RopTransportSend ROP Success Response Buffer,NoPropertiesReturned (1 byte)]This value specifies whether property values are returned.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2487");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2487
            Site.CaptureRequirementIfAreEqual<int>(
                ropTransportSendResponse.PropertyValueCount,
                ropTransportSendResponse.PropertyValues.Length,
                2487,
                @"[In RopTransportSend ROP Success Response Buffer,PropertyValueCount (2 bytes)]This value specifies the number of structures returned in the PropertyValues field.");

            foreach (TaggedPropertyValue taggedPropertyValue in ropTransportSendResponse.PropertyValues)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R468");

                // Verify MS-OXCROPS requirement: MS-OXCDATA_R468
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(PropertyTag),
                    taggedPropertyValue.PropertyTag.GetType(),
                    "MS-OXCDATA",
                    468,
                    @"[In TaggedPropertyValue Structure] PropertyTag (4 bytes): A PropertyTag structure, as specified in section 2.9, giving the values of the PropertyId and PropertyType fields for the property.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R181");

                // Verify MS-OXCROPS requirement: MS-OXCDATA_R181
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    taggedPropertyValue.PropertyTag.PropertyType.GetType(),
                    "MS-OXCDATA",
                    181,
                    @"[In PropertyTag Structure] PropertyType (2 bytes): An unsigned integer that identifies the data type of the property value, as specified by the table in section 2.11.1.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R182");

                // Verify MS-OXCROPS requirement: MS-OXCDATA_R182
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    taggedPropertyValue.PropertyTag.PropertyId.GetType(),
                    "MS-OXCDATA",
                    182,
                    @"[In PropertyTag Structure] PropertyId (2 bytes): An unsigned integer that identifies the property.");
            }
        }

        /// <summary>
        /// Verify RopTransportSend Failure Response
        /// </summary>
        /// <param name="ropTransportSendResponse">The failure response of RopTransportSend request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopTransportSend request</param>
        private void VerifyRopTransportSendFailureResponse(RopTransportSendResponse ropTransportSendResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4632");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4632
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportSendResponse.RopId.GetType(),
                4632,
                @"[In RopTransportSend ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4634");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4634
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopTransportSend,
                ropTransportSendResponse.RopId,
                4634,
                @"[In RopTransportSend ROP Failure Response Buffer]RopId (1 byte): For this operation[RopTransportSend], this field is set to 0x4A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4635");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4635
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportSendResponse.InputHandleIndex.GetType(),
                4635,
                @"[In RopTransportSend ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4636");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4636
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropTransportSendResponse.InputHandleIndex,
                4636,
                @"[In RopTransportSend ROP Failure Response Buffer]InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4637");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4637
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropTransportSendResponse.ReturnValue.GetType(),
                4637,
                @"[In RopTransportSend ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4639");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4639
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropTransportSendResponse.ReturnValue,
                4639,
                @"[In RopTransportSend ROP Failure Response Buffer]ReturnValue (4 bytes): For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopTransportNewMail
        /// <summary>
        /// Verify RopTransportNewMail Response
        /// </summary>
        /// <param name="ropTransportNewMailResponse">The response of RopTransportNewMail request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopTransportNewMail request</param>
        private void VerifyRopTransportNewMailResponse(RopTransportNewMailResponse ropTransportNewMailResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2511");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2511
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportNewMailResponse.RopId.GetType(),
                2511,
                @"[In RopTransportNewMail ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2513");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2513
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopTransportNewMail,
                ropTransportNewMailResponse.RopId,
                2513,
                @"[In RopTransportNewMail ROP Response Buffer,RopId (1 byte)]For this operation[[RopTransportNewMail]], this field is set to 0x51.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2514");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2514
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTransportNewMailResponse.InputHandleIndex.GetType(),
                2514,
                @"[In RopTransportNewMail ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2515");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2515
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropTransportNewMailResponse.InputHandleIndex,
                2515,
                @"[In RopTransportNewMail ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2517");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2517
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropTransportNewMailResponse.ReturnValue.GetType(),
                2517,
                @"[In RopTransportNewMail ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetTransportFolder
        /// <summary>
        /// Verify RopGetTransportFolder Success Response
        /// </summary>
        /// <param name="ropGetTransportFolderResponse">The response of RopGetTransportFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetTransportFolder request</param>
        private void VerifyRopGetTransportFolderSuccessResponse(RopGetTransportFolderResponse ropGetTransportFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2529");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2529
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetTransportFolderResponse.RopId.GetType(),
                2529,
                @"[In RopGetTransportFolder ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2531");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2531
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetTransportFolder,
                ropGetTransportFolderResponse.RopId,
                2531,
                @"[In RopGetTransportFolder ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetTransportFolder], this field is set to 0x6D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2532");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2532
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetTransportFolderResponse.InputHandleIndex.GetType(),
                2532,
                @"[In RopGetTransportFolder ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2533");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2533
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetTransportFolderResponse.InputHandleIndex,
                2533,
                @"[In RopGetTransportFolder ROP Success Response Buffer,InputHandleIndex (1 byte)]This field MUST be set to the InputHandleIndex specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2535");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2535
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetTransportFolderResponse.ReturnValue.GetType(),
                2535,
                @"[In RopGetTransportFolder ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2537");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2537
            Site.CaptureRequirementIfAreEqual<int>(
                8,
                Marshal.SizeOf(ropGetTransportFolderResponse.FolderId),
                2537,
                @"[In RopGetTransportFolder ROP Success Response Buffer]FolderId (8 bytes): 64-bit identifier.");
        }

        /// <summary>
        /// Verify RopGetTransportFolder Failure Response
        /// </summary>
        /// <param name="ropGetTransportFolderResponse">The response of RopGetTransportFolder request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetTransportFolder request</param>
        private void VerifyRopGetTransportFolderFailureResponse(RopGetTransportFolderResponse ropGetTransportFolderResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2539");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2539
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetTransportFolderResponse.RopId.GetType(),
                2539,
                @"[In RopGetTransportFolder ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2541");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2541
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetTransportFolder,
                ropGetTransportFolderResponse.RopId,
                2541,
                @"[In RopGetTransportFolder ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetTransportFolder], this field is set to 0x6D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2542");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2542
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetTransportFolderResponse.InputHandleIndex.GetType(),
                2542,
                @"[In RopGetTransportFolder ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2543");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2543
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetTransportFolderResponse.InputHandleIndex,
                2543,
                @"[In RopGetTransportFolder ROP Failure Response Buffer,InputHandleIndex (1 byte)]This field MUST be set to the InputHandleIndex specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2545");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2545
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetTransportFolderResponse.ReturnValue.GetType(),
                2545,
                @"[In RopGetTransportFolder ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Refer to MS-OXCROPS: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve 
            // the Server object and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4713");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4713
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropGetTransportFolderResponse.ReturnValue,
                    4713,
                    @"[In Appendix B: Product Behavior] For some ROPs, Implementation does use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid. (<23> Section 3.2.5.1: For some ROPs, Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.)");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2547");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2547
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropGetTransportFolderResponse.ReturnValue,
                    2547,
                    @"[In RopGetTransportFolder ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
            }
        }
        #endregion

        #region RopOptionsData
        /// <summary>
        /// Verify RopOptionsData Success Response
        /// </summary>
        /// <param name="ropOptionsDataResponse">The success response of RopOptionsData request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopOptionsData request</param>
        private void VerifyRopOptionsDataSuccessResponse(RopOptionsDataResponse ropOptionsDataResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2562");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2562
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOptionsDataResponse.RopId.GetType(),
                2562,
                @"[In RopOptionsData ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2564");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2564
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOptionsData,
                ropOptionsDataResponse.RopId,
                2564,
                @"[In RopOptionsData ROP Success Response Buffer,RopId (1 byte)]For this operation[RopOptionsData], this field is set to 0x6F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2565");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2565
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOptionsDataResponse.InputHandleIndex.GetType(),
                2565,
                @"[In RopOptionsData ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2566");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2566
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropOptionsDataResponse.InputHandleIndex,
                2566,
                @"[In RopOptionsData ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2568");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2568
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOptionsDataResponse.ReturnValue.GetType(),
                2568,
                @"[In RopOptionsData ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4687");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4687
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropOptionsDataResponse.ReturnValue,
                4687,
                @"[In RopOptionsData ROP Success Response Buffer]ReturnValue (4 bytes):For this response, this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2570");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2570
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOptionsDataResponse.Reserved.GetType(),
                2570,
                @"[In  RopOptionsData ROP Success Response Buffer]Reserved (1 byte): Reserved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2571");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2571
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)ResponseFlags.Reserved,
                ropOptionsDataResponse.Reserved,
                2571,
                @"[In RopOptionsData ROP Success Response Buffer,Reserved (1 byte)]This field MUST be set to 0x01.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2572");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2572
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOptionsDataResponse.OptionsInfoSize.GetType(),
                2572,
                @"[In RopOptionsData ROP Success Response Buffer]OptionsInfoSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2574");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2574
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropOptionsDataResponse.OptionsInfo.GetType(),
                2574,
                @"[In RopOptionsData ROP Success Response Buffer]OptionsInfo (variable): An array of bytes.");

            if (Common.IsRequirementEnabled(2578, this.Site))
            {
                bool isVerifyR2578 = true;

                if (ropOptionsDataResponse.OptionsInfo.Length != 0)
                {
                    foreach (byte item in ropOptionsDataResponse.OptionsInfo)
                    {
                        if (item != 0)
                        {
                            isVerifyR2578 = false;
                            break;
                        }
                    }
                }

                // Add the comment information for debugging
                Site.Log.Add(LogEntryKind.Comment, "Verify MS-OXCROPS_R2578,OptionsInfo.Length:{0}", ropOptionsDataResponse.OptionsInfo.Length);

                // Verify MS-OXCROPS requirement: 2578
                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2578,
                    2578,
                    @"[In Appendix B: Product Behavior] Implementation does return this field[OptionsInfo] as an empty array. (Microsoft Exchange Server 2013 follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(2928, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2928, the actual length of the OptionsInfo is {0}", ropOptionsDataResponse.OptionsInfo.Length);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2928
                bool isVerifyR2928 = ropOptionsDataResponse.OptionsInfo.Length > 0;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR2928,
                    2928,
                    @"[In Appendix B: Product Behavior] Implementation does not return [OptionsInfo (variable) field as] an empty array in the OptionsInfo field. (<7> Section 2.2.7.9.2: Exchange 2007, and Exchange 2010 do not return an empty array in the OptionsInfo field.)");
            }

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2580");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2580
            // SUT is exchange 2007 and exchange 2010 if program go through here, OptionsInfo does not return an empty array
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropOptionsDataResponse.OptionsInfoSize,
                (ushort)ropOptionsDataResponse.OptionsInfo.Length,
                2580,
                @"[In RopOptionsData ROP Success Response Buffer,OptionsInfo (variable)]The size of this field, in bytes, is specified by the OptionsInfoSize field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2581");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2581
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropOptionsDataResponse.HelpFileSize.GetType(),
                2581,
                @"[In RopOptionsData ROP Success Response Buffer]HelpFileSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2583");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2583
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropOptionsDataResponse.HelpFile.GetType(),
                2583,
                @"[In RopOptionsData ROP Success Response Buffer]HelpFile (variable): An array of bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2585");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2585
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropOptionsDataResponse.HelpFileSize,
                (ushort)ropOptionsDataResponse.HelpFile.Length,
                2585,
                @"[In RopOptionsData ROP Success Response Buffer,HelpFile (variable)]The size of this field, in bytes, is specified by the HelpFileSize field.");
        }

        /// <summary>
        /// Verify RopOptionsData Failure Response
        /// </summary>
        /// <param name="ropOptionsDataResponse">The failure response of RopOptionsData request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopOptionsData request</param>
        private void VerifyRopOptionsDataFailureResponse(RopOptionsDataResponse ropOptionsDataResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4640");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4640
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOptionsDataResponse.RopId.GetType(),
                4640,
                @"[In RopOptionsData ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4642");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4642
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOptionsData,
                ropOptionsDataResponse.RopId,
                4642,
                @"[In RopOptionsData ROP Failure Response Buffer]RopId (1 byte): For this operation[RopOptionsData], this field is set to 0x6F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4643");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4643
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOptionsDataResponse.InputHandleIndex.GetType(),
                4643,
                @"[In RopOptionsData ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4644");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4644
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropOptionsDataResponse.InputHandleIndex,
                4644,
                @"[In RopOptionsData ROP Failure Response Buffer]InputHandleIndex (1 byte): This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4645");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4645
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOptionsDataResponse.ReturnValue.GetType(),
                4645,
                @"[In RopOptionsData ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            if (Common.IsRequirementEnabled(4691, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4691,the actual value of the returnValue is {0}", ropOptionsDataResponse.ReturnValue);

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4691
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropOptionsDataResponse.ReturnValue,
                    4691,
                    @"[In Appendix B: Product Behavior] ReturnValue (4 bytes): For this response[Failure Response], Implementation does set this field to a value other than 0x00000000. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(4690, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4690");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4690
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropOptionsDataResponse.ReturnValue,
                    4690,
                    @"[In Appendix B: Product Behavior] Implementation does set the ReturnValue field for the RopOptionsData ROP response (section 2.2.7.9) to 0x00000000 regardless of the failure of the ROP. (<8> Section 2.2.7.9.3: Exchange 2007 sets the ReturnValue field for the RopOptionsData ROP response (section 2.2.7.9) to 0x00000000 regardless of the failure of the ROP.)");
            }
        }
        #endregion

        #endregion

        #region Property ROPs

        #region RopGetPropertyIdsFromNames
        /// <summary>
        /// Verify RopGetPropertyIdsFromNames Success Response
        /// </summary>
        /// <param name="ropGetPropertyIdsFromNamesResponse">The success response of RopGetPropertyIdsFromNames request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertyIdsFromNames request</param>
        private void VerifyRopGetPropertyIdsFromNamesSuccessResponse(RopGetPropertyIdsFromNamesResponse ropGetPropertyIdsFromNamesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2609");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2609
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertyIdsFromNamesResponse.RopId.GetType(),
                2609,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2611");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2611
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertyIdsFromNames,
                ropGetPropertyIdsFromNamesResponse.RopId,
                2611,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetPropertyIdsFromNames], this field is set to 0x56.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2612");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2612
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertyIdsFromNamesResponse.InputHandleIndex.GetType(),
                2612,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2613");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2613
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertyIdsFromNamesResponse.InputHandleIndex,
                2613,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer, InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2615");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2615
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertyIdsFromNamesResponse.ReturnValue.GetType(),
                2615,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2617");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2617
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPropertyIdsFromNamesResponse.ReturnValue,
                2617,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2618");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2618
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetPropertyIdsFromNamesResponse.PropertyIdCount.GetType(),
                2618,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer]PropertyIdCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2620");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2620
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyId[]),
                ropGetPropertyIdsFromNamesResponse.PropertyIds.GetType(),
                2620,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer]PropertyIds (variable): An array of unsigned 16-bit integers.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2621");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2621
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetPropertyIdsFromNamesResponse.PropertyIdCount,
                ropGetPropertyIdsFromNamesResponse.PropertyIds.Length,
                2621,
                @"[In RopGetPropertyIdsFromNames ROP Success Response Buffer,PropertyIds (variable)]The number of structures contained in this field is specified by the PropertyIdCount field.");
        }

        /// <summary>
        /// Verify RopGetPropertyIdsFromNames Failure Response
        /// </summary>
        /// <param name="ropGetPropertyIdsFromNamesResponse">The failure response of RopGetPropertyIdsFromNames request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertyIdsFromNames request</param>
        private void VerifyRopGetPropertyIdsFromNamesFailureResponse(RopGetPropertyIdsFromNamesResponse ropGetPropertyIdsFromNamesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2624");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2624
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertyIdsFromNamesResponse.RopId.GetType(),
                2624,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2626");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2626
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertyIdsFromNames,
                ropGetPropertyIdsFromNamesResponse.RopId,
                2626,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetPropertyIdsFromNames], this field is set to 0x56.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2627");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2627
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertyIdsFromNamesResponse.InputHandleIndex.GetType(),
                2627,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2628");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2628
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertyIdsFromNamesResponse.InputHandleIndex,
                2628,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2630");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2630
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertyIdsFromNamesResponse.ReturnValue.GetType(),
                2630,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2632");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2632
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPropertyIdsFromNamesResponse.ReturnValue,
                2632,
                @"[In RopGetPropertyIdsFromNames ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopGetNamesFromPropertyIds

        /// <summary>
        /// Verify RopGetNamesFromPropertyIds Success Response
        /// </summary>
        /// <param name="ropGetNamesFromPropertyIdsResponse">The success response of RopGetNamesFromPropertyIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetNamesFromPropertyIds request</param>
        private void VerifyRopGetNamesFromPropertyIdsSuccessResponse(RopGetNamesFromPropertyIdsResponse ropGetNamesFromPropertyIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2649");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2649
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetNamesFromPropertyIdsResponse.RopId.GetType(),
                2649,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2651");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2651
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetNamesFromPropertyIds,
                ropGetNamesFromPropertyIdsResponse.RopId,
                2651,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer,RopId (1 byte)] For this operation[RopGetNamesFromPropertyIds], this field is set to 0x55.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2652");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2652
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetNamesFromPropertyIdsResponse.InputHandleIndex.GetType(),
                2652,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2653");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2653
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetNamesFromPropertyIdsResponse.InputHandleIndex,
                2653,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the InputHandleIndex specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2655");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2655
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetNamesFromPropertyIdsResponse.ReturnValue.GetType(),
                2655,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2657");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2657
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetNamesFromPropertyIdsResponse.ReturnValue,
                2657,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2658");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2658
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetNamesFromPropertyIdsResponse.PropertyNameCount.GetType(),
                2658,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer]PropertyNameCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2660");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2660
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyName[]),
                ropGetNamesFromPropertyIdsResponse.PropertyNames.GetType(),
                2660,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer]PropertyNames (variable): A list of PropertyName structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2661");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2661
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetNamesFromPropertyIdsResponse.PropertyNameCount,
                ropGetNamesFromPropertyIdsResponse.PropertyNames.Length,
                2661,
                @"[In RopGetNamesFromPropertyIds ROP Success Response Buffer,PropertyNames (variable)]The number of structures contained in this field is specified by the PropertyNameCount field.");
        }

        /// <summary>
        /// Verify RopGetNamesFromPropertyIds Failure Response
        /// </summary>
        /// <param name="ropGetNamesFromPropertyIdsResponse">The failure response of RopGetNamesFromPropertyIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetNamesFromPropertyIds request</param>
        private void VerifyRopGetNamesFromPropertyIdsFailureResponse(RopGetNamesFromPropertyIdsResponse ropGetNamesFromPropertyIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2664");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2664
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetNamesFromPropertyIdsResponse.RopId.GetType(),
                2664,
                @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2666");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2666
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetNamesFromPropertyIds,
                ropGetNamesFromPropertyIdsResponse.RopId,
                2666,
                @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetNamesFromPropertyIds], this field is set to 0x55.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2667");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2667
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetNamesFromPropertyIdsResponse.InputHandleIndex.GetType(),
                2667,
                @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2668");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2668
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetNamesFromPropertyIdsResponse.InputHandleIndex,
                2668,
                @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2670");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2670
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetNamesFromPropertyIdsResponse.ReturnValue.GetType(),
                2670,
                @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Refer to MS-OXCROPS: For some ROPs, Exchange 2003 and Exchange 2007 use different methods to resolve 
            // the Server object and, therefore, do not fail the ROP if the index is invalid.
            if (Common.IsRequirementEnabled(4713, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4713");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4713
                Site.CaptureRequirementIfAreEqual<uint>(
                    SuccessReturnValue,
                    ropGetNamesFromPropertyIdsResponse.ReturnValue,
                    4713,
                    @"[In Appendix B: Product Behavior] For some ROPs, Implementation does use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid. (<23> Section 3.2.5.1: For some ROPs, Exchange 2007 use different methods to resolve the Server object and, therefore, do not fail the ROP if the index is invalid.)");
            }
            else
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2672");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2672
                Site.CaptureRequirementIfAreNotEqual<uint>(
                    SuccessReturnValue,
                    ropGetNamesFromPropertyIdsResponse.ReturnValue,
                    2672,
                    @"[In RopGetNamesFromPropertyIds ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
            }
        }

        #endregion

        #region RopGetPropertiesSpecific

        /// <summary>
        /// Verify RopGetPropertiesSpecific Success Response
        /// </summary>
        /// <param name="ropGetPropertiesSpecificResponse">The success response of RopGetPropertiesSpecific request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertiesSpecific request</param>
        /// <param name="propertyTags">The field of PropertyTags in RopGetPropertiesSpecific request</param>
        private void VerifyRopGetPropertiesSpecificSuccessResponse(RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse, byte inputHandleIndex, PropertyTag[] propertyTags)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2693");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2693
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesSpecificResponse.RopId.GetType(),
                2693,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2695");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2695
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesSpecific,
                ropGetPropertiesSpecificResponse.RopId,
                2695,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesSpecific], this field is set to 0x07.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2696");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2696
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesSpecificResponse.InputHandleIndex.GetType(),
                2696,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2697");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2697
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesSpecificResponse.InputHandleIndex,
                2697,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2699");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2699
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesSpecificResponse.ReturnValue.GetType(),
                2699,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2701");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2701
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesSpecificResponse.ReturnValue,
                2701,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2702");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2702
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyRow),
                ropGetPropertiesSpecificResponse.RowData.GetType(),
                2702,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer]RowData (variable): A PropertyRow structure.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2704");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2704
            Site.CaptureRequirementIfAreEqual<int>(
                propertyTags.Length,
                ropGetPropertiesSpecificResponse.RowData.PropertyValues.Count,
                2704,
                @"[In RopGetPropertiesSpecific ROP Success Response Buffer,RowData (variable)]the columns used for these rows[RowData (variable)] were those specified in the PropertyTags field in the ROP request.");
        }

        /// <summary>
        /// Verify RopGetPropertiesSpecific Failure Response
        /// </summary>
        /// <param name="ropGetPropertiesSpecificResponse">The failure response of RopGetPropertiesSpecific request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertiesSpecific request</param>
        private void VerifyRopGetPropertiesSpecificFailureResponse(RopGetPropertiesSpecificResponse ropGetPropertiesSpecificResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2705");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2705
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesSpecificResponse.RopId.GetType(),
                2705,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2707");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2707
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesSpecific,
                ropGetPropertiesSpecificResponse.RopId,
                2707,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesSpecific], this field is set to 0x07.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2708");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2708
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesSpecificResponse.InputHandleIndex.GetType(),
                2708,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2709");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2709
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesSpecificResponse.InputHandleIndex,
                2709,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2711");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2711
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesSpecificResponse.ReturnValue.GetType(),
                2711,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2713");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2713
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesSpecificResponse.ReturnValue,
                2713,
                @"[In RopGetPropertiesSpecific ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopGetPropertiesAll

        /// <summary>
        /// Verify RopGetPropertiesAll Success Response
        /// </summary>
        /// <param name="ropGetPropertiesAllResponse">The success response of RopGetPropertiesAll request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertiesAll request</param>
        private void VerifyRopGetPropertiesAllSuccessResponse(RopGetPropertiesAllResponse ropGetPropertiesAllResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2728");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2728
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesAllResponse.RopId.GetType(),
                2728,
                @"[In RopGetPropertiesAll ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2730");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2730
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesAll,
                ropGetPropertiesAllResponse.RopId,
                2730,
                @"[In RopGetPropertiesAll ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesAll], this field is set to 0x08.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2731");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2731
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesAllResponse.InputHandleIndex.GetType(),
                2731,
                @"[In RopGetPropertiesAll ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2732");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2732
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesAllResponse.InputHandleIndex,
                2732,
                @"[In RopGetPropertiesAll ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2734");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2734
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesAllResponse.ReturnValue.GetType(),
                2734,
                @"[In RopGetPropertiesAll ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2736");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2736
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesAllResponse.ReturnValue,
                2736,
                @"[In RopGetPropertiesAll ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2737");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2737
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetPropertiesAllResponse.PropertyValueCount.GetType(),
                2737,
                @"[In RopGetPropertiesAll ROP Success Response Buffer]PropertyValueCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2739");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2739
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(TaggedPropertyValue[]),
                ropGetPropertiesAllResponse.PropertyValues.GetType(),
                2739,
                @"[In RopGetPropertiesAll ROP Success Response Buffer]PropertyValues (variable): An array of TaggedPropertyValue structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2740");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2740
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetPropertiesAllResponse.PropertyValueCount,
                ropGetPropertiesAllResponse.PropertyValues.Length,
                2740,
                @"[In RopGetPropertiesAll ROP Success Response Buffer,PropertyValues (variable)]The number of structures contained in this field is specified by the PropertyValueCount field.");
        }

        /// <summary>
        /// Verify RopGetPropertiesAll Failure Response
        /// </summary>
        /// <param name="ropGetPropertiesAllResponse">The failure response of RopGetPropertiesAll request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertiesAll request</param>
        private void VerifyRopGetPropertiesAllFailureResponse(RopGetPropertiesAllResponse ropGetPropertiesAllResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2743");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2743
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesAllResponse.RopId.GetType(),
                2743,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2745");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2745
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesAll,
                ropGetPropertiesAllResponse.RopId,
                2745,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesAll], this field is set to 0x08.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2746");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2746
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesAllResponse.InputHandleIndex.GetType(),
                2746,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2747");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2747
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesAllResponse.InputHandleIndex,
                2747,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2749");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2749
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesAllResponse.ReturnValue.GetType(),
                2749,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2751");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2751
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesAllResponse.ReturnValue,
                2751,
                @"[In RopGetPropertiesAll ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopGetPropertiesList

        /// <summary>
        /// Verify RopGetPropertiesList Success Response
        /// </summary>
        /// <param name="ropGetPropertiesListResponse">The success response of RopGetPropertiesList request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetPropertiesList request</param>
        private void VerifyRopGetPropertiesListSuccessResponse(RopGetPropertiesListResponse ropGetPropertiesListResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2762");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2762
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesListResponse.RopId.GetType(),
                2762,
                @"[In RopGetPropertiesList ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2764");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2764
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesList,
                ropGetPropertiesListResponse.RopId,
                2764,
                @"[In RopGetPropertiesList ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesList], this field is set to 0x09.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2765");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2765
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesListResponse.InputHandleIndex.GetType(),
                2765,
                @"[In RopGetPropertiesList ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2766");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2766
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesListResponse.InputHandleIndex,
                2766,
                @"[In RopGetPropertiesList ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2768");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2768
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesListResponse.ReturnValue.GetType(),
                2768,
                @"[In RopGetPropertiesList ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2770");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2770
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesListResponse.ReturnValue,
                2770,
                @"[In RopGetPropertiesList ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2771");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2771
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropGetPropertiesListResponse.PropertyTagCount.GetType(),
                2771,
                @"[In RopGetPropertiesList ROP Success Response Buffer]PropertyTagCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2773");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2773
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyTag[]),
                ropGetPropertiesListResponse.PropertyTags.GetType(),
                2773,
                @"[In RopGetPropertiesList ROP Success Response Buffer]PropertyTags (variable): An array of PropertyTag structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2774");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2774
            Site.CaptureRequirementIfAreEqual<int>(
                ropGetPropertiesListResponse.PropertyTagCount,
                ropGetPropertiesListResponse.PropertyTags.Length,
                2774,
                @"[In RopGetPropertiesList ROP Success Response Buffer,PropertyTags (variable)]The number of structures contained in this field is specified by the PropertyTagCount field.");
        }

        /// <summary>
        /// Verify RopGetPropertiesList Failure Response
        /// </summary>
        /// <param name="ropGetPropertiesListResponse">The failure response of RopGetPropertiesList request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in ROP RopGetPropertiesList request</param>
        private void VerifyRopGetPropertiesListFailureResponse(RopGetPropertiesListResponse ropGetPropertiesListResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2777");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2777
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesListResponse.RopId.GetType(),
                2777,
                @"[In RopGetPropertiesList ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2779");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2779
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPropertiesList,
                ropGetPropertiesListResponse.RopId,
                2779,
                @"[In RopGetPropertiesList ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetPropertiesList], this field is set to 0x09.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2780");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2780
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPropertiesListResponse.InputHandleIndex.GetType(),
                2780,
                @"[In RopGetPropertiesList ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2781");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2781
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetPropertiesListResponse.InputHandleIndex,
                2781,
                @"[In RopGetPropertiesList ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2783");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2783
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPropertiesListResponse.ReturnValue.GetType(),
                2783,
                @"[In RopGetPropertiesList ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2785");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2785
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetPropertiesListResponse.ReturnValue,
                2785,
                @"[In RopGetPropertiesList ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopSetProperties

        /// <summary>
        /// Verify RopSetProperties Success Response
        /// </summary>
        /// <param name="ropSetPropertiesResponse">The success response of RopSetProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetProperties request</param>
        private void VerifyRopSetPropertiesSuccessResponse(RopSetPropertiesResponse ropSetPropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2805");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2805
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesResponse.RopId.GetType(),
                2805,
                @"[In RopSetProperties ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2807");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2807
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetProperties,
                ropSetPropertiesResponse.RopId,
                2807,
                @"[In RopSetProperties ROP Success Response Buffer,RopId (1 byte)]For this operation[ RopSetProperties], this field is set to 0x0A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2808");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2808
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesResponse.InputHandleIndex.GetType(),
                2808,
                @"[In RopSetProperties ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2809");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2809
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetPropertiesResponse.InputHandleIndex,
                2809,
                @"[In RopSetProperties ROP Success Response Buffer,InputHandleIndex (1 byte)]
This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2811");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2811
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetPropertiesResponse.ReturnValue.GetType(),
                2811,
                @"[In RopSetProperties ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2813");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2813
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetPropertiesResponse.ReturnValue,
                2813,
                @"[In RopSetProperties ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2814");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2814
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropSetPropertiesResponse.PropertyProblemCount.GetType(),
                2814,
                @"[In RopSetProperties ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2816");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2816
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyProblem[]),
                ropSetPropertiesResponse.PropertyProblems.GetType(),
                2816,
                @"[In RopSetProperties ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2817");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2817
            Site.CaptureRequirementIfAreEqual<int>(
                ropSetPropertiesResponse.PropertyProblemCount,
                ropSetPropertiesResponse.PropertyProblems.Length,
                2817,
                @"[In RopSetProperties ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field is specified by the PropertyProblemCount field.");
        }

        /// <summary>
        /// Verify RopSetProperties Failure Response
        /// </summary>
        /// <param name="ropSetPropertiesResponse">The failure response of RopSetProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetProperties request</param>
        private void VerifyRopSetPropertiesFailureResponse(RopSetPropertiesResponse ropSetPropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2819");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2819
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesResponse.RopId.GetType(),
                2819,
                @"[In RopSetProperties ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2821");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2821
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetProperties,
                ropSetPropertiesResponse.RopId,
                2821,
                @"[In RopSetProperties ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopSetProperties], this field is set to 0x0A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2822");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2822
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesResponse.InputHandleIndex.GetType(),
                2822,
                @"[In RopSetProperties ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2823");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2823
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetPropertiesResponse.InputHandleIndex,
                2823,
                @"[In RopSetProperties ROP Failure Response Buffer,InputHandleIndex (1 byte)]
This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2825");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2825
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetPropertiesResponse.ReturnValue.GetType(),
                2825,
                @"[In RopSetProperties ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2827");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2827
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetPropertiesResponse.ReturnValue,
                2827,
                @"[In RopSetProperties ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopSetPropertiesNoReplicate

        /// <summary>
        /// Verify RopSetPropertiesNoReplicate Success Response
        /// </summary>
        /// <param name="ropSetPropertiesNoReplicateResponse">The success response of RopSetPropertiesNoReplicate request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetPropertiesNoReplicate request</param>
        private void VerifyRopSetPropertiesNoReplicateSuccessResponse(RopSetPropertiesNoReplicateResponse ropSetPropertiesNoReplicateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2847");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2847
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesNoReplicateResponse.RopId.GetType(),
                2847,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2849");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2849
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetPropertiesNoReplicate,
                ropSetPropertiesNoReplicateResponse.RopId,
                2849,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer,RopId (1 byte)]For this operation[RopSetPropertiesNoReplicate], this field is set to 0x79.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2850");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2850
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesNoReplicateResponse.InputHandleIndex.GetType(),
                2850,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2851");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2851
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetPropertiesNoReplicateResponse.InputHandleIndex,
                2851,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2853");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2853
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetPropertiesNoReplicateResponse.ReturnValue.GetType(),
                2853,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2855");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2855
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSetPropertiesNoReplicateResponse.ReturnValue,
                2855,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2856");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2856
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropSetPropertiesNoReplicateResponse.PropertyProblemCount.GetType(),
                2856,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2858");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2858
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyProblem[]),
                ropSetPropertiesNoReplicateResponse.PropertyProblems.GetType(),
                2858,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2859");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2859
            Site.CaptureRequirementIfAreEqual<int>(
                ropSetPropertiesNoReplicateResponse.PropertyProblemCount,
                ropSetPropertiesNoReplicateResponse.PropertyProblems.Length,
                2859,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field is specified by the PropertyProblemCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2857");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2857
            Site.CaptureRequirementIfAreEqual<int>(
                ropSetPropertiesNoReplicateResponse.PropertyProblems.Length,
                ropSetPropertiesNoReplicateResponse.PropertyProblemCount,
                2857,
                @"[In RopSetPropertiesNoReplicate ROP Success Response Buffer,PropertyProblemCount (2 bytes)]This value specifies the number of PropertyProblem structures in the PropertyProblems field.");
        }

        /// <summary>
        /// Verify RopSetPropertiesNoReplicate Failure Response
        /// </summary>
        /// <param name="ropSetPropertiesNoReplicateResponse">The failure response of SetPropertiesNoReplicate request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetPropertiesNoReplicate request</param>
        private void VerifyRopSetPropertiesNoReplicateFailureResponse(RopSetPropertiesNoReplicateResponse ropSetPropertiesNoReplicateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2861");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2861
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesNoReplicateResponse.RopId.GetType(),
                2861,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2863");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2863
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetPropertiesNoReplicate,
                ropSetPropertiesNoReplicateResponse.RopId,
                2863,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopSetPropertiesNoReplicate], this field is set to 0x79.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2864");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2864
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetPropertiesNoReplicateResponse.InputHandleIndex.GetType(),
                2864,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2865");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2865
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetPropertiesNoReplicateResponse.InputHandleIndex,
                2865,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2867");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2867
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetPropertiesNoReplicateResponse.ReturnValue.GetType(),
                2867,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2869");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2869
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSetPropertiesNoReplicateResponse.ReturnValue,
                2869,
                @"[In RopSetPropertiesNoReplicate ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopDeleteProperties

        /// <summary>
        /// Verify RopDeleteProperties Success Response
        /// </summary>
        /// <param name="ropDeletePropertiesResponse">The success response of RopDeleteProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeleteProperties request</param>
        private void VerifyRopDeletePropertiesSuccessResponse(RopDeletePropertiesResponse ropDeletePropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2886");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2886
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesResponse.RopId.GetType(),
                2886,
                @"[In RopDeleteProperties ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2888");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2888
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeleteProperties,
                ropDeletePropertiesResponse.RopId,
                2888,
                @"[In RopDeleteProperties ROP Success Response Buffer,RopId (1 byte)]For this operation[RopDeleteProperties], this field is set to 0x0B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2889");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2889
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesResponse.InputHandleIndex.GetType(),
                2889,
                @"[In RopDeleteProperties ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2890");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2890
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeletePropertiesResponse.InputHandleIndex,
                2890,
                @"[In RopDeleteProperties ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2892");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2892
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeletePropertiesResponse.ReturnValue.GetType(),
                2892,
                @"[In RopDeleteProperties ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2894");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2894
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropDeletePropertiesResponse.ReturnValue,
                2894,
                @"[In RopDeleteProperties ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2895");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2895
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropDeletePropertiesResponse.PropertyProblemCount.GetType(),
                2895,
                @"[In RopDeleteProperties ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // According to the de-serialize approach of ropDeletePropertiesResponse,if the PropertyProblemCount's value is greater than 0, the error operation occurs.
            if (ropDeletePropertiesResponse.PropertyProblemCount > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2897");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2897
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(PropertyProblem[]),
                    ropDeletePropertiesResponse.PropertyProblems.GetType(),
                    2897,
                    @"[In RopDeleteProperties ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2898");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2898
                Site.CaptureRequirementIfAreEqual<int>(
                    ropDeletePropertiesResponse.PropertyProblemCount,
                    ropDeletePropertiesResponse.PropertyProblems.Length,
                    2898,
                    @"[In RopDeleteProperties ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field is specified by the PropertyProblemCount field.");
            }
        }

        /// <summary>
        /// Verify RopDeleteProperties Failure Response
        /// </summary>
        /// <param name="ropDeletePropertiesResponse">The failure response of RopDeleteProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeleteProperties request</param>
        private void VerifyRopDeletePropertiesFailureResponse(RopDeletePropertiesResponse ropDeletePropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2900");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2900
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesResponse.RopId.GetType(),
                2900,
                @"[In RopDeleteProperties ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2902");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2902
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeleteProperties,
                ropDeletePropertiesResponse.RopId,
                2902,
                @"[In RopDeleteProperties ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopDeleteProperties], this field is set to 0x0B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2903");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2903
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesResponse.InputHandleIndex.GetType(),
                2903,
                @"[In RopDeleteProperties ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2904");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2904
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeletePropertiesResponse.InputHandleIndex,
                2904,
                @"[In RopDeleteProperties ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2906");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2906
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeletePropertiesResponse.ReturnValue.GetType(),
                2906,
                @"[In RopDeleteProperties ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2908");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2908
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropDeletePropertiesResponse.ReturnValue,
                2908,
                @"[In RopDeleteProperties ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopDeletePropertiesNoReplicate

        /// <summary>
        /// Verify RopDeletePropertiesNoReplicate Success Response
        /// </summary>
        /// <param name="ropDeletePropertiesNoReplicateResponse">The success response of RopDeletePropertiesNoReplicate request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeletePropertiesNoReplicate request</param>
        private void VerifyRopDeletePropertiesNoReplicateSuccessResponse(RopDeletePropertiesNoReplicateResponse ropDeletePropertiesNoReplicateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2930");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2930
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesNoReplicateResponse.RopId.GetType(),
                2930,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2932");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2932
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeletePropertiesNoReplicate,
                ropDeletePropertiesNoReplicateResponse.RopId,
                2932,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer,RopId (1 byte)]For this operation[RopDeletePropertiesNoReplicate], this field is set to 0x7A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2933");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2933
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesNoReplicateResponse.InputHandleIndex.GetType(),
                2933,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2934");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2934
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeletePropertiesNoReplicateResponse.InputHandleIndex,
                2934,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2936");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2936
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeletePropertiesNoReplicateResponse.ReturnValue.GetType(),
                2936,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2938");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2938
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropDeletePropertiesNoReplicateResponse.ReturnValue,
                2938,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2939");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2939
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropDeletePropertiesNoReplicateResponse.PropertyProblemCount.GetType(),
                2939,
                @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // According to the de-serialize approach of ropDeletePropertiesNoReplicateResponse,if the PropertyProblemCount's value is greater than 0, the error operation occurs.
            if (ropDeletePropertiesNoReplicateResponse.PropertyProblemCount > 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2941");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2941
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(PropertyProblem[]),
                    ropDeletePropertiesNoReplicateResponse.PropertyProblems.GetType(),
                    2941,
                    @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2942");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2942
                Site.CaptureRequirementIfAreEqual<int>(
                    ropDeletePropertiesNoReplicateResponse.PropertyProblemCount,
                    ropDeletePropertiesNoReplicateResponse.PropertyProblems.Length,
                    2942,
                    @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field is specified by the PropertyProblemCount field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2940");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R2940
                Site.CaptureRequirementIfAreEqual<int>(
                    ropDeletePropertiesNoReplicateResponse.PropertyProblems.Length,
                    ropDeletePropertiesNoReplicateResponse.PropertyProblemCount,
                    2940,
                    @"[In RopDeletePropertiesNoReplicate ROP Success Response Buffer,PropertyProblemCount (2 bytes)]This value specifies the number of PropertyProblem structures in the PropertyProblems field.");

                foreach (PropertyProblem propertyProblem in ropDeletePropertiesNoReplicateResponse.PropertyProblems)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R46");

                    // Verify MS-OXCROPS requirement: MS-OXCDATA_R46
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(ushort),
                        propertyProblem.Index.GetType(),
                        "MS-OXCDATA",
                        46,
                        @"[In PropertyProblem Structure] Index (2 bytes): An unsigned integer.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R48");

                    // Verify MS-OXCROPS requirement: MS-OXCDATA_R48
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(PropertyTag),
                        propertyProblem.PropertyTag.GetType(),
                        "MS-OXCDATA",
                        48,
                        @"[In PropertyProblem Structure] PropertyTag (4 bytes): A PropertyTag structure, as specified in section 2.9.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCDATA_R50");

                    // Verify MS-OXCROPS requirement: MS-OXCDATA_R50
                    Site.CaptureRequirementIfAreEqual<Type>(
                        typeof(uint),
                        propertyProblem.ErrorCode.GetType(),
                        "MS-OXCDATA",
                        50,
                        @"[In PropertyProblem Structure] ErrorCode (4 bytes): An unsigned integer.");
                }
            }
        }

        /// <summary>
        /// Verify RopDeletePropertiesNoReplicate Failure Response
        /// </summary>
        /// <param name="ropDeletePropertiesNoReplicateResponse">The failure response of RopDeletePropertiesNoReplicate request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopDeletePropertiesNoReplicate request</param>
        private void VerifyRopDeletePropertiesNoReplicateFailureResponse(RopDeletePropertiesNoReplicateResponse ropDeletePropertiesNoReplicateResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2944");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2944
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesNoReplicateResponse.RopId.GetType(),
                2944,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2946");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2946
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopDeletePropertiesNoReplicate,
                ropDeletePropertiesNoReplicateResponse.RopId,
                2946,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopDeletePropertiesNoReplicate], this field is set to 0x7A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2947");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2947
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropDeletePropertiesNoReplicateResponse.InputHandleIndex.GetType(),
                2947,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2948");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2948
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropDeletePropertiesNoReplicateResponse.InputHandleIndex,
                2948,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2950");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2950
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropDeletePropertiesNoReplicateResponse.ReturnValue.GetType(),
                2950,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2952");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2952
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropDeletePropertiesNoReplicateResponse.ReturnValue,
                2952,
                @"[In RopDeletePropertiesNoReplicate ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }

        #endregion

        #region RopQueryNamedProperties

        /// <summary>
        /// Verify RopQueryNamedProperties Success Response
        /// </summary>
        /// <param name="ropQueryNamedPropertiesResponse">The success response of RopQueryNamedProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQueryNamedProperties request</param>
        private void VerifyRopQueryNamedPropertiesSuccessResponse(RopQueryNamedPropertiesResponse ropQueryNamedPropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2972");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2972
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryNamedPropertiesResponse.RopId.GetType(),
                2972,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2974");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2974
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryNamedProperties,
                ropQueryNamedPropertiesResponse.RopId,
                2974,
                @"[In RopQueryNamedProperties ROP Success Response Buffer,RopId (1 byte)]For this operation[RopQueryNamedProperties], this field is set to 0x5F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2975");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2975
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryNamedPropertiesResponse.InputHandleIndex.GetType(),
                2975,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2976");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2976
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryNamedPropertiesResponse.InputHandleIndex,
                2976,
                @"[In RopQueryNamedProperties ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2978");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2978
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryNamedPropertiesResponse.ReturnValue.GetType(),
                2978,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2980,the ReturnValue:{0}", ropQueryNamedPropertiesResponse.ReturnValue);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2980
            bool isVerifyR2980 =
                ropQueryNamedPropertiesResponse.ReturnValue == SuccessReturnValue ||
                ropQueryNamedPropertiesResponse.ReturnValue == ReturnValueForRopQueryNamedProperties;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2980,
                2980,
                @"[In RopQueryNamedProperties ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000 or 0x00040380.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2981");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2981
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropQueryNamedPropertiesResponse.IdCount.GetType(),
                2981,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]IdCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2983");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2983
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyId[]),
                ropQueryNamedPropertiesResponse.PropertyIds.GetType(),
                2983,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]PropertyIds (variable): An array of unsigned 16-bit integers.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2984");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2984
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropQueryNamedPropertiesResponse.IdCount,
                (ushort)ropQueryNamedPropertiesResponse.PropertyIds.Length,
                2984,
                @"[In RopQueryNamedProperties ROP Success Response Buffer, PropertyIds (variable)] The number of integers in the array is specified by the IdCount field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2987");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2987
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyName[]),
                ropQueryNamedPropertiesResponse.PropertyNames.GetType(),
                2987,
                @"[In RopQueryNamedProperties ROP Success Response Buffer]PropertyNames (variable): A list of PropertyName structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2988");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2988
            Site.CaptureRequirementIfAreEqual<ushort>(
                ropQueryNamedPropertiesResponse.IdCount,
                (ushort)ropQueryNamedPropertiesResponse.PropertyNames.Length,
                2988,
                @"[In RopQueryNamedProperties ROP Success Response Buffer,PropertyNames (variable)]The number of structures contained in this field is specified by the IdCount field.");
        }

        /// <summary>
        /// Verify RopQueryNamedProperties Failure Response
        /// </summary>
        /// <param name="ropQueryNamedPropertiesResponse">The failure response of RopQueryNamedProperties request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopQeuryNamedProperties request</param>
        private void VerifyRopQueryNamedPropertiesFailureResponse(RopQueryNamedPropertiesResponse ropQueryNamedPropertiesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2991");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2991
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryNamedPropertiesResponse.RopId.GetType(),
                2991,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2993");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2993
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopQueryNamedProperties,
                ropQueryNamedPropertiesResponse.RopId,
                2993,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopQueryNamedProperties], this field is set to 0x5F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2994");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2994
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropQueryNamedPropertiesResponse.InputHandleIndex.GetType(),
                2994,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2995");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2995
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropQueryNamedPropertiesResponse.InputHandleIndex,
                2995,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2997");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2997
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropQueryNamedPropertiesResponse.ReturnValue.GetType(),
                2997,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R2999,the return value:{0}", ropQueryNamedPropertiesResponse.ReturnValue);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R2999
            bool isVerifyR2999 =
                (ropQueryNamedPropertiesResponse.ReturnValue != SuccessReturnValue) &&
                (ropQueryNamedPropertiesResponse.ReturnValue != ReturnValueForRopQueryNamedProperties);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR2999,
                2999,
                @"[In RopQueryNamedProperties ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000 or 0x00040380.");
        }
        #endregion

        #region RopCopyProperties
        /// <summary>
        /// Verify RopCopyProperties Success Response
        /// </summary>
        /// <param name="ropCopyPropertiesResponse">The success response of RopCopyProperties request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyProperties request</param>
        private void VerifyRopCopyPropertiesSuccessResponse(RopCopyPropertiesResponse ropCopyPropertiesResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3024");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3024
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.RopId.GetType(),
                3024,
                @"[In RopCopyProperties ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3026");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3026
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyProperties,
                ropCopyPropertiesResponse.RopId,
                3026,
                @"[In RopCopyProperties ROP Success Response Buffer,RopId (1 byte)]For this operation[RopCopyProperties], this field is set to 0x67.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3027");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3027
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.SourceHandleIndex.GetType(),
                3027,
                @"[In RopCopyProperties ROP Success Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3028");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3028
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyPropertiesResponse.SourceHandleIndex,
                3028,
                @"[In RopCopyProperties ROP Success Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3030");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3030
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyPropertiesResponse.ReturnValue.GetType(),
                3030,
                @"[In RopCopyProperties ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3032");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3032
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCopyPropertiesResponse.ReturnValue,
                3032,
                @"[In RopCopyProperties ROP Success Response Buffer,ReturnValue (4 bytes) ]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3033");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3033
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropCopyPropertiesResponse.PropertyProblemCount.GetType(),
                3033,
                @"[In RopCopyProperties ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3035");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3035
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyProblem[]),
                ropCopyPropertiesResponse.PropertyProblems.GetType(),
                3035,
                @"[In RopCopyProperties ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3036");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3036
            Site.CaptureRequirementIfAreEqual<int>(
                ropCopyPropertiesResponse.PropertyProblemCount,
                ropCopyPropertiesResponse.PropertyProblems.Length,
                3036,
                @"[In RopCopyProperties ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field[PropertyProblems (variable)] is specified by the PropertyProblemCount field.");
        }

        /// <summary>
        /// Verify RopCopyProperties Null Destination Failure Response
        /// </summary>
        /// <param name="ropCopyPropertiesResponse">The null destination failure response of RopCopyProperties request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyProperties request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopCopyProperties request</param>
        private void VerifyRopCopyPropertiesNullDestinationFailureResponse(RopCopyPropertiesResponse ropCopyPropertiesResponse, byte sourceHandleIndex, uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3038");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3038
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.RopId.GetType(),
                3038,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3040");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3040
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyProperties,
                ropCopyPropertiesResponse.RopId,
                3040,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer,RopId (1 byte)]For this operation[RopCopyProperties], this field is set to 0x67.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3041");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3041
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.SourceHandleIndex.GetType(),
                3041,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3042");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3042
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyPropertiesResponse.SourceHandleIndex,
                3042,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3044");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3044
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyPropertiesResponse.ReturnValue.GetType(),
                3044,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3046");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3046
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyPropertiesResponse.ReturnValue,
                3046,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer,ReturnValue (4 bytes)]For this response[Null Destination Failure Response], this field is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3047");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3047
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyPropertiesResponse.DestHandleIndex.GetType(),
                3047,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer]DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3048");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3048
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropCopyPropertiesResponse.DestHandleIndex,
                3048,
                @"[In RopCopyProperties ROP Null Destination Failure Response Buffer,DestHandleIndex (4 bytes)]This index MUST be set to the value specified in the DestHandleIndex field in the request.");
        }

        /// <summary>
        /// Verify RopCopyProperties Failure Response
        /// </summary>
        /// <param name="ropCopyPropertiesResponse">The failure response of RopCopyProperties request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyProperties request</param>
        private void VerifyRopCopyPropertiesFailureResponse(RopCopyPropertiesResponse ropCopyPropertiesResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3050");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3050
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.RopId.GetType(),
                3050,
                @"[In RopCopyProperties ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3052");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3052
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyProperties,
                ropCopyPropertiesResponse.RopId,
                3052,
                @"[In RopCopyProperties ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopCopyProperties], this field is set to 0x67.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3053");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3053
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyPropertiesResponse.SourceHandleIndex.GetType(),
                3053,
                @"[In RopCopyProperties ROP Failure Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3054");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3054
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyPropertiesResponse.SourceHandleIndex,
                3054,
                @"[In RopCopyProperties ROP Failure Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3056");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3056
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyPropertiesResponse.ReturnValue.GetType(),
                3056,
                @"[In RopCopyProperties ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3058,the ReturnValue:{0}", ropCopyPropertiesResponse.ReturnValue);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3058
            bool isVerifyR3058 =
                (ropCopyPropertiesResponse.ReturnValue != SuccessReturnValue) &&
                (ropCopyPropertiesResponse.ReturnValue != ReturnValueForRopMoveFolderResponseAndMoveCopyMessage);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3058,
                3058,
                @"[In RopCopyProperties ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000 or 0x00000503.");
        }
        #endregion

        #region RopCopyTo
        /// <summary>
        /// Verify RopCopyTo Success Response 
        /// </summary>
        /// <param name="ropCopyToResponse">The success response of RopCopyTo request</param>
        /// <param name="sourceHandleIndex">The field of sourceHandleIndex in RopCopyTo request</param>
        private void VerifyRopCopyToSuccessResponse(RopCopyToResponse ropCopyToResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3085");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3085
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.RopId.GetType(),
                3085,
                @"[In RopCopyTo ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3087");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3087
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyTo,
                ropCopyToResponse.RopId,
                3087,
                @"[In RopCopyTo ROP Success Response Buffer,RopId (1 byte)]For this operation[RopCopyTo], this field is set to 0x39.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3088");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3088
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.SourceHandleIndex.GetType(),
                3088,
                @"[In RopCopyTo ROP Success Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3089");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3089
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyToResponse.SourceHandleIndex,
                3089,
                @"[In RopCopyTo ROP Success Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3091");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3091
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToResponse.ReturnValue.GetType(),
                3091,
                @"[In RopCopyTo ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3093");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3093
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropCopyToResponse.ReturnValue,
                3093,
                @"[In RopCopyTo ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3094");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3094
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropCopyToResponse.PropertyProblemCount.GetType(),
                3094,
                @"[In RopCopyTo ROP Success Response Buffer]PropertyProblemCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3096");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3096
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(PropertyProblem[]),
                ropCopyToResponse.PropertyProblems.GetType(),
                3096,
                @"[In RopCopyTo ROP Success Response Buffer]PropertyProblems (variable): An array of PropertyProblem structures.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3097");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3097
            Site.CaptureRequirementIfAreEqual<int>(
                ropCopyToResponse.PropertyProblemCount,
                ropCopyToResponse.PropertyProblems.Length,
                3097,
                @"[In RopCopyTo ROP Success Response Buffer,PropertyProblems (variable)]The number of structures contained in this field is specified by the PropertyProblemCount field.");
        }

        /// <summary>
        /// Verify RopCopyTo Null Destination Failure Response
        /// </summary>
        /// <param name="ropCopyToResponse">The null destination failure response of RopCopyTo request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyTo request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopCopyTo request</param>
        private void VerifyRopCopyToNullDestinationFailureResponse(RopCopyToResponse ropCopyToResponse, byte sourceHandleIndex, uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3099");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3099
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.RopId.GetType(),
                3099,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3101");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3101
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyTo,
                ropCopyToResponse.RopId,
                3101,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer,RopId (1 byte)]For this operation[RopCopyTo], this field is set to 0x39.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3102");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3102
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.SourceHandleIndex.GetType(),
                3102,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3103");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3103
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyToResponse.SourceHandleIndex,
                3103,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3105");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3105
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToResponse.ReturnValue.GetType(),
                3105,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3107");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3107
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyToResponse.ReturnValue,
                3107,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer,ReturnValue (4 bytes)]For this response[Null Destination Failure Response], this field is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3108");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3108
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToResponse.DestHandleIndex.GetType(),
                3108,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer]DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3109");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3109
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropCopyToResponse.DestHandleIndex,
                3109,
                @"[In RopCopyTo ROP Null Destination Failure Response Buffer,DestHandleIndex (4 bytes)]This index MUST be set to the DestHandleIndex specified in the request.");
        }

        /// <summary>
        /// Verify RopCopyTo Failure Response
        /// </summary>
        /// <param name="ropCopyToResponse">The failure response of RopCopyTo request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyTo request</param>
        private void VerifyRopCopyToFailureResponse(RopCopyToResponse ropCopyToResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3111");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3111
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.RopId.GetType(),
                3111,
                @"[In RopCopyTo ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3113");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3113
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyTo,
                ropCopyToResponse.RopId,
                3113,
                @"[In RopCopyTo ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopCopyTo], this field is set to 0x39.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3114");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3114
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToResponse.SourceHandleIndex.GetType(),
                3114,
                @"[In RopCopyTo ROP Failure Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3115");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3115
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyToResponse.SourceHandleIndex,
                3115,
                @"[In RopCopyTo ROP Failure Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3117");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3117
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToResponse.ReturnValue.GetType(),
                3117,
                @"[In RopCopyTo ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3119,the ReturnValue:{0}", ropCopyToResponse.ReturnValue);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3119
            bool isVerifyR3119 =
                (ropCopyToResponse.ReturnValue != SuccessReturnValue) &&
                (ropCopyToResponse.ReturnValue != ReturnValueForRopMoveFolderResponseAndMoveCopyMessage);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3119,
                3119,
                @"[In RopCopyTo ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000 or 0x00000503.");
        }
        #endregion

        #region RopProgress
        /// <summary>
        /// Verify RopProgress Success Response
        /// </summary>
        /// <param name="ropProgressResponse">The success response of RopProgress request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopProgress request</param>
        private void VerifyRopProgressSuccessResponse(RopProgressResponse ropProgressResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3132");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3132
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropProgressResponse.RopId.GetType(),
                3132,
                @"[In RopProgress ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3134");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3134
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopProgress,
                ropProgressResponse.RopId,
                3134,
                @"[In RopProgress ROP Success Response Buffer,RopId (1 byte)]For this operation[RopProgress], this field is set to 0x50.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3135");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3135
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropProgressResponse.InputHandleIndex.GetType(),
                3135,
                @"[In RopProgress ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3136");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3136
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropProgressResponse.InputHandleIndex,
                3136,
                @"[In RopProgress ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3138");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3138
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropProgressResponse.ReturnValue.GetType(),
                3138,
                @"[In RopProgress ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3140");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3140
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropProgressResponse.ReturnValue,
                3140,
                @"[In RopProgress ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3141");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3141
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropProgressResponse.LogonId.GetType(),
                3141,
                @"[In RopProgress ROP Success Response Buffer]LogonId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3143");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3143
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropProgressResponse.CompletedTaskCount.GetType(),
                3143,
                @"[In RopProgress ROP Success Response Buffer]CompletedTaskCount (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3145");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3145
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropProgressResponse.TotalTaskCount.GetType(),
                3145,
                @"[In RopProgress ROP Success Response Buffer]TotalTaskCount (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopProgress Failure Response
        /// </summary>
        /// <param name="ropProgressResponse">The failure response of RopProgress request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopProgress request</param>
        private void VerifyRopProgressFailureResponse(RopProgressResponse ropProgressResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3147");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3147
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropProgressResponse.RopId.GetType(),
                3147,
                @"[In RopProgress ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3149");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3149
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopProgress,
                ropProgressResponse.RopId,
                3149,
                @"[In RopProgress ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopProgress], this field is set to 0x50.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3150");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3150
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropProgressResponse.InputHandleIndex.GetType(),
                3150,
                @"[In RopProgress ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3151");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3151
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropProgressResponse.InputHandleIndex,
                3151,
                @"[In RopProgress ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3153");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3153
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropProgressResponse.ReturnValue.GetType(),
                3153,
                @"[In RopProgress ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3155");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3155
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropProgressResponse.ReturnValue,
                3155,
                @"[In RopProgress ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this field is set to a value other than 0x00000000.");
        }
        #endregion
        #endregion

        #region Stream ROPs

        #region RopOpenStream
        /// <summary>
        /// Verify RopOpenStream Success Response
        /// </summary>
        /// <param name="ropOpenStreamResponse">The success response of RopOpenStream request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenStream request</param>
        private void VerifyRopOpenStreamSuccessResponse(RopOpenStreamResponse ropOpenStreamResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3175");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3175
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenStreamResponse.RopId.GetType(),
                3175,
                @"[In RopOpenStream ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3177");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3177
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenStream,
                ropOpenStreamResponse.RopId,
                3177,
                @"[In RopOpenStream ROP Success Response Buffer,RopId (1 byte)]For this operation[RopOpenStream], this field is set to 0x2B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3178");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3178
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenStreamResponse.OutputHandleIndex.GetType(),
                3178,
                @"[In RopOpenStream ROP Success Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3179");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3179
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenStreamResponse.OutputHandleIndex,
                3179,
                @"[In RopOpenStream ROP Success Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3181");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenStreamResponse.ReturnValue.GetType(),
                3181,
                @"[In RopOpenStream ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3183");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3183
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropOpenStreamResponse.ReturnValue,
                3183,
                @"[In RopOpenStream ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3184");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3184
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenStreamResponse.StreamSize.GetType(),
                3184,
                @"[In RopOpenStream ROP Success Response Buffer]StreamSize (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopOpenStream Failure Response
        /// </summary>
        /// <param name="ropOpenStreamResponse">The response of RopOpenStream failure request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopOpenStream request</param>
        private void VerifyRopOpenStreamFailureResponse(RopOpenStreamResponse ropOpenStreamResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3186");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3186
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenStreamResponse.RopId.GetType(),
                3186,
                @"[In RopOpenStream ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3188");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3188
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopOpenStream,
                ropOpenStreamResponse.RopId,
                3188,
                @"[In RopOpenStream ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopOpenStream], this field is set to 0x2B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3189");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3189
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropOpenStreamResponse.OutputHandleIndex.GetType(),
                3189,
                @"[In RopOpenStream ROP Failure Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3190");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3190
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropOpenStreamResponse.OutputHandleIndex,
                3190,
                @"[In RopOpenStream ROP Failure Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3192");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3192
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropOpenStreamResponse.ReturnValue.GetType(),
                3192,
                @"[In RopOpenStream ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3194");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3194
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropOpenStreamResponse.ReturnValue,
                3194,
                @"[In Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this
                field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopReadStream
        /// <summary>
        /// Verify RopReadStream Response
        /// </summary>
        /// <param name="ropReadStreamResponse">The response of RopReadStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopReadStream request</param>
        private void VerifyRopReadStreamResponse(RopReadStreamResponse ropReadStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3216");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3216
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadStreamResponse.RopId.GetType(),
                3216,
                @"[In RopReadStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3218");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3218
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopReadStream,
                ropReadStreamResponse.RopId,
                3218,
                @"[In RopReadStream ROP Response Buffer,RopId (1 byte)]For this operation[RopReadStream], this field is set to 0x2C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3219");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3219
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropReadStreamResponse.InputHandleIndex.GetType(),
                3219,
                @"[In RopReadStream ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3220");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3220
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropReadStreamResponse.InputHandleIndex,
                3220,
                @"[In RopReadStream ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3222");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3222
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropReadStreamResponse.ReturnValue.GetType(),
                3222,
                @"[In RopReadStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3224");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3224
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropReadStreamResponse.DataSize.GetType(),
                3224,
                @"[In RopReadStream ROP Response Buffer]DataSize (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3229");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3229
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropReadStreamResponse.Data.GetType(),
                3229,
                @"[In RopReadStream ROP Response Buffer]Data (variable): An array of bytes.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3231");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3231
            Site.CaptureRequirementIfAreEqual<int>(
                ropReadStreamResponse.DataSize,
                ropReadStreamResponse.Data.Length,
                3231,
                @"[In RopReadStream ROP Response Buffer,Data (variable)]The size of this field, in bytes, is specified by the DataSize field.");
        }
        #endregion

        #region RopWriteStream
        /// <summary>
        /// Verify RopWriteStream Response
        /// </summary>
        /// <param name="ropWriteStreamResponse">The response of RopWriteStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopWriteStream request</param>
        private void VerifyRopWriteStreamResponse(RopWriteStreamResponse ropWriteStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3247");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3247
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropWriteStreamResponse.RopId.GetType(),
                3247,
                @"[In RopWriteStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3249");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3249
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopWriteStream,
                ropWriteStreamResponse.RopId,
                3249,
                @"[In RopWriteStream ROP Response Buffer,RopId (1 byte)]For this operation[RopWriteStream], this field is set to 0x2D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3250");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3250
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropWriteStreamResponse.InputHandleIndex.GetType(),
                3250,
                @"[In RopWriteStream ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3251");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3251
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropWriteStreamResponse.InputHandleIndex,
                3251,
                @"[In RopWriteStream ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3253");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3253
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropWriteStreamResponse.ReturnValue.GetType(),
                3253,
                @"[In RopWriteStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3255");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3255
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropWriteStreamResponse.WrittenSize.GetType(),
                3255,
                @"[In RopWriteStream ROP Response Buffer]WrittenSize (2 bytes): An unsigned integer.");
        }

        #endregion

        #region RopCommitStream
        /// <summary>
        /// Verify RopCommitStream Response
        /// </summary>
        /// <param name="ropCommitStreamResponse">The response of RopCommitStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopCommitStream request</param>
        private void VerifyRopCommitStreamResponse(RopCommitStreamResponse ropCommitStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3267");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3267
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCommitStreamResponse.RopId.GetType(),
                3267,
                @"[In RopCommitStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3269");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3269
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCommitStream,
                ropCommitStreamResponse.RopId,
                3269,
                @"[In RopCommitStream ROP Response Buffer,RopId (1 byte)]For this operation[RopCommitStream], this field is set to 0x5D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3270");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3270
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCommitStreamResponse.InputHandleIndex.GetType(),
                3270,
                @"[In RopCommitStream ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3271");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3271
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropCommitStreamResponse.InputHandleIndex,
                3271,
                @"[In RopCommitStream ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3273");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3273
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCommitStreamResponse.ReturnValue.GetType(),
                3273,
                @"[In RopCommitStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetStreamSize
        /// <summary>
        /// Verify RopGetStreamSize Success Response
        /// </summary>
        /// <param name="ropGetStreamSizeResponse">The success response of RopGetStreamSize request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStreamSize request</param>
        private void VerifyRopGetStreamSizeSuccessResponse(RopGetStreamSizeResponse ropGetStreamSizeResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3285");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3285
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStreamSizeResponse.RopId.GetType(),
                3285,
                @"[In RopGetStreamSize ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3287");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3287
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStreamSize,
                ropGetStreamSizeResponse.RopId,
                3287,
                @"[In RopGetStreamSize ROP Success Response Buffer,RopId (1 byte)]For this operation[RopGetStreamSize], this field is set to 0x5E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3288");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3288
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStreamSizeResponse.InputHandleIndex.GetType(),
                3288,
                @"[In RopGetStreamSize ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3289");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3289
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStreamSizeResponse.InputHandleIndex,
                3289,
                @"[In RopGetStreamSize ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3291");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3291
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStreamSizeResponse.ReturnValue.GetType(),
                3291,
                @"[In RopGetStreamSize ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3293");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3293
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetStreamSizeResponse.ReturnValue,
                3293,
                @"[In RopGetStreamSize ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3294");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3294
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStreamSizeResponse.StreamSize.GetType(),
                3294,
                @"[In RopGetStreamSize ROP Success Response Buffer]StreamSize (4 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopGetStreamSize Failure Response
        /// </summary>
        /// <param name="ropGetStreamSizeResponse">The failure response of RopGetStreamSize request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetStreamSize request</param>
        private void VerifyRopGetStreamSizeFailureResponse(RopGetStreamSizeResponse ropGetStreamSizeResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3296");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3296
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStreamSizeResponse.RopId.GetType(),
                3296,
                @"[In RopGetStreamSize ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3298");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3298
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetStreamSize,
                ropGetStreamSizeResponse.RopId,
                3298,
                @"[In RopGetStreamSize ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopGetStreamSize], this field is set to 0x5E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3299");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3299
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetStreamSizeResponse.InputHandleIndex.GetType(),
                3299,
                @"[In RopGetStreamSize ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3300");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3300
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetStreamSizeResponse.InputHandleIndex,
                3300,
                @"[In RopGetStreamSize ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3302");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3302
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetStreamSizeResponse.ReturnValue.GetType(),
                3302,
                @"[In RopGetStreamSize ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3304");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3304
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetStreamSizeResponse.ReturnValue,
                3304,
                @"[In RopGetStreamSize ROP Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSetStreamSize
        /// <summary>
        /// Verify RopSetStreamSize Response
        /// </summary>
        /// <param name="ropSetStreamSizeResponse">The response of RopSetStreamSize request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetStreamSize request</param>
        private void VerifyRopSetStreamSizeResponse(RopSetStreamSizeResponse ropSetStreamSizeResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3321");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3321
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetStreamSizeResponse.RopId.GetType(),
                3321,
                @"[In RopSetStreamSize ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3323");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3323
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetStreamSize,
                ropSetStreamSizeResponse.RopId,
                3323,
                @"[In RopSetStreamSize ROP Response Buffer,RopId (1 byte)]For this operation[RopSetStreamSize], this field is set to 0x2F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3324");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3324
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetStreamSizeResponse.InputHandleIndex.GetType(),
                3324,
                @"[In RopSetStreamSize ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3325");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3325
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetStreamSizeResponse.InputHandleIndex,
                3325,
                @"[In RopSetStreamSize ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3327");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3327
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetStreamSizeResponse.ReturnValue.GetType(),
                3327,
                @"[In RopSetStreamSize ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }

        #endregion

        #region RopSeekStream
        /// <summary>
        /// Verify RopSeekStream Success Response
        /// </summary>
        /// <param name="ropSeekStreamResponse">The success response of RopSeekStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekStream request</param>
        private void VerifyRopSeekStreamSuccessResponse(RopSeekStreamResponse ropSeekStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3344");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3344
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekStreamResponse.RopId.GetType(),
                3344,
                @"[In RopSeekStream ROP Success Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3346");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3346
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekStream,
                ropSeekStreamResponse.RopId,
                3346,
                @"[In RopSeekStream ROP Success Response Buffer,RopId (1 byte)]For this operation[[RopSeekStream]], this field is set to 0x2E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3347");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3347
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekStreamResponse.InputHandleIndex.GetType(),
                3347,
                @"[In RopSeekStream ROP Success Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3348");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3348
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekStreamResponse.InputHandleIndex,
                3348,
                @"[In RopSeekStream ROP Success Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3350");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3350
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekStreamResponse.ReturnValue.GetType(),
                3350,
                @"[In RopSeekStream ROP Success Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3352");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3352
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSeekStreamResponse.ReturnValue,
                3352,
                @"[In RopSeekStream ROP Success Response Buffer,ReturnValue (4 bytes)]For this response[Success Response ], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3353");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3353
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropSeekStreamResponse.NewPosition.GetType(),
                3353,
                @"[In RopSeekStream ROP Success Response Buffer]NewPosition (8 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopSeekStream Failure Response
        /// </summary>
        /// <param name="ropSeekStreamResponse">The failure response of RopSeekStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSeekStream request</param>
        private void VerifyRopSeekStreamFailureResponse(RopSeekStreamResponse ropSeekStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3355");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3355
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekStreamResponse.RopId.GetType(),
                3355,
                @"[In RopSeekStream ROP Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3357");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3357
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSeekStream,
                ropSeekStreamResponse.RopId,
                3357,
                @"[In RopSeekStream ROP Failure Response Buffer,RopId (1 byte)]For this operation[RopSeekStream], this field is set to 0x2E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3358");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3358
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSeekStreamResponse.InputHandleIndex.GetType(),
                3358,
                @"[In RopSeekStream ROP Failure Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3359");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3359
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSeekStreamResponse.InputHandleIndex,
                3359,
                @"[In RopSeekStream ROP Failure Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3361");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3361
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSeekStreamResponse.ReturnValue.GetType(),
                3361,
                @"[In RopSeekStream ROP Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3363");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3363
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSeekStreamResponse.ReturnValue,
                3363,
                @"[In Failure Response Buffer,ReturnValue (4 bytes)]For this response[Failure Response ], this
                field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopCopyToStream
        /// <summary>
        /// Verify RopCopyToStream Response
        /// </summary>
        /// <param name="ropCopyToStreamResponse">The response of RopCopyToStream request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyToStream request</param>
        private void VerifyRopCopyToStreamResponse(RopCopyToStreamResponse ropCopyToStreamResponse, byte sourceHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3379");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3379
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToStreamResponse.RopId.GetType(),
                3379,
                @"[In RopCopyToStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3381");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3381
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyToStream,
                ropCopyToStreamResponse.RopId,
                3381,
                @"[In RopCopyToStream ROP Response Buffer,RopId (1 byte)]For this operation[RopCopyToStream], this field is set to 0x3A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3382");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3382
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToStreamResponse.SourceHandleIndex.GetType(),
                3382,
                @"[In RopCopyToStream ROP Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3383");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3383
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyToStreamResponse.SourceHandleIndex,
                3383,
                @"[In RopCopyToStream ROP Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3385");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3385
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToStreamResponse.ReturnValue.GetType(),
                3385,
                @"[In RopCopyToStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3387");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3387
            Site.CaptureRequirementIfAreNotEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyToStreamResponse.ReturnValue,
                3387,
                @"[In RopCopyToStream ROP Response Buffer,ReturnValue (4 bytes)]For this response[except Null Destination Failure Response ], this field is set to a value other than 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3388");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3388
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropCopyToStreamResponse.ReadByteCount.GetType(),
                3388,
                @"[In RopCopyToStream ROP Response Buffer]ReadByteCount (8 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3390");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3390
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropCopyToStreamResponse.WrittenByteCount.GetType(),
                3390,
                @"[In RopCopyToStream ROP Response Buffer]WrittenByteCount (8 bytes): An unsigned integer.");
        }

        /// <summary>
        /// Verify RopCopyToStream Null Destination Failure Response
        /// </summary>
        /// <param name="ropCopyToStreamResponse">The null destination failure response of RopCopyToStream request</param>
        /// <param name="sourceHandleIndex">The field of SourceHandleIndex in RopCopyToStream request</param>
        /// <param name="destHandleIndex">The field of DestHandleIndex in RopCopyToStream request</param>
        private void VerifyRopCopyToStreamNullDestinationFailureResponse(RopCopyToStreamResponse ropCopyToStreamResponse, byte sourceHandleIndex, uint destHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3392");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3392
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToStreamResponse.RopId.GetType(),
                3392,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3394");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3394
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCopyToStream,
                ropCopyToStreamResponse.RopId,
                3394,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer,RopId (1 byte)]For this operation[RopCopyToStream], this field is set to 0x3A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3395");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3395
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCopyToStreamResponse.SourceHandleIndex.GetType(),
                3395,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]SourceHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3396");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3396
            Site.CaptureRequirementIfAreEqual<byte>(
                sourceHandleIndex,
                ropCopyToStreamResponse.SourceHandleIndex,
                3396,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer,SourceHandleIndex (1 byte)]This index MUST be set to the value specified in the SourceHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3398");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3398
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToStreamResponse.ReturnValue.GetType(),
                3398,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3400");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3400
            Site.CaptureRequirementIfAreEqual<uint>(
                ReturnValueForRopMoveFolderResponseAndMoveCopyMessage,
                ropCopyToStreamResponse.ReturnValue,
                3400,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer,ReturnValue (4 bytes)]For this response[Null Destination Failure Response], this field is set to 0x00000503.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3401");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3401
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCopyToStreamResponse.DestHandleIndex.GetType(),
                3401,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]DestHandleIndex (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3402");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3402
            Site.CaptureRequirementIfAreEqual<uint>(
                destHandleIndex,
                ropCopyToStreamResponse.DestHandleIndex,
                3402,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer,DestHandleIndex (4 bytes)]This index MUST be set to the specified in the DestHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3404");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3404
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropCopyToStreamResponse.ReadByteCount.GetType(),
                3404,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]ReadByteCount (8 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3412");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3412
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropCopyToStreamResponse.WrittenByteCount.GetType(),
                3412,
                @"[In RopCopyToStream ROP Null Destination Failure Response Buffer]WrittenByteCount (8 bytes): An unsigned integer.");

            if (Common.IsRequirementEnabled(3411, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3411,the ReadByteCount:{0}", ropCopyToStreamResponse.ReadByteCount);

                // Verify MS-OXCROPS requirement: 3411
                Site.CaptureRequirementIfAreEqual<ulong>(
                    ReadOrWrittenByteCountForRopCopyToStream,
                    ropCopyToStreamResponse.ReadByteCount,
                    3411,
                    @"[In Appendix B: Product Behavior] For this response [of read from the source object ] , Implementation does set this field[ReadByteCount] to 0x0000000000000000. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(4692, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4692");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4692
                Site.CaptureRequirementIfAreNotEqual<ulong>(
                    ReadOrWrittenByteCountForRopCopyToStream,
                    ropCopyToStreamResponse.ReadByteCount,
                    4692,
                    @"[In Appendix B: Product Behavior] Implementation does not set the field[ReadByteCount] to 0x0000000000000000, but the non zero value has no meaning.  (<11> Section 2.2.9.8.3: Exchange 2007 do not set the field to 0x0000000000000000, but the nonzero value has no meaning.)");
            }

            if (Common.IsRequirementEnabled(3415, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3415,the WrittenByteCount:{0}", ropCopyToStreamResponse.WrittenByteCount);

                // Verify MS-OXCROPS requirement: 3415
                Site.CaptureRequirementIfAreEqual<ulong>(
                    ReadOrWrittenByteCountForRopCopyToStream,
                    ropCopyToStreamResponse.WrittenByteCount,
                    3415,
                    @"[In Appendix B: Product Behavior] For this response [of written to the destination object], Implementation does set this field[WrittenByteCount] to 0x0000000000000000. (Microsoft Exchange Server 2010 and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(4693, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4693");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R4693
                Site.CaptureRequirementIfAreNotEqual<ulong>(
                    ReadOrWrittenByteCountForRopCopyToStream,
                    ropCopyToStreamResponse.WrittenByteCount,
                    4693,
                    @"[In Appendix B: Product Behavior] Implementation does not set the field[WrittenByteCount] to 0x0000000000000000, but the non zero value has no meaning. (<12> Section 2.2.9.8.3: Exchange 2007 do not set the field to 0x0000000000000000, but the nonzero value has no meaning.)");
            }
        }

        #endregion

        #region RopLockRegionStream
        /// <summary>
        /// Verify RopLockRegionStream Response
        /// </summary>
        /// <param name="ropLockRegionStreamResponse">The response of RopLockRegionStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopLockRegionStream request</param>
        private void VerifyRopLockRegionStreamResponse(RopLockRegionStreamResponse ropLockRegionStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3433");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3433
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLockRegionStreamResponse.RopId.GetType(),
                3433,
                @"[In RopLockRegionStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3435");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3435
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopLockRegionStream,
                ropLockRegionStreamResponse.RopId,
                3435,
                @"[In RopLockRegionStream ROP Response Buffer,RopId (1 byte)]For this operation[RopLockRegionStream], this field is set to 0x5B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3436");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3436
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropLockRegionStreamResponse.InputHandleIndex.GetType(),
                3436,
                @"[In RopLockRegionStream ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3437");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3437
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropLockRegionStreamResponse.InputHandleIndex,
                3437,
                @"[In RopLockRegionStream ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3439");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3439
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropLockRegionStreamResponse.ReturnValue.GetType(),
                3439,
                @"[In RopLockRegionStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopUnlockRegionStream
        /// <summary>
        /// Verify RopUnlockRegionStream Response
        /// </summary>
        /// <param name="ropUnlockRegionStreamResponse">The response of RopUnlockRegionStream request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopUnlockRegionStream request</param>
        private void VerifyRopUnlockRegionStreamResponse(RopUnlockRegionStreamResponse ropUnlockRegionStreamResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3458");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3458
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropUnlockRegionStreamResponse.RopId.GetType(),
                3458,
                @"[In RopUnlockRegionStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3460");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3460
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopUnlockRegionStream,
                ropUnlockRegionStreamResponse.RopId,
                3460,
                @"[In RopUnlockRegionStream ROP Response Buffer,RopId (1 byte)]For this operation[RopUnlockRegionStream], this field is set to 0x5C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3461");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3461
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropUnlockRegionStreamResponse.InputHandleIndex.GetType(),
                3461,
                @"[In RopUnlockRegionStream ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3462");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3462
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropUnlockRegionStreamResponse.InputHandleIndex,
                3462,
                @"[In RopUnlockRegionStream ROP Response Buffer,InputHandleIndex (1 byte)]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3464");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3464
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropUnlockRegionStreamResponse.ReturnValue.GetType(),
                3464,
                @"[In RopUnlockRegionStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopCloneStream
        /// <summary>
        /// Verify RopCloneStream Response
        /// </summary>
        /// <param name="ropCloneStreamResponse">The response of RopCloneStream request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopCloneStream request</param>
        private void VerifyRopCloneStreamResponse(RopCloneStreamResponse ropCloneStreamResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3495");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3495
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCloneStreamResponse.RopId.GetType(),
                3495,
                @"[In RopCloneStream ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3497");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3497
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopCloneStream,
                ropCloneStreamResponse.RopId,
                3497,
                @"[In RopCloneStream ROP Response Buffer,RopId (1 byte)]For this operation[RopCloneStream], this field is set to 0x3B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3498");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3498
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropCloneStreamResponse.OutputHandleIndex.GetType(),
                3498,
                @"[In RopCloneStream ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3499");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3499
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropCloneStreamResponse.OutputHandleIndex,
                3499,
                @"[In RopCloneStream ROP Response Buffer,OutputHandleIndex (1 byte)]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3501");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3501
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropCloneStreamResponse.ReturnValue.GetType(),
                3501,
                @"[In RopCloneStream ROP Response Buffer]ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Permission ROPs

        #region RopModifyPermissions
        /// <summary>
        /// Verify RopModifyPermissions Response
        /// </summary>
        /// <param name="ropModifyPermissionsResponse">The response of RopModifyPermissions request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopModifyPermissions request</param>
        private void VerifyRopModifyPermissionsResponse(RopModifyPermissionsResponse ropModifyPermissionsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3530");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3530
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyPermissionsResponse.RopId.GetType(),
                3530,
                @"[In RopModifyPermissions ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3532");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3532
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopModifyPermissions,
                ropModifyPermissionsResponse.RopId,
                3532,
                @"[In RopModifyPermissions ROP Response Buffer, RopId (1 byte):] For this operation [RopModifyPermissions], this field is set to 0x40.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3533");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3533
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyPermissionsResponse.InputHandleIndex.GetType(),
                3533,
                @"[In RopModifyPermissions ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3534");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3534
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropModifyPermissionsResponse.InputHandleIndex,
                3534,
                @"[In RopModifyPermissions ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3536");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3536
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropModifyPermissionsResponse.ReturnValue.GetType(),
                3536,
                @"[In RopModifyPermissions ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetPermissionsTable
        /// <summary>
        /// Verify RopGetPermissionsTable Response
        /// </summary>
        /// <param name="ropGetPermissionsTableResponse">The response of RopGetPermissionsTable request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopGetPermissionsTable request</param>
        private void VerifyRopGetPermissionsTableResponse(RopGetPermissionsTableResponse ropGetPermissionsTableResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3554");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3554
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPermissionsTableResponse.RopId.GetType(),
                3554,
                @"[In RopGetPermissionsTable ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3556");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3556
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetPermissionsTable,
                ropGetPermissionsTableResponse.RopId,
                3556,
                @"[In RopGetPermissionsTable ROP Response Buffer, RopId (1 byte):] For this operation[RopGetPermissionsTable], this field is set to 0x3E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3557");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3557
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetPermissionsTableResponse.OutputHandleIndex.GetType(),
                3557,
                @"[In RopGetPermissionsTable ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3558");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3558
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropGetPermissionsTableResponse.OutputHandleIndex,
                3558,
                @"[In RopGetPermissionsTable ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3560");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3560
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetPermissionsTableResponse.ReturnValue.GetType(),
                3560,
                @"[In RopGetPermissionsTable ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Rule ROPs

        #region RopModifyRules
        /// <summary>
        /// Verify RopModifyRules Response
        /// </summary>
        /// <param name="ropModifyRulesResponse">The response of RopModifyRules request</param>
        private void VerifyRopModifyRulesResponse(RopModifyRulesResponse ropModifyRulesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3588");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3588
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyRulesResponse.RopId.GetType(),
                3588,
                @"[In RopModifyRules ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3590");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3590
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopModifyRules,
                ropModifyRulesResponse.RopId,
                3590,
                @"[In RopModifyRules ROP Response Buffer, RopId (1 byte):] For this operation [RopModifyRules] this field is set to 0x41.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3591");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3591
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropModifyRulesResponse.InputHandleIndex.GetType(),
                3591,
                @"[In RopModifyRules ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3594");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3594
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropModifyRulesResponse.ReturnValue.GetType(),
                3594,
                @"[In RopModifyRules ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetRulesTable
        /// <summary>
        /// Verify RopGetRulesTable Response
        /// </summary>
        /// <param name="ropGetRulesTableResponse">The response of RopGetRulesTable request</param>
        private void VerifyRopGetRulesTableResponse(RopGetRulesTableResponse ropGetRulesTableResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3612");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3612
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetRulesTableResponse.RopId.GetType(),
                3612,
                @"[In RopGetRulesTable ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3614");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3614
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetRulesTable,
                ropGetRulesTableResponse.RopId,
                3614,
                @"[In RopGetRulesTable ROP Response Buffer,RopId (1 byte):] For this operation[RopGetRulesTable] this field is set to 0x3F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3615");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3615
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetRulesTableResponse.OutputHandleIndex.GetType(),
                3615,
                @"[In RopGetRulesTable ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3618");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3618
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetRulesTableResponse.ReturnValue.GetType(),
                3618,
                @"[In RopGetRulesTable ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopUpdateDeferredActionMessages
        /// <summary>
        /// Verify RopUpdateDeferredActionMessages Response
        /// </summary>
        /// <param name="ropUpdateDeferredActionMessagesResponse">The response of RopUpdateDeferredActionMessages request</param>
        private void VerifyRopUpdateDeferredActionMessagesResponse(RopUpdateDeferredActionMessagesResponse ropUpdateDeferredActionMessagesResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3640");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3640
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropUpdateDeferredActionMessagesResponse.RopId.GetType(),
                3640,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3642");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3642
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopUpdateDeferredActionMessages,
                ropUpdateDeferredActionMessagesResponse.RopId,
                3642,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer, RopId (1 byte):] For this operation [ RopUpdateDeferredActionMessages] this field is set to 0x57.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3643");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3643
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropUpdateDeferredActionMessagesResponse.InputHandleIndex.GetType(),
                3643,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3646");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3646
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropUpdateDeferredActionMessagesResponse.ReturnValue.GetType(),
                3646,
                @"[In RopUpdateDeferredActionMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Fast Transfer ROPs

        #region RopFastTransferDestinationConfigure
        /// <summary>
        /// Verify RopFastTransferDestinationConfigure Response 
        /// </summary>
        /// <param name="ropFastTransferDestinationConfigureResponse">The response of RopFastTransferDestinationConfigure request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopFastTransferDestinationConfigure request</param>
        private void VerifyRopFastTransferDestinationConfigureResponse(RopFastTransferDestinationConfigureResponse ropFastTransferDestinationConfigureResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3667");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3667
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferDestinationConfigureResponse.RopId.GetType(),
                3667,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3669");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3669
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferDestinationConfigure,
                ropFastTransferDestinationConfigureResponse.RopId,
                3669,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer, RopId (1 byte):] For this operation [RopFastTransferDestinationConfigure] this field is set to 0x53.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3670");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3670
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferDestinationConfigureResponse.OutputHandleIndex.GetType(),
                3670,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3671");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3671
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropFastTransferDestinationConfigureResponse.OutputHandleIndex,
                3671,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3673");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3673
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferDestinationConfigureResponse.ReturnValue.GetType(),
                3673,
                @"[In RopFastTransferDestinationConfigure ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopFastTransferDestinationPutBuffer
        /// <summary>
        /// Verify RopFastTransferDestinationPutBuffer Response
        /// </summary>
        /// <param name="ropFastTransferDestinationPutBufferResponse">The response of RopFastTransferDestinationPutBuffer request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopFastTransferDestinationPutBuffer request</param>
        private void VerifyRopFastTransferDestinationPutBufferResponse(RopFastTransferDestinationPutBufferResponse ropFastTransferDestinationPutBufferResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3690");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3690
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferDestinationPutBufferResponse.RopId.GetType(),
                3690,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3705");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3705
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferDestinationPutBufferResponse.RopId.GetType(),
                3705,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] Reserved (1 byte): Reserved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3692");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3692
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferDestinationPutBuffer,
                ropFastTransferDestinationPutBufferResponse.RopId,
                3692,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer, RopId (1 byte):] For this operation[RopFastTransferDestinationPutBuffer] this field is set to 0x54.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3693");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3693
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferDestinationPutBufferResponse.InputHandleIndex.GetType(),
                3693,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3694");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3694
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropFastTransferDestinationPutBufferResponse.InputHandleIndex,
                3694,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3696");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3696
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferDestinationPutBufferResponse.ReturnValue.GetType(),
                3696,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3698");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3698
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferDestinationPutBufferResponse.TransferStatus.GetType(),
                3698,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer]TransferStatus (2 bytes): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3699,TransferStatus's value:{0}", ropFastTransferDestinationPutBufferResponse.TransferStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3699
            bool isVerifyR3699 = (ropFastTransferDestinationPutBufferResponse.TransferStatus & (ushort)TransferStatus.Error) == (ushort)TransferStatus.Error
                                 || (ropFastTransferDestinationPutBufferResponse.TransferStatus & (ushort)TransferStatus.Done) == (ushort)TransferStatus.Done
                                 || (ropFastTransferDestinationPutBufferResponse.TransferStatus & (ushort)TransferStatus.Partial) == (ushort)TransferStatus.Partial;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3699,
                3699,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer, TransferStatus (2 bytes):] The possible values[the value of Error is 0x0000, the value of Partial is 0x0001 and the value of Done is 0x0003] for this enumeration are specified in [MS-OXCFXICS] section 2.2.3.1.2.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3701");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3701
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferDestinationPutBufferResponse.InProgressCount.GetType(),
                3701,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] InProgressCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3703");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3703
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferDestinationPutBufferResponse.TotalStepCount.GetType(),
                3703,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer] TotalStepCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3706,the Reserved's value:{0}", ropFastTransferDestinationPutBufferResponse.Reserved);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3706
            bool isVerifyR3706 =
                (typeof(byte) == ropFastTransferDestinationPutBufferResponse.Reserved.GetType()) &&
                (ropFastTransferDestinationPutBufferResponse.Reserved == ReservedOneByte);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3706,
                3706,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer, Reserved (1 byte):] The server MUST set this field to 0x00.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3707");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3707
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferDestinationPutBufferResponse.BufferUsedSize.GetType(),
                3707,
                @"[In RopFastTransferDestinationPutBuffer ROP Response Buffer]BufferUsedSize (2 bytes): An unsigned integer.");
        }
        #endregion

        #region RopFastTransferSourceGetBuffer
        /// <summary>
        /// Verify ropFastTransferSourceGetBufferResponse Response
        /// </summary>
        /// <param name="ropFastTransferSourceGetBufferResponse">The response of RopFastTransferSourceGetBufferResponse request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopFastTransferSourceGetBufferResponse request</param>
        private void VerifyRopFastTransferSourceGetBufferResponse(RopFastTransferSourceGetBufferResponse ropFastTransferSourceGetBufferResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3724");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3724
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceGetBufferResponse.RopId.GetType(),
                3724,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3726");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3726
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferSourceGetBuffer,
                ropFastTransferSourceGetBufferResponse.RopId,
                3726,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, RopId (1 byte):] For this operation [RopFastTransferSourceGetBuffer] this field is set to 0x4E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3727");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3727
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceGetBufferResponse.InputHandleIndex.GetType(),
                3727,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3728");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3728
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropFastTransferSourceGetBufferResponse.InputHandleIndex,
                3728,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3730");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3730
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferSourceGetBufferResponse.ReturnValue.GetType(),
                3730,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3732");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3732
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferSourceGetBufferResponse.TransferStatus.GetType(),
                3732,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TransferStatus (2 bytes): An enumeration.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3733,TransferStatus's value:{0}", ropFastTransferSourceGetBufferResponse.TransferStatus);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3733
            bool isVerifyR3733 = (ropFastTransferSourceGetBufferResponse.TransferStatus == (ushort)TransferStatus.Error)
                                 || (ropFastTransferSourceGetBufferResponse.TransferStatus == (ushort)TransferStatus.Partial)
                                 || (ropFastTransferSourceGetBufferResponse.TransferStatus == (ushort)TransferStatus.Done);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3733,
                3733,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, TransferStatus (2 bytes):] The possible values[the value of Error is 0x0000, the value of Partial is 0x0001 and the value of Done is 0x0003] for this enumeration are specified in [MS-OXCFXICS] section 2.2.3.1.1.5.2. ");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3736");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3736
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferSourceGetBufferResponse.InProgressCount.GetType(),
                3736,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]InProgressCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3738");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3738
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferSourceGetBufferResponse.TotalStepCount.GetType(),
                3738,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TotalStepCount (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3741,the Reserved's value:{0}", ropFastTransferSourceGetBufferResponse.Reserved);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3741
            bool isVerifyR3741 =
                (ropFastTransferSourceGetBufferResponse.Reserved.GetType() == typeof(byte)) &&
                (ropFastTransferSourceGetBufferResponse.Reserved == ReservedOneByte);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR3741,
                3741,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, Reserved (1 byte):] The server MUST set this field to 0x00.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3740");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3740
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceGetBufferResponse.Reserved.GetType(),
                3740,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer] Reserved (1 byte): Reserved.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3742");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3742
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropFastTransferSourceGetBufferResponse.TransferBufferSize.GetType(),
                3742,
                @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TransferBufferSize (2 bytes): An unsigned integer.");

            if (ropFastTransferSourceGetBufferResponse.ReturnValue != ReturnValueForRopFastTransferSourceGetBufferResponse)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3745");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3745
                // The ropFastTransferSourceGetBufferResponse.TransferBuffer isNotNull, indicating this filed is present.
                Site.CaptureRequirementIfIsNotNull(
                    ropFastTransferSourceGetBufferResponse.TransferBuffer,
                    3745,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, TransferBuffer (variable):] That is present if the ReturnValue is not 0x00000480.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3746");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3746
                // When the return value is not 0x00000480,TransferBuffer present.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    ropFastTransferSourceGetBufferResponse.TransferBufferSize,
                    (ushort)ropFastTransferSourceGetBufferResponse.TransferBuffer.Length,
                    3746,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, TransferBuffer (variable):] If present, the size of this field, in bytes, is specified by the TransferBufferSize field.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3744");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3744
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(byte[]),
                    ropFastTransferSourceGetBufferResponse.TransferBuffer.GetType(),
                    3744,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer]TransferBuffer (variable): An array of bytes.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3957");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R3957
                // When the value of the BackoffTime is null indicating this filed is not present.
                // Else the value of this filed is not equal null
                Site.CaptureRequirementIfIsTrue(
                    ropFastTransferSourceGetBufferResponse.BackoffTime == null,
                    3957,
                    @"[In RopFastTransferSourceGetBuffer ROP Response Buffer, BackoffTime (4 bytes):This field is] not present otherwise[if the ReturnValue is not 0x00000480].");
            }
        }
        #endregion

        #region RopFastTransferSourceCopyFolder
        /// <summary>
        /// Verify RopFastTransferSourceCopyFolder Response
        /// </summary>
        /// <param name="ropFastTransferSourceCopyFolderResponse">The response of RopFastTransferSourceCopyFolder request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopFastTransferSourceCopyFolder request</param>
        private void VerifyRopFastTransferSourceCopyFolderResponse(RopFastTransferSourceCopyFolderResponse ropFastTransferSourceCopyFolderResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3769");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3769
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyFolderResponse.RopId.GetType(),
                3769,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3771");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3771
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferSourceCopyFolder,
                ropFastTransferSourceCopyFolderResponse.RopId,
                3771,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer, RopId (1 byte):] For this operation[RopFastTransferSourceCopyFolder] this field is set to 0x4C.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3772");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3772
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyFolderResponse.OutputHandleIndex.GetType(),
                3772,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3773");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3773
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropFastTransferSourceCopyFolderResponse.OutputHandleIndex,
                3773,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3775");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3775
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferSourceCopyFolderResponse.ReturnValue.GetType(),
                3775,
                @"[In RopFastTransferSourceCopyFolder ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopFastTransferSourceCopyMessages
        /// <summary>
        /// Verify RopFastTransferSourceCopyMessages Response
        /// </summary>
        /// <param name="ropFastTransferSourceCopyMessagesResponse">The response of RopFastTransferSourceCopyMessages request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopFastTransferSourceCopyMessages request</param>
        private void VerifyRopFastTransferSourceCopyMessagesResponse(RopFastTransferSourceCopyMessagesResponse ropFastTransferSourceCopyMessagesResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3801");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3801
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyMessagesResponse.RopId.GetType(),
                3801,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3803");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3803
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferSourceCopyMessages,
                ropFastTransferSourceCopyMessagesResponse.RopId,
                3803,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer, RopId (1 byte):] For this operation[RopFastTransferSourceCopyMessages] this field is set to 0x4B.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3804");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3804
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyMessagesResponse.OutputHandleIndex.GetType(),
                3804,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3805");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3805
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropFastTransferSourceCopyMessagesResponse.OutputHandleIndex,
                3805,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the OutputHandleIndex specified in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3807");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3807
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferSourceCopyMessagesResponse.ReturnValue.GetType(),
                3807,
                @"[In RopFastTransferSourceCopyMessages ROP Response Buffer] ReturnValue (4 bytes): An unsigned  integer.");
        }
        #endregion

        #region RopFastTransferSourceCopyTo
        /// <summary>
        /// Verify RopFastTransferSourceCopyTo Response
        /// </summary>
        /// <param name="ropFastTransferSourceCopyToResponse">The response of RopFastTransferSourceCopyToResponse request</param>
        /// <param name="outputHandleIndex">The field of outputHandleIndex in RopFastTransferSourceCopyToResponse request</param>
        private void VerifyRopFastTransferSourceCopyToResponse(RopFastTransferSourceCopyToResponse ropFastTransferSourceCopyToResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3836");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3836
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyToResponse.RopId.GetType(),
                3836,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3838");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3838
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferSourceCopyTo,
                ropFastTransferSourceCopyToResponse.RopId,
                3838,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer, RopId (1 byte):] For this operation[RopFastTransferSourceCopyTo] this field is set to 0x4D.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3839");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3839
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyToResponse.OutputHandleIndex.GetType(),
                3839,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3840");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3840
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropFastTransferSourceCopyToResponse.OutputHandleIndex,
                3840,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3842");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3842
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferSourceCopyToResponse.ReturnValue.GetType(),
                3842,
                @"[In RopFastTransferSourceCopyTo ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopFastTransferSourceCopyProperties
        /// <summary>
        /// Verify RopFastTransferSourceCopyProperties Response
        /// </summary>
        /// <param name="ropFastTransferSourceCopyPropertiesResponse">The response of RopFastTransferSourceCopyProperties request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopFastTransferSourceCopyProperties request</param>
        private void VerifyRopFastTransferSourceCopyPropertiesResponse(
            RopFastTransferSourceCopyPropertiesResponse ropFastTransferSourceCopyPropertiesResponse,
            byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3871");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3871
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyPropertiesResponse.RopId.GetType(),
                3871,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3873");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3873
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopFastTransferSourceCopyProperties,
                ropFastTransferSourceCopyPropertiesResponse.RopId,
                3873,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer, RopId (1 byte):] For this operation[RopFastTransferSourceCopyProperties] this field is set to 0x69.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3874");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3874
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropFastTransferSourceCopyPropertiesResponse.OutputHandleIndex.GetType(),
                3874,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3875");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3875
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropFastTransferSourceCopyPropertiesResponse.OutputHandleIndex,
                3875,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3877");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3877
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropFastTransferSourceCopyPropertiesResponse.ReturnValue.GetType(),
                3877,
                @"[In RopFastTransferSourceCopyProperties ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopTellVersion
        /// <summary>
        /// Verify RopTellVersion Response
        /// </summary>
        /// <param name="ropTellVersionResponse">the response of RopTellVersion request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopTellVersion request</param>
        private void VerifyRopTellVersionResponse(RopTellVersionResponse ropTellVersionResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3892");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3892
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTellVersionResponse.RopId.GetType(),
                3892,
                @"[In RopTellVersion ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3894");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3894
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopTellVersion,
                ropTellVersionResponse.RopId,
                3894,
                @"[In RopTellVersion ROP Response Buffer, RopId (1 byte):] For this operation[RopTellVersion] this field is set to 0x86.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3895");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3895
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropTellVersionResponse.InputHandleIndex.GetType(),
                3895,
                @"[In RopTellVersion ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3896");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3896
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropTellVersionResponse.InputHandleIndex,
                3896,
                @"[In RopTellVersion ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the vale specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3898");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3898
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropTellVersionResponse.ReturnValue.GetType(),
                3898,
                @"[In RopTellVersion ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Incremental Change Synchronization ROPs

        #region RopSynchronizationConfigure
        /// <summary>
        /// Verify RopSynchronizationConfigure Response
        /// </summary>
        /// <param name="ropSynchronizationConfigureResponse">The response of RopSynchronizationConfigure request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopSynchronizationConfigure request</param>
        private void VerifyRopSynchronizationConfigureResponse(RopSynchronizationConfigureResponse ropSynchronizationConfigureResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3937");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3937
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationConfigureResponse.RopId.GetType(),
                3937,
                @"[In RopSynchronizationConfigure ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3939");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3939
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationConfigure,
                ropSynchronizationConfigureResponse.RopId,
                3939,
                @"[In RopSynchronizationConfigure ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationConfigure] this field is set to 0x70.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3940");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3940
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationConfigureResponse.OutputHandleIndex.GetType(),
                3940,
                @"[In RopSynchronizationConfigure ROP Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3941");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3941
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropSynchronizationConfigureResponse.OutputHandleIndex,
                3941,
                @"[In RopSynchronizationConfigure ROP Response Buffer, OutputHandleIndex (1 byte):] This Index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3943");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3943
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationConfigureResponse.ReturnValue.GetType(),
                3943,
                @"[In RopSynchronizationConfigure ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationImportMessageChange
        /// <summary>
        /// Verify RopSynchronizationImportMessageChange Success Response
        /// </summary>
        /// <param name="ropSynchronizationImportMessageChangeResponse">The success response of RopSynchronizationImportMessageChange request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopSynchronizationImportMessageChange request</param>
        private void VerifyRopSynchronizationImportMessageChangeSuccessResponse(
            RopSynchronizationImportMessageChangeResponse ropSynchronizationImportMessageChangeResponse,
            byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3969");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3969
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageChangeResponse.RopId.GetType(),
                3969,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3971");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3971
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportMessageChange,
                ropSynchronizationImportMessageChangeResponse.RopId,
                3971,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportMessageChange] this field is set to 0x72.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3972");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3972
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageChangeResponse.OutputHandleIndex.GetType(),
                3972,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3973");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3973
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropSynchronizationImportMessageChangeResponse.OutputHandleIndex,
                3973,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3975");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3975
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportMessageChangeResponse.ReturnValue.GetType(),
                3975,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3977");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3977
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportMessageChangeResponse.ReturnValue,
                3977,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer, ReturnValue (4 bytes):] For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3978");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3978
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropSynchronizationImportMessageChangeResponse.MessageId.GetType(),
                3978,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer] MessageId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4694");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4694
            Site.CaptureRequirementIfAreEqual<ulong?>(
                MessageIdForRops,
                ropSynchronizationImportMessageChangeResponse.MessageId,
                4694,
                @"[In RopSynchronizationImportMessageChange ROP Success Response Buffer] MessageId (8 bytes): This field MUST be set to 0x0000000000000000.");
        }

        /// <summary>
        /// Verify RopSynchronizationImportMessageChange Failure Response
        /// </summary>
        /// <param name="ropSynchronizationImportMessageChangeResponse">The failure response of RopSynchronizationImportMessageChange request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopSynchronizationImportMessageChange request</param>
        private void VerifyRopSynchronizationImportMessageChangeFailureResponse(
            RopSynchronizationImportMessageChangeResponse ropSynchronizationImportMessageChangeResponse,
            byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3980");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3980
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageChangeResponse.RopId.GetType(),
                3980,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3982");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3982
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportMessageChange,
                ropSynchronizationImportMessageChangeResponse.RopId,
                3982,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportMessageChange] this field is set to 0x72.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3983");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3983
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageChangeResponse.OutputHandleIndex.GetType(),
                3983,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer] OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3984");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3984
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropSynchronizationImportMessageChangeResponse.OutputHandleIndex,
                3984,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3986");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3986
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportMessageChangeResponse.ReturnValue.GetType(),
                3986,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R3988");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R3988
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportMessageChangeResponse.ReturnValue,
                3988,
                @"[In RopSynchronizationImportMessageChange ROP Failure Response Buffer, ReturnValue (4 bytes):] For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSynchronizationImportReadStateChanges
        /// <summary>
        /// Verify RopSynchronizationImportReadStateChanges Response
        /// </summary>
        /// <param name="ropSynchronizationImportReadStateChangesResponse">The response of RopSynchronizationImportReadStateChanges request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopSynchronizationImportReadStateChanges request</param>
        private void VerifyRopSynchronizationImportReadStateChangesResponse(
            RopSynchronizationImportReadStateChangesResponse ropSynchronizationImportReadStateChangesResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4013");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4013
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportReadStateChangesResponse.RopId.GetType(),
                4013,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4015");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4015
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportReadStateChanges,
                ropSynchronizationImportReadStateChangesResponse.RopId,
                4015,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportReadStateChanges] this field is set to 0x80.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4016");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4016
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportReadStateChangesResponse.InputHandleIndex.GetType(),
                4016,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4017");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4017
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportReadStateChangesResponse.InputHandleIndex,
                4017,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4019");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4019
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportReadStateChangesResponse.ReturnValue.GetType(),
                4019,
                @"[In RopSynchronizationImportReadStateChanges ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationImportHierarchyChange
        /// <summary>
        /// Verify RopSynchronizationImportHierarchyChange Success Response
        /// </summary>
        /// <param name="ropSynchronizationImportHierarchyChangeResponse">The success response of RopSynchronizationImportHierarchyChange request</param>
        /// <param name="inputHandleIndex">The field of inputHandleIndex in RopSynchronizationImportHierarchyChange request</param>
        private void VerifyRopSynchronizationImportHierarchyChangeSuccessResponse(
            RopSynchronizationImportHierarchyChangeResponse ropSynchronizationImportHierarchyChangeResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4043");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4043
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportHierarchyChangeResponse.RopId.GetType(),
                4043,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4045");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4045
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportHierarchyChange,
                ropSynchronizationImportHierarchyChangeResponse.RopId,
                4045,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportHierarchyChange] this field is set to 0x73.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4046");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4046
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportHierarchyChangeResponse.InputHandleIndex.GetType(),
                4046,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4047");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4047
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportHierarchyChangeResponse.InputHandleIndex,
                4047,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4049");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4049
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportHierarchyChangeResponse.ReturnValue.GetType(),
                4049,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4051");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4051
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportHierarchyChangeResponse.ReturnValue,
                4051,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer, ReturnValue (4 bytes):] For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4052");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4052
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropSynchronizationImportHierarchyChangeResponse.FolderId.GetType(),
                4052,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer] FolderId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4695");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4695
            Site.CaptureRequirementIfAreEqual<ulong>(
                FolderIdForRopSynchronizationImportHierarchyChange,
                ropSynchronizationImportHierarchyChangeResponse.FolderId,
                4695,
                @"[In RopSynchronizationImportHierarchyChange ROP Success Response Buffer] FolderId (8 bytes): This field MUST be set to 0x0000000000000000.");
        }

        /// <summary>
        /// Verify RopSynchronizationImportHierarchyChange Failure Response
        /// </summary>
        /// <param name="ropSynchronizationImportHierarchyChangeResponse">The failure response of RopSynchronizationImportHierarchyChange request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationImportHierarchyChange request</param>
        private void VerifyRopSynchronizationImportHierarchyChangeFailureResponse(
            RopSynchronizationImportHierarchyChangeResponse ropSynchronizationImportHierarchyChangeResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4054");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4054
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportHierarchyChangeResponse.RopId.GetType(),
                4054,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4056");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4056
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportHierarchyChange,
                ropSynchronizationImportHierarchyChangeResponse.RopId,
                4056,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportHierarchyChange] this field is set to 0x73.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4057");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4057
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportHierarchyChangeResponse.InputHandleIndex.GetType(),
                4057,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4058");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4058
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportHierarchyChangeResponse.InputHandleIndex,
                4058,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4060");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4060
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportHierarchyChangeResponse.ReturnValue.GetType(),
                4060,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4062");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4062
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportHierarchyChangeResponse.ReturnValue,
                4062,
                @"[In RopSynchronizationImportHierarchyChange ROP Failure Response Buffer, ReturnValue (4 bytes):] For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSynchronizationImportDeletes
        /// <summary>
        /// Verify RopSynchronizationImportDeletes Response
        /// </summary>
        /// <param name="ropSynchronizationImportDeletesResponse">The response of RopSynchronizationImportDeletes request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationImportDeletes request</param>
        private void VerifyRopSynchronizationImportDeletesResponse(RopSynchronizationImportDeletesResponse ropSynchronizationImportDeletesResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4083");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4083
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportDeletesResponse.RopId.GetType(),
                4083,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4085");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4085
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportDeletes,
                ropSynchronizationImportDeletesResponse.RopId,
                4085,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportDeletes] this field is set to 0x74.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4086");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4086
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportDeletesResponse.InputHandleIndex.GetType(),
                4086,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer]InputHandleIndex (1 byte): An unsigned  integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4087");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4087
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportDeletesResponse.InputHandleIndex,
                4087,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4089");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4089
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportDeletesResponse.ReturnValue.GetType(),
                4089,
                @"[In RopSynchronizationImportDeletes ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationImportMessageMove
        /// <summary>
        /// Verify RopSynchronizationImportMessageMove Success Response
        /// </summary>
        /// <param name="ropSynchronizationImportMessageMoveResponse">The success response of RopSynchronizationImportMessageMove request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationImportMessageMove request</param>
        private void VerifyRopSynchronizationImportMessageMoveSuccessResponse(
            RopSynchronizationImportMessageMoveResponse ropSynchronizationImportMessageMoveResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4125");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4125
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageMoveResponse.RopId.GetType(),
                4125,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4127");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4127
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportMessageMove,
                ropSynchronizationImportMessageMoveResponse.RopId,
                4127,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportMessageMove] this field is set to 0x78.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4128");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4128
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageMoveResponse.InputHandleIndex.GetType(),
                4128,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4129");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4129
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportMessageMoveResponse.InputHandleIndex,
                4129,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4131");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4131
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportMessageMoveResponse.ReturnValue.GetType(),
                4131,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4133");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4133
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportMessageMoveResponse.ReturnValue,
                4133,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer, ReturnValue (4 bytes):] For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4134");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4134
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ulong),
                ropSynchronizationImportMessageMoveResponse.MessageId.GetType(),
                4134,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer] MessageId (8 bytes): An identifier.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4696");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4696
            Site.CaptureRequirementIfAreEqual<ulong?>(
                MessageIdForRops,
                 ropSynchronizationImportMessageMoveResponse.MessageId,
                4696,
                @"[In RopSynchronizationImportMessageMove ROP Success Response Buffer] MessageId (8 bytes): This field MUST be set to 0x0000000000000000.");
        }

        /// <summary>
        /// Verify RopSynchronizationImportMessageMove Failure Response
        /// </summary>
        /// <param name="ropSynchronizationImportMessageMoveResponse">The failure response of RopSynchronizationImportMessageMove request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationImportMessageMove request</param>
        private void VerifyRopSynchronizationImportMessageMoveFailureResponse(
            RopSynchronizationImportMessageMoveResponse ropSynchronizationImportMessageMoveResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4136");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4136
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageMoveResponse.RopId.GetType(),
                4136,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer] RopId (1 byte):  An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4138");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4138
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationImportMessageMove,
                ropSynchronizationImportMessageMoveResponse.RopId,
                4138,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationImportMessageMove] this field is set to 0x78.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4139");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4139
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationImportMessageMoveResponse.InputHandleIndex.GetType(),
                4139,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4140");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4140
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationImportMessageMoveResponse.InputHandleIndex,
                4140,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4142");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4142
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationImportMessageMoveResponse.ReturnValue.GetType(),
                4142,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4144");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4144
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropSynchronizationImportMessageMoveResponse.ReturnValue,
                4144,
                @"[In RopSynchronizationImportMessageMove ROP Failure Response Buffer, ReturnValue (4 bytes):] For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #region RopSynchronizationOpenCollector
        /// <summary>
        /// Verify RopSynchronizationOpenCollector Response
        /// </summary>
        /// <param name="ropSynchronizationOpenCollectorResponse">The response of RopSynchronizationOpenCollector request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopSynchronizationOpenCollector request</param>
        private void VerifyRopSynchronizationOpenCollectorResponse(RopSynchronizationOpenCollectorResponse ropSynchronizationOpenCollectorResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4160");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4160
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationOpenCollectorResponse.RopId.GetType(),
                4160,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4162");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4162
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationOpenCollector,
                ropSynchronizationOpenCollectorResponse.RopId,
                4162,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationOpenCollector] this field is set to 0x7E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4163");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4163
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationOpenCollectorResponse.OutputHandleIndex.GetType(),
                4163,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4164");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4164
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropSynchronizationOpenCollectorResponse.OutputHandleIndex,
                4164,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4166");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4166
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationOpenCollectorResponse.ReturnValue.GetType(),
                4166,
                @"[In RopSynchronizationOpenCollector ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationGetTransferState
        /// <summary>
        /// Verify RopSynchronizationGetTransferState Response
        /// </summary>
        /// <param name="ropSynchronizationGetTransferStateResponse">The response of RopSynchronizationGetTransferState request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopSynchronizationGetTransferState request</param>
        private void VerifyRopSynchronizationGetTransferStateResponse(RopSynchronizationGetTransferStateResponse ropSynchronizationGetTransferStateResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4181");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4181
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationGetTransferStateResponse.RopId.GetType(),
                4181,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4183");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4183
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationGetTransferState,
                ropSynchronizationGetTransferStateResponse.RopId,
                4183,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationGetTransferState] this field is set to 0x82.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4184");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4184
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationGetTransferStateResponse.OutputHandleIndex.GetType(),
                4184,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4185");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4185
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropSynchronizationGetTransferStateResponse.OutputHandleIndex,
                4185,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer, OutputHandleIndex (1 byte):]This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4187");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4187
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationGetTransferStateResponse.ReturnValue.GetType(),
                4187,
                @"[In RopSynchronizationGetTransferState ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationUploadStateStreamBegin
        /// <summary>
        /// Verify RopSynchronizationUploadStateStreamBegin Response
        /// </summary>
        /// <param name="ropSynchronizationUploadStateStreamBeginResponse">The response of RopSynchronizationUploadStateStreamBegin request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationUploadStateStreamBegin request</param>
        private void VerifyRopSynchronizationUploadStateStreamBeginResponse(
            RopSynchronizationUploadStateStreamBeginResponse ropSynchronizationUploadStateStreamBeginResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4203");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4203
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamBeginResponse.RopId.GetType(),
                4203,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4205");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4205
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationUploadStateStreamBegin,
                ropSynchronizationUploadStateStreamBeginResponse.RopId,
                4205,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationUploadStateStreamBegin] this field is set to 0x75.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4206");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4206
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamBeginResponse.InputHandleIndex.GetType(),
                4206,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4207");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4207
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationUploadStateStreamBeginResponse.InputHandleIndex,
                4207,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4209");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4209
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationUploadStateStreamBeginResponse.ReturnValue.GetType(),
                4209,
                @"[In RopSynchronizationUploadStateStreamBegin ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationUploadStateStreamContinue
        /// <summary>
        /// Verify RopSynchronizationUploadStateStreamContinue Response
        /// </summary>
        /// <param name="ropSynchronizationUploadStateStreamContinueResponse">The response of RopSynchronizationUploadStateStreamContinue request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationUploadStateStreamContinue request</param>
        private void VerifyRopSynchronizationUploadStateStreamContinueResponse(
            RopSynchronizationUploadStateStreamContinueResponse ropSynchronizationUploadStateStreamContinueResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4226");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4226
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamContinueResponse.RopId.GetType(),
                4226,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4228");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4228
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationUploadStateStreamContinue,
                ropSynchronizationUploadStateStreamContinueResponse.RopId,
                4228,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer,RopId (1 byte):] For this operation[RopSynchronizationUploadStateStreamContinue] this field is set to 0x76.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4229");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4229
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamContinueResponse.InputHandleIndex.GetType(),
                4229,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4230");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4230
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationUploadStateStreamContinueResponse.InputHandleIndex,
                4230,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer, InputHandleIndex (1 byte):]This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4232");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4232
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationUploadStateStreamContinueResponse.ReturnValue.GetType(),
                4232,
                @"[In RopSynchronizationUploadStateStreamContinue ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSynchronizationUploadStateStreamEnd
        /// <summary>
        /// Verify RopSynchronizationUploadStateStreamEnd Response
        /// </summary>
        /// <param name="ropSynchronizationUploadStateStreamEndResponse">The response of RopSynchronizationUploadStateStreamEnd request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSynchronizationUploadStateStreamEnd request</param>
        private void VerifyRopSynchronizationUploadStateStreamEndResponse(
            RopSynchronizationUploadStateStreamEndResponse ropSynchronizationUploadStateStreamEndResponse,
            byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4244");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4244
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamEndResponse.RopId.GetType(),
                4244,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4246");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4246
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSynchronizationUploadStateStreamEnd,
                ropSynchronizationUploadStateStreamEndResponse.RopId,
                4246,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer, RopId (1 byte):] For this operation[RopSynchronizationUploadStateStreamEnd] this field is set to 0x77.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4247");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4247
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSynchronizationUploadStateStreamEndResponse.InputHandleIndex.GetType(),
                4247,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4248");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4248
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSynchronizationUploadStateStreamEndResponse.InputHandleIndex,
                4248,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4250");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4250
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSynchronizationUploadStateStreamEndResponse.ReturnValue.GetType(),
                4250,
                @"[In RopSynchronizationUploadStateStreamEnd ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopSetLocalReplicaMidsetDeleted
        /// <summary>
        /// Verify RopSetLocalReplicaMidsetDeleted Response
        /// </summary>
        /// <param name="ropSetLocalReplicaMidsetDeletedResponse">The response of RopSetLocalReplicaMidsetDeleted request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopSetLocalReplicaMidsetDeleted request</param>
        private void VerifyRopSetLocalReplicaMidsetDeletedResponse(RopSetLocalReplicaMidsetDeletedResponse ropSetLocalReplicaMidsetDeletedResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4276");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4276
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetLocalReplicaMidsetDeletedResponse.RopId.GetType(),
                4276,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4278");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4278
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopSetLocalReplicaMidsetDeleted,
                ropSetLocalReplicaMidsetDeletedResponse.RopId,
                4278,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer, RopId (1 byte):] For this operation[RopSetLocalReplicaMidsetDeleted] this field is set to 0x93.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4279");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4279
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropSetLocalReplicaMidsetDeletedResponse.InputHandleIndex.GetType(),
                4279,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer]InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4280");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4280
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropSetLocalReplicaMidsetDeletedResponse.InputHandleIndex,
                4280,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4282");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4282
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropSetLocalReplicaMidsetDeletedResponse.ReturnValue.GetType(),
                4282,
                @"[In RopSetLocalReplicaMidsetDeleted ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopGetLocalReplicaIds
        /// <summary>
        /// Verify RopGetLocalReplicaIds Success Response
        /// </summary>
        /// <param name="ropGetLocalReplicaIdsResponse">The success response of RopGetLocalReplicaIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetLocalReplicaIds request</param>
        private void VerifyRopGetLocalReplicaIdsSuccessResponse(RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4296");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4296
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetLocalReplicaIdsResponse.RopId.GetType(),
                4296,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4298");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4298
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetLocalReplicaIds,
                ropGetLocalReplicaIdsResponse.RopId,
                4298,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer, RopId (1 byte):] For this operation[RopGetLocalReplicaIds], this field is set to 0x7F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4299");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4299
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetLocalReplicaIdsResponse.InputHandleIndex.GetType(),
                4299,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4300");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4300
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetLocalReplicaIdsResponse.InputHandleIndex,
                4300,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in the InputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4303");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4303
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetLocalReplicaIdsResponse.ReturnValue.GetType(),
                4303,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4305");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4305
            Site.CaptureRequirementIfAreEqual<uint>(
                SuccessReturnValue,
                ropGetLocalReplicaIdsResponse.ReturnValue,
                4305,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer, ReturnValue (4 bytes):] For this response[Success Response], this field is set to 0x00000000.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4306, actual type of ReplGuid is :{0}", ropGetLocalReplicaIdsResponse.ReplGuid.GetType().Name);

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4306
            bool isVerifyR4306 = this.IsGUID(ropGetLocalReplicaIdsResponse.ReplGuid);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR4306,
                4306,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer] ReplGuid (16 bytes): GUID.");

            // Add the debug information
            Site.Log.Add(
                LogEntryKind.Debug,
                "Verify MS-OXCROPS_R4309,The size of bytes is:{0},The type of GlobalCount is:{0}",
                ropGetLocalReplicaIdsResponse.GlobalCount.Length,
                ropGetLocalReplicaIdsResponse.GetType().ToString());

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4309
            bool isVerifyR4309 =
                (ropGetLocalReplicaIdsResponse.GlobalCount.GetType() == typeof(byte[])) &&
                (ropGetLocalReplicaIdsResponse.GlobalCount.Length == 6);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR4309,
                4309,
                @"[In RopGetLocalReplicaIds ROP Success Response Buffer] GlobalCount (6 bytes): An array of bytes.");
        }

        /// <summary>
        /// Verify RopGetLocalReplicaIds Failure Response
        /// </summary>
        /// <param name="ropGetLocalReplicaIdsResponse">The failure response of RopGetLocalReplicaIds request</param>
        /// <param name="inputHandleIndex">The field of InputHandleIndex in RopGetLocalReplicaIds request</param>
        private void VerifyRopGetLocalReplicaIdsFailureResponse(RopGetLocalReplicaIdsResponse ropGetLocalReplicaIdsResponse, byte inputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4311");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4311
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetLocalReplicaIdsResponse.RopId.GetType(),
                4311,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4313");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4313
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopGetLocalReplicaIds,
                ropGetLocalReplicaIdsResponse.RopId,
                4313,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer, RopId (1 byte):] For this operation[RopGetLocalReplicaIds] this field is set to 0x7F.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4314");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4314
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropGetLocalReplicaIdsResponse.InputHandleIndex.GetType(),
                4314,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer] InputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4315");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4315
            Site.CaptureRequirementIfAreEqual<byte>(
                inputHandleIndex,
                ropGetLocalReplicaIdsResponse.InputHandleIndex,
                4315,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer, InputHandleIndex (1 byte):] This index MUST be set to the value specified in theInputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4317");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4317
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropGetLocalReplicaIdsResponse.ReturnValue.GetType(),
                4317,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer] ReturnValue (4 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4319");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4319
            Site.CaptureRequirementIfAreNotEqual<uint>(
                SuccessReturnValue,
                ropGetLocalReplicaIdsResponse.ReturnValue,
                4319,
                @"[In RopGetLocalReplicaIds ROP Failure Response Buffer, ReturnValue (4 bytes):] For this response[Failure Response], this field is set to a value other than 0x00000000.");
        }
        #endregion

        #endregion

        #region Notification ROPs

        #region RopRegisterNotification
        /// <summary>
        /// Verify RopRegisterNotification Response
        /// </summary>
        /// <param name="ropRegisterNotificationResponse">The response of RopRegisterNotification request</param>
        /// <param name="outputHandleIndex">The field of OutputHandleIndex in RopRegisterNotification request</param>
        private void VerifyRopRegisterNotificationResponse(RopRegisterNotificationResponse ropRegisterNotificationResponse, byte outputHandleIndex)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4346");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4346
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRegisterNotificationResponse.RopId.GetType(),
                4346,
                @"[In RopRegisterNotification ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4348");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4348
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopRegisterNotification,
                ropRegisterNotificationResponse.RopId,
                4348,
                @"[In RopRegisterNotification ROP Response Buffer, RopId (1 byte):] For this operation[RopRegisterNotification] this field is set to 0x29.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4349");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4349
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropRegisterNotificationResponse.OutputHandleIndex.GetType(),
                4349,
                @"[In RopRegisterNotification ROP Response Buffer]OutputHandleIndex (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4350");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4350
            Site.CaptureRequirementIfAreEqual<byte>(
                outputHandleIndex,
                ropRegisterNotificationResponse.OutputHandleIndex,
                4350,
                @"[In RopRegisterNotification ROP Response Buffer, OutputHandleIndex (1 byte):] This index MUST be set to the value specified in the OutputHandleIndex field in the request.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4352");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4352
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropRegisterNotificationResponse.ReturnValue.GetType(),
                4352,
                @"[In RopRegisterNotification ROP Response Buffer] ReturnValue (4 bytes): An unsigned integer.");
        }
        #endregion

        #region RopNotify
        /// <summary>
        /// Verify RopNotify Response
        /// </summary>
        /// <param name="ropNotifyResponse">The response of RopNotify request</param>
        private void VerifyRopNotifyResponse(RopNotifyResponse ropNotifyResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4423");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4423
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropNotifyResponse.RopId.GetType(),
                4423,
                @"[In RopNotify ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4425");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4425
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopNotify,
                ropNotifyResponse.RopId,
                4425,
                @"[In RopNotify ROP Response Buffer, RopId (1 byte):] For this operation[RopNotify] this field is set to 0x2A.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4426");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4426
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(uint),
                ropNotifyResponse.NotificationHandle.GetType(),
                4426,
                @"[In RopNotify ROP Response Buffer]NotificationHandle (4 bytes): A Server object handle.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4428");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4428
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropNotifyResponse.LogonId.GetType(),
                4428,
                @"[In RopNotify ROP Response Buffer] LogonId (1 byte): An unsigned integer.");
        }
        #endregion

        #region RopPending
        /// <summary>
        /// Verify RopPending Response
        /// </summary>
        /// <param name="ropPendingResponse">The response of RopPending request</param>
        private void VerifyRopPendingResponse(RopPendingResponse ropPendingResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4434");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4434
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropPendingResponse.RopId.GetType(),
                4434,
                @"[In RopPending ROP Response Buffer] RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4436");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4436
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopPending,
                ropPendingResponse.RopId,
                4436,
                @"[In RopPending ROP Response Buffer, RopId (1 byte):] For this operation[RopPending] this field is set to 0x6E.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4437");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4437
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropPendingResponse.SessionIndex.GetType(),
                4437,
                @"[In RopPending ROP Response Buffer] SessionIndex (2 bytes): An unsigned integer.");
        }
        #endregion

        #endregion

        #region Other ROPs

        #region RopBufferTooSmall
        /// <summary>
        /// Verify RopBufferTooSmall Response
        /// </summary>
        /// <param name="ropBufferTooSmallResponse">The response of RopBufferTooSmall request</param>
        private void VerifyRopBufferTooSmallResponse(RopBufferTooSmallResponse ropBufferTooSmallResponse)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4653");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4653
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte),
                ropBufferTooSmallResponse.RopId.GetType(),
                4653,
                @"[In RopBufferTooSmall ROP Response Buffer]RopId (1 byte): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4655");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4655
            Site.CaptureRequirementIfAreEqual<byte>(
                (byte)RopId.RopBufferTooSmall,
                ropBufferTooSmallResponse.RopId,
                4655,
                @"[In RopBufferTooSmall ROP Response Buffer, RopId (1 byte):] For this operation[RopBufferTooSmall] this field is set to 0xFF.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4443");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4443
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                ropBufferTooSmallResponse.SizeNeeded.GetType(),
                4443,
                @"[In RopBufferTooSmall ROP Response Buffer] SizeNeeded (2 bytes): An unsigned integer.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4445");

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4445
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(byte[]),
                ropBufferTooSmallResponse.RequestBuffers.GetType(),
                4445,
                @"[In RopBufferTooSmall ROP Response Buffer]RequestBuffers (variable): An array of bytes.");
        }
        #endregion

        #endregion

        #region Message Processing Events and Sequencing Rules
        /// <summary>
        /// Verify Message Processing Events and Sequencing Rules
        /// </summary>
        /// <param name="serverOHTable">Server object handle table that include the modified handle that server created for the server object</param>
        private void VerifyMessageProcessingEventsAndSequencingRules(List<uint> serverOHTable)
        {
            bool isNotEqual = true;
            for (int i = 0; i < serverOHTable.Count; i++)
            {
                if (i != serverOHTable.Count - 1)
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R4560, the actual server object handle is [{0}: {1}]", i, serverOHTable[i]);
                    
                    // The last index value 0xFFFFFFFF used to initialize unused entries 
                    // of a Server object handle table for multiple ROPS
                    if (serverOHTable[i] == HandleValueForUnusedEntries)
                    {
                        isNotEqual = false;
                        break;
                    }
                }
            }

            // Verify MS-OXCROPS requirement: MS-OXCROPS_R4560
            bool isVerifyR4560 = isNotEqual;

            Site.CaptureRequirementIfIsTrue(
                isVerifyR4560,
                4560,
                @"[In Processing a ROP Input Buffer] Because the Server object handle value 0xFFFFFFFF is used to initialize unused entries of a Server object handle table, a server MUST NOT assign that value to a created Server Object.");
        }
        #endregion

        #region  Insufficient Output Buffer Space

        /// <summary>
        /// Verify RPC error when the pcbOut parameter of EcDoRpcExt2 is set to the maximum value.
        /// </summary>
        /// <param name="status">The status of the RPC operation.</param>
        private void VerifyFailRPCForMaxPcbOut(uint status)
        {
            if (Common.IsRequirementEnabled(454509, this.Site)
                 && (this.oxcropsClient.MapiContext.TransportSequence.ToLower().Equals("ncacn_ip_tcp", StringComparison.InvariantCultureIgnoreCase)
                 || this.oxcropsClient.MapiContext.TransportSequence.ToLower().Equals("ncacn_http", StringComparison.InvariantCultureIgnoreCase)))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R454509");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R454509
                Site.CaptureRequirementIfAreEqual<uint>(
                    OxcRpcErrorCode.ECResponseTooBig,
                    status,
                    454509,
                    @"[In Appendix B: Product Behavior] If one of the ROP responses will not fit in the ROP output buffer when either the pcbOut parameter of EcDoRpcExt2 response is set to the maximum value, then implementation does fail the EcDoRpcExt2 method with a return value of 0x0000047D. (Exchange 2010  and above follow this behavior.)");
            }

            if (Common.IsRequirementEnabled(20009, this.Site)
               && this.oxcropsClient.MapiContext.TransportSequence.ToLower().Equals("mapi_http", StringComparison.InvariantCultureIgnoreCase))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R20009");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R20009
                Site.CaptureRequirementIfAreEqual<uint>(
                    OxcRpcErrorCode.ECResponseTooBig,
                    status,
                    20009,
                    @"[In Appendix B: Product Behavior] If one of the ROP responses will not fit in the ROP output buffer when the RopBufferSize field of the Execute request type success response body, as specified in [MS-OXCMAPIHTTP] section 2.2.4.2.2, is set to the maximum value, then implementation does fail the Execute request type with a value of 0x0000047D in the StatusCode field. (Exchange 2013 SP1  and above follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify RPC error when the RPC buffer is with insufficient space.
        /// </summary>
        /// <param name="status">The status of the RPC operation.</param>
        private void VerifyFailRPCForInsufficientOutputBuffer(uint status)
        {
            if (Common.IsRequirementEnabled(469306, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R469306");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R469306
                Site.CaptureRequirementIfAreEqual<uint>(
                    OxcRpcErrorCode.ECBufferTooSmall,
                    status,
                    469306,
                    @"[In Appendix B: Product Behavior] If the first ROP in the RPC buffer creates a response that exceeds the maximum ROP buffer size and there are no additional ROPs in the RPC buffer, then the implementation does return the ecBufferTooSmall error to fail the ROP instead of returning the RopBufferTooSmall ROP (section 2.2.15.1). (Update Rollup 3 for Exchange 2010 SP2 follow this behavior.)");
            }
        }

        /// <summary>
        /// Verify RPC error when the field MaximumByteCount of RopReadStream request is greater than 0x80000000.
        /// </summary>
        /// <param name="status">The status of the RPC operation.</param>
        private void VerifyMaximumByteCountExceedError(uint status)
        {
            if (Common.IsRequirementEnabled(321501, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCROPS_R321501");

                // Verify MS-OXCROPS requirement: MS-OXCROPS_R321501
                Site.CaptureRequirementIfAreEqual<uint>(
                    OxcRpcErrorCode.ECRpcFormat,
                    status,
                    321501,
                    @"[In Appendix B: Product Behavior] If MaximumByteCount is greater than 0x80000000, then Implementation does fail with error code 0x000004B6. (Microsoft Exchange Server 2007 and above follow this behavior.)");
            }
        }

        #endregion
    }
}