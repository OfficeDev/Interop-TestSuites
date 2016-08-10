namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// Server role and both role Adapter requirements capture code for MS-OXCRPC.
    /// </summary>
    public partial class MS_OXCRPCAdapter : ManagedAdapterBase, IMS_OXCRPCAdapter
    {
        #region Verify Simple Data

        /// <summary>
        /// Verify the BIG_RANGE_ULONG simple data structure
        /// </summary>
        /// <param name="bigRangeUlong">An unsigned long that MUST be between 0x0 and 0x40000</param>
        private void VerifySimpleDataBig(ulong bigRangeUlong)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R20, BIG_RANGE_ULONG value is: {0}", bigRangeUlong);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R20
            bool isVerifyR20 = (bigRangeUlong >= 0x0) && (bigRangeUlong <= 0x40000);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR20,
                20,
                @"[In BIG_RANGE_ULONG Data Type] The BIG_RANGE_ULONG data type is an unsigned long that MUST be between 0x0 and 0x40000.");
        }

        /// <summary>
        /// Verify the SMALL_RANGE_ULONG simple data structure
        /// </summary>
        /// <param name="smallRangeUlong">An unsigned long that MUST be between 0x0 and 0x1008.</param>
        private void VerifySimpleDataSmall(ulong smallRangeUlong)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R21, SMALL_RANGE_ULONG value is: {0}", smallRangeUlong);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R21
            bool isVerifyR21 = (smallRangeUlong >= 0x0) &&
                (smallRangeUlong <= 0x1008);

            Site.CaptureRequirementIfIsTrue(
                isVerifyR21,
                21,
                @"[In SMALL_RANGE_ULONG Data Type] The SMALL_RANGE_ULONG data type is an unsigned long that MUST be between 0x0 and 0x1008.");
        }

        #endregion Verify Simple Data

        #region Verify Transport

        /// <summary>
        ///  Verify the referenced transport "ncacn_http" in this Open Specification.
        /// </summary>
        /// <param name="returnedBindingHandle"> A valid RPC binding handle </param>
        private void VerifyNcacnHttp(IntPtr returnedBindingHandle)
        {
            // IMS-OXCRPCAdapter interface implementation will initialize the client and server, build the transport tunnel between client and server.
            // If the returned RPC Binding is not intptr.zero, which indicates that initialize the client and server successfully, these requirements will be verified.
            if (returnedBindingHandle != IntPtr.Zero)
            {
                if (Common.IsRequirementEnabled(1936, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1936.The RPC binding handle is {0}.", returnedBindingHandle);

                    // If the type of the transport is ncacn_http and server returns RPC Binding which is not intptr.zero, this requirement will be verified.
                    Site.CaptureRequirement(
                        1936,
                        @"[In Appendix B: Product Behavior] Implementation does support this given protocol sequence [ncacn_http]. <1> Section 2.1: Exchange 2007 and above follow this behavior.");
                }

                // The well-known endpoint 6001 is a constant in the transport class used to bind to server when the transport sequence is ncacn_http. 
                // So if code can reach here, it indicates that the well-known endpoint 6001 is used for transport sequence ncacn_http, and R4 can be verified.
                this.Site.CaptureRequirement(
                    4,
                    @"[In Transport] For the network protocol sequence ncacn_http, this protocol MUST use the well-known endpoint 6001.");
            }
        }

        /// <summary>
        ///  Verify the referenced transport "ncacn_ip_tcp" in this protocol.
        /// </summary>
        /// <param name="returnedBindingHandle"> A valid RPC binding handle </param>
        private void VerifyNcacnIpTcp(IntPtr returnedBindingHandle)
        {
            // IMS-OXCRPCAdapter interface implementation will initialize the client and server, build the transport tunnel between client and server. 
            // If the returned RPC Binding is not intptr.zero, which indicates that initialize the client and server successfully, these requirements will be verified. 
            if (returnedBindingHandle != IntPtr.Zero)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R9, the RPC binding handle is {0}.", returnedBindingHandle);

                // Client will use the one RPC dynamic endpoint for protocol sequence "ncacn_ip_tcp", if the returned RPC Binding is not intptr.zero, 
                // which indicates that initialize the client and server successfully, so this requirement will be verified.
                Site.CaptureRequirement(
                    9,
                    @"[In Transport] For ncacn_ip_tcp, this protocol MUST use RPC dynamic endpoints, as defined in Part 4 of [C706].");

                if (Common.IsRequirementEnabled(1889, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1889, the RPC binding handle is {0}.", returnedBindingHandle);

                    // If the type of the transport is ncacn_ip_tcp and server returns RPC Binding which is not intptr.zero, this requirement will be verified.
                    Site.CaptureRequirement(
                        1889,
                        @"[In Appendix B: Product Behavior] Implementation does support this given protocol sequence [ncacn_ip_tcp]. ( Microsoft Exchange Server 2007 and 2010 follow this behavior.)");
                }
            }
        }

        /// <summary>
        ///  Verify UUID, little-endian byte order and identity of the caller in this Open Specification.
        /// </summary>
        /// <param name="bindingHandle"> A valid RPC binding handle </param>
        private void VerifyCommonRequirements(IntPtr bindingHandle)
        {
            if (bindingHandle != IntPtr.Zero)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R10, the RPC binding handle is {0}.", bindingHandle);

                // Client will use the UUID, if the returned RPC Binding is not intptr.zero, which indicates that initialize the client and server successfully, then these requirements will be verified. 
                Site.CaptureRequirement(
                    10,
                    @"[In Transport] This protocol MUST use the UUID specified in section 1.9.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1, the RPC binding handle is {0}.", bindingHandle);

                // If server returns RPC Binding which is not intptr.zero, it indicates that buffers and fields are depicted in little-endian byte order.
                Site.CaptureRequirement(
                    1,
                    @"[In Structures] Unless otherwise specified, buffers and fields in section 2.2.2.1 through section 2.2.2.2.22 are depicted in little-endian byte order.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R14, the RPC binding handle is {0}.", bindingHandle);

                // If server returns RPC Binding which is not intptr.zero, it indicates that the Open Specification uses the underlying RPC protocol to retrieve the identity of the caller that made the method call.
                Site.CaptureRequirement(
                    14,
                    @"[In Transport] The protocol [MS-OXCRPC] uses the underlying RPC protocol to retrieve the identity of the caller that made the method call, as specified in [MS-RPCE].");
            }
        }

        #endregion Verify Transport

        #region Verify response buffer

        /// <summary>
        /// Verify the RPC header structure.
        /// </summary>
        /// <param name="flagValues">All flag values before decompressing.</param>
        /// <param name="header">The structure to store information which is used to determine how to interpret the data following the header.</param>
        /// <param name="isExOrExt2">A Boolean value, true means use method EcDoConnectEx, false means use method EcDoRpcExt2</param>
        /// <param name="isAuxOutOrOut">A Boolean value, true means field is rgbAuxOut, false means field is rgbOut</param>
        /// <param name="isLast">A Boolean value, true means the last header, otherwise is false</param>
        private void VerifyRPCHeaderExt(
            List<short> flagValues,
            RPC_HEADER_EXT header,
            bool isExOrExt2,
            bool isAuxOutOrOut,
            bool isLast)
        {
            #region Verify whether rgbAuxOut includes an RPC_HEADER_EXT

            // Use method EcDoConnectEx, output buffer is rgbAuxOut
            if (isExOrExt2 && isAuxOutOrOut)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R575");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R575
                // According to the Open Specification, RPC_HEADER_EXT can be determined by the version field. The version of the header MUST be set to 0x0000
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    header.Version,
                    575,
                    @"[In EcDoConnectEx Method (Opnum 10)] [rgbAuxOut] The server MUST include an RPC_HEADER_EXT structure before the auxiliary payload data.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R747");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R747
                // According to the Open Specification, RPC_HEADER_EXT can be determined by the version field.
                // The version of the header MUST be set to 0x0000
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    header.Version,
                    747,
                    @"[In rgbAuxOut Output Buffer] The rgbAuxOut parameter output buffer contains an RPC_HEADER_EXT structure, as specified in section 2.2.2.1, followed by payload data.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R748");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R748
                // According to the Open Specification, Flag field is from third byte in RPC_HEADER_EXT. Last flag value is 0x0004
                Site.CaptureRequirementIfAreEqual<int>(
                    0x0004,
                    header.Flags & 0x0004,
                    748,
                    @"[In rgbAuxOut Output Buffer] The RPC_HEADER_EXT structure MUST contain the Last flag in the Flags field.");
            }

            // Use method EcDoRpcExt2, output buffer is rgbAuxOut
            if ((!isExOrExt2) && isAuxOutOrOut)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1649");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1649
                // Version field of RPC_HEADER_EXT is 0x0000 indicates that the output buffer rgbAuxOut has an RPC_HEADER_EXT structure which is the same as that of the rgbAuxOut input buffer in the EcDoConnectEx method.
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    header.Version,
                    1649,
                    @"[In rgbAuxOut Output Buffer] The format of the rgbAuxOut parameter input buffer for the EcDoRpcExt2 method, as specified in section 3.1.4.2, is the same as that of the rgbAuxOut parameter input buffer for the EcDoConnectEx method, as specified in section 3.1.4.1.1.1.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R682");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R682
                // According to the Open Specification, RPC_HEADER_EXT can be determined by the version field.
                // The version of the header MUST be set to 0x0000
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    header.Version,
                    682,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [rgbAuxOut] The server MUST include a RPC_HEADER_EXT header before the auxiliary payload data.");
            }

            // Use method EcDoRpcExt2, output buffer is rgbOut
            if ((!isExOrExt2) && (!isAuxOutOrOut))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R658");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R658
                // According to the Open Specification, RPC_HEADER_EXT can be determined by the version field.
                // The version of the header MUST be set to 0x0000
                Site.CaptureRequirementIfAreEqual<ushort>(
                    0x0000,
                    header.Version,
                    658,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [rgbOut] Like the ROP request payload, the ROP response payload is also prefixed by a RPC_HEADER_EXT header.");
            }
            #endregion Verify whether rgbAuxOut includes an RPC_HEADER_EXT

            #region Verify Version field

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R33");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R33
            Site.CaptureRequirementIfAreEqual<ushort>(
                0x0000,
                header.Version,
                33,
                @"[In RPC_HEADER_EXT Structure] [Version (2 bytes)] This value MUST be set to 0x0000.");

            #endregion Verify Version field

            #region Verify Flags field
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R34");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R34
            // If the Flags is an unsigned short value, then it must be 2 bytes, and this requirement will be verified.
            Site.CaptureRequirementIfAreEqual<Type>(
                typeof(ushort),
                header.Flags.GetType(),
                34,
                @"[In RPC_HEADER_EXT Structure] Flags (2 bytes): The flags that specify how data that follows this header [RPC_HEADER_EXT] MUST be interpreted.");

            if (header.Flags != 0x0000)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R35, the value of Flags buffer is:{0}.", header.Flags);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R35
                bool isVerifyR35 = ((header.Flags & 0x0001) == 0x0001) ||
                    ((header.Flags & 0x0002) == 0x0002) ||
                    ((header.Flags & 0x0004) == 0x0004);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR35,
                    35,
                    @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] The flags [Compressed, XorMagic, Last] in the following table are valid.");
            }

            // Flag values used here is the values before decompressing. All flag values in header.Flags are the values after decompressing 
            // and 0x0001 which indicates compressed data will no longer exists.
            foreach (short flag in flagValues)
            {
                if ((flag & 0x0001) == 0x0001)
                {
                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R4794
                    // Compared with SizeActual, the value of Size field doesn't include the length of RPC_HEADER_EXT structure according to the Open Specification.
                    // And the value of Size field is updated to the value of SizeActual when decompressing. So if code can reach here, R4794 can be verified directly.
                    this.Site.CaptureRequirement(
                        4794,
                        @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] [Compressed 0x0001] If this flag is set, the value of the Size field MUST be less than the value of the SizeActual field.");
                }
            }

            // Use method EcDoRpcExt2, output buffer is rgbOut, except last header
            if ((!isExOrExt2) && (!isAuxOutOrOut) && (!isLast))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R778");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R778
                // According to Open Specification, rgbOut is only on the method EcDoRpcExt2
                // Last flag value is 0x0004
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0x0004,
                    header.Flags & 0x0004,
                    778,
                    @"[In rgbOut Output Buffer] All RPC_HEADER_EXT structures in the output buffer except for the last MUST NOT contain the Last flag in the Flags field of the RPC_HEADER_EXT structure.");
            }

            // Use method EcDoRpcExt2, output buffer is rgbOut, header is last header
            if (!isExOrExt2 && !isAuxOutOrOut && isLast)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R779");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R779
                // According to Open Specification, rgbOut is only on the method EcDoRpcExt2
                // Last flag value is 0x0004
                Site.CaptureRequirementIfAreEqual<int>(
                    0x0004,
                    header.Flags & 0x0004,
                    779,
                    @"[In rgbOut Output Buffer] The last RPC_HEADER_EXT structure in the output buffer MUST contain the Last flag in its Flags field.");
            }

            if ((header.Flags & (ushort)RpcHeaderExtFlags.Last) == 0x0004)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R40, the value of Flags buffer is:{0}.", header.Flags);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R40
                Site.CaptureRequirementIfIsTrue(
                    isLast,
                    40,
                    @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] [Last 0x0004] No other RPC_HEADER_EXT structure follows the data of the current RPC_HEADER_EXT structure.");
            }

            #endregion Verify Flags field

            #region Verify Size and SizeActual fields

            // Verify the Compressed (0x0001) bit is not set in the Flags field.
            if ((header.Flags & 0x0001) != 0x0001)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R38");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R38
                Site.CaptureRequirementIfAreEqual<ushort>(
                    header.Size,
                    header.SizeActual,
                    38,
                    @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] [Compressed 0x0001] If this flag is not set, the Size and SizeActual fields MUST be the same.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R46");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R46
                Site.CaptureRequirementIfAreEqual<ushort>(
                    header.Size,
                    header.SizeActual,
                    46,
                    @"[In RPC_HEADER_EXT Structure] [SizeActual (2 bytes)] If the Compressed flag is not set, this value [SizeActual] MUST be equal to the value of the Size field.");
            }
            #endregion Verify Size and SizeActual fields
        }

        /// <summary>
        /// Verify multiple RPC_HEADER_EXTs in the rgbOut buffer
        /// </summary>
        /// <param name="headerCount">The count of RPC_HEADER_EXT</param>
        private void VerifyMultiRPCHeader(int headerCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R969, Header count value is: {0}", headerCount);

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R969
            // When rgbOut is parsed successfully through method ParseResponseBuffer, the count of payloads can be represented by the count of header in the payload.
            Site.CaptureRequirementIfIsTrue(
                headerCount <= 96,
                969,
                @"[In Extended Buffer Packing] The server MUST NOT place more than 96 individual payloads into a single rgbOut parameter response.");

            // headerCount > 1 means that additional data exists
            if (headerCount > 1)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R775");

                // Because each RPC_HEADER_EXT is tested in the VerifyRPCHeaderExt, if the code can arrive here, it means that each buffer contains an RPC_HEADER_EXT.
                Site.CaptureRequirement(
                    775,
                    @"[In rgbOut Output Buffer] The rgbOut parameter output buffer can contain multiple extended buffers in a single output buffer. Each of the extended buffers contains an RPC_HEADER_EXT structure, as specified in section 2.2.2.1, followed by Payload data.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R965");

                // Because each RPC_HEADER_EXT is tested in the VerifyRPCHeaderExt, if the code can arrive here, it means that each buffer contains an RPC_HEADER_EXT.
                Site.CaptureRequirement(
                    965,
                    @"[In Extended Buffer Packing] The additional ROP response is placed into the rgbOut parameter buffer following the previous header and associated payload with its own RPC_HEADER_EXT structure.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R777");

                // Because each RPC_HEADER_EXT is tested in the VerifyRPCHeaderExt, if the code can arrive here, it means that each buffer contains an RPC_HEADER_EXT.
                Site.CaptureRequirement(
                    777,
                    @"[In rgbOut Output Buffer] Each of the extended buffers contains an RPC_HEADER_EXT structure, as specified in section 2.2.2.1, followed by Payload data.");
            }
        }

        /// <summary>
        /// Verify the auxiliary buffer returned by calling the method EcDoConnectEx
        /// </summary>
        /// <param name="rgbAuxOutValue">Each item in this list contains the server topology information.</param>
        private void VerifyRgbAuxOutPayLoadOnEcDoConnectEx(List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R581");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R581
            // When pcbAuxOut is not 0(i.e., rgbAuxOut is not null) and the returned rgbAuxOutValue from method ParserRgbAuxOut is not null, 
            // the format of rgbAuxOut is consistent with Open Specification.
            // rgbAuxOut is determined by pcbAuxOut according to the implementation of ParserRgbAuxOut method, 
            // so this requirement will be covered if rgbAuxOutValue count is not 0(i.e., this method succeeds).
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                581,
                @"[In EcDoConnectEx Method (Opnum 10)] [pcbAuxOut] On output, this parameter contains the size of the data to be returned in the rgbAuxOut parameter.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R574");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R574
            // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                574,
                @"[In EcDoConnectEx Method (Opnum 10)] rgbAuxOut: On output, the server can return auxiliary payload data to the client in this parameter.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1061");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1061
            // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                1061,
                @"[In Server Topology Information] Each of these auxiliary blocks [AUX_CLIENT_CONTROL, AUX_OSVERSIONINFO, AUX_EXORGINFO and AUX_SERVER_SESSION_INFO] MUST be preceded by a properly formatted AUX_HEADER structure, as specified in section 2.2.2.2.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R990");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R990
            // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                990,
                @"[In Auxiliary Buffer] All information in the auxiliary buffer MUST be added with an AUX_HEADER structure preceding the actual auxiliary information.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1597");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1597
            // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                1597,
                @"[In Auxiliary Buffer] All information in the auxiliary buffer MUST be added with an AUX_HEADER structure preceding the actual auxiliary information.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R993");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R993
            // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                993,
                @"[In Auxiliary Buffer] The AUX_HEADER structure block contains the length of the AUX_HEADER structure itself plus the following auxiliary block structure in its Size field.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1060");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1060 "rgbAuxOutValue.Count! = 0" indicates that there are blocks in rgbAuxOut auxiliary buffer sent from the server to the client.
            Site.CaptureRequirementIfAreNotEqual<int>(
                0,
                rgbAuxOutValue.Count,
                1060,
                @"[In Server Topology Information] The auxiliary blocks [AUX_CLIENT_CONTROL, AUX_OSVERSIONINFO, AUX_EXORGINFO and AUX_EXORGINFO] sent from the server to the client in the rgbAuxOut parameter auxiliary buffer on the EcDoConnectEx method to provide server topology information are described in the following table.");

            // isAuxOsVersionInFo specifies whether block AUX_OSVERSIONINFO exists in the rgbAuxOut.
            bool isAuxOsVersionInFo = false;
            bool isAuxEndpointCapabilitiesExist = false;

            foreach (AUX_SERVER_TOPOLOGY_STRUCTURE rgbAuxOut1 in rgbAuxOutValue)
            {
                if (rgbAuxOut1.Header.Version == 0x01)
                {
                    // AUX_SERVER_SESSION_INFO, AUX_TYPE_ENDPOINT_CAPABILITIES, AUX_TYPE_EXORGINO, AUX_TYPE_OSVERSIONINFO and AUX_TYPE_CLIENT_CONTROL are blocks returned from server to client when server version is 0x01 according to the Open Specification.
                    if (rgbAuxOut1.Header.Type == 0x4B ||
                        rgbAuxOut1.Header.Type == 0x48 ||
                        rgbAuxOut1.Header.Type == 0x17 ||
                        rgbAuxOut1.Header.Type == 0x16 ||
                        rgbAuxOut1.Header.Type == 0x0A)
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R55, the type of payload data that follows the AUX_HEADER is {0}.", rgbAuxOut1.Header.Type);

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R55
                        // According to Open Specification, the type of AUX_HEADER should be one of specific values that are defined in Open Specification if the version of AUX_HEADER is 0x01.
                        // So R55 will be verified if the version of AUX_HEADER is 0x01 and the type of AUX_HEADER is one of the specific values.
                        Site.CaptureRequirement(
                            55,
                            @"[In AUX_HEADER Structure] [Type (1 byte)] The block type names, associated Type field values, and the corresponding auxiliary block structure that follows the AUX_HEADER structure when the Version field is AUX_VERSION_1 are listed in the following table [The types are: AUX_TYPE_PERF_REQUESTID 0x01, AUX_TYPE_PERF_CLIENTDINFO 0x02, AUX_TYPE_PERF_SERVERINFO 0x03, AUX_TYPE_PERF_SESSIONINFO 0x04, AUX_TYPE_PERF_DEFMDB_SUCCESS 0x05, AUX_TYPE_PERF_DEFGC_SUCCESS 0x06, AUX_TYPE_PERF_MDB_SUCCESS 0x07, AUX_TYPE_PERF_GC_SUCCESS 0x08, AUX_TYPE_PERF_FAILURE 0x09, AUX_TYPE_CLIENT_CONTROL 0x0A, AUX_TYPE_PERF_PROCESSINFO 0x0B, AUX_TYPE_PERF_BG_DEFMDB_SUCCESS 0x0C, AUX_TYPE_PERF_BG_DEFGC_SUCCESS 0x0D, AUX_TYPE_PERF_BG_MDB_SUCCESS 0x0E, AUX_TYPE_PERF_BG_GC_SUCCESS 0x0F, AUX_TYPE_PERF_BG_FAILURE 0x10, AUX_TYPE_PERF_FG_DEFMDB_SUCCESS 0x11, AUX_TYPE_PERF_FG_DEFGC_SUCCESS 0x12, AUX_TYPE_PERF_FG_MDB_SUCCESS 0x13, AUX_TYPE_PERF_FG_GC_SUCCESS 0x14, AUX_TYPE_PERF_FG_FAILURE 0x15, AUX_TYPE_OSVERSIONINFO 0x16, AUX_TYPE_EXORGINO 0x17, AUX_TYPE_PERF_ACCOUNTINFO 0x18, AUX_TYPE_ENDPOINT_CAPABILITIES 0x48, AUX_CLIENT_CONNECTION_INFO 0x4A, AUX_SERVER_SESSION_INFO 0X4B, AUX_PROTOCOL_DEVICE_IDENTIFICATION 0X4E].");
                    }
                }

                // Check whether block AUX_EXORGINFO exists in the rgbAuxOut
                // According to the Open Specification, version 0x01 and type 0x17 indicate that block AUX_EXORGINFO exists in rgbAuxOut payload.
                if (rgbAuxOut1.Header.Version == 0x01 && rgbAuxOut1.Header.Type == 0x17)
                {
                    int orgFlags = BitConverter.ToInt32(rgbAuxOut1.Payload, 0);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R283");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R283
                    // Check if the OrgFlags field is PUBLIC_FOLDERS_ENABLED(0x00000001).
                    Site.CaptureRequirementIfAreEqual<int>(
                        0x00000001,
                        orgFlags & 0x00000001,
                        283,
                        @"[In AUX_EXORGINFO Auxiliary Block Structure] [OrgFlags (4 bytes)] The following table [table in section 2.2.2.19] specifies the valid values [PUBLIC_FOLDERS_ENABLED 0x00000001].");

                    if (Common.IsRequirementEnabled(1692, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1692, block AUX_EXORGINFO does {0} exist in rgbAuxOut payload.", rgbAuxOut1.Header.Version == 0x01 && rgbAuxOut1.Header.Type == 0x17 ? string.Empty : "not");

                        // If code can reach here, it indicates that block AUX_EXORGINFO can be returned to client.
                        Site.CaptureRequirement(
                            1692,
                            @"[In Appendix B: Product Behavior] Implementation does return the AUX_EXORGINFO block in the rgbAuxOut buffer on the EcDoConnectEx call. (Microsoft Exchange Server 2007 and above follow this behavior.)");
                    }
                }

                // Check whether block AUX_OSVERSIONINFO exists in the rgbAuxOut
                // According to the Open Specification, version 0x01 and type 0x16 indicate that block AUX_OSVERSIONINFO exists in rgbAuxOut payload.
                if (rgbAuxOut1.Header.Version == 0x01 && rgbAuxOut1.Header.Type == 0x16)
                {
                    isAuxOsVersionInFo = true;

                    if (Common.IsRequirementEnabled(1918, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1918, the header version is {0}, the header type is {1}.", rgbAuxOut1.Header.Version, rgbAuxOut1.Header.Type);

                        // Verify MS-OXCRPC requirement: MS-OXCRPC_R1918
                        // Check if rgbAuxOut contains the AUX_OSVERSIONINFO block.
                        Site.CaptureRequirementIfIsTrue(
                            isAuxOsVersionInFo,
                            1918,
                            @"[In Appendix B: Product Behavior] Implementation does return the data of the servers operating system version and operating system service pack information in AUX_OSVERSIONINFO. (Microsoft Exchange Server 2007 follows this behavior.)");
                    }
                }

                if (Common.IsRequirementEnabled(1438, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1438, the header version is {0}, the header type is {1}.", rgbAuxOut1.Header.Version, rgbAuxOut1.Header.Type);

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1438
                    // Check if rgbAuxOut contains the AUX_OSVERSIONINFO block.
                    Site.CaptureRequirementIfIsFalse(
                        isAuxOsVersionInFo,
                        1438,
                        @"[In Appendix B: Product Behavior] Implementation doesn't support sending the AUX_OSVERSIONINFO block. <16> Section 3.1.4.1.2.1: Exchange 2010, Exchange 2013, and Exchange 2016 do not support sending the AUX_OSVERSIONINFO block.");
                }

                // Check whether block AUX_ENDPOINT_CAPABILITIES exists in the rgbAuxOut
                // According to the Open Specification, version 0x01 and type 0x48 indicate that block AUX_ENDPOINT_CAPABILITIES exists in rgbAuxOut payload.
                if (rgbAuxOut1.Header.Version == 0x01 && rgbAuxOut1.Header.Type == 0x48)
                {
                    isAuxEndpointCapabilitiesExist = true;

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1521");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1521
                    Site.CaptureRequirementIfAreEqual<int>(
                        0x00000001,
                        BitConverter.ToInt32(rgbAuxOut1.Payload, 0),
                        1521,
                        @"[In AUX_ENDPOINT_CAPABILITIES Auxiliary Block Structure] EndpointCapabilityFlag (4 bytes): The valid flag values [ENDPOINT_CAPABILITIES_SINGLE_ENDPOINT 0x00000001] are specified in the following table.");

                    if (Common.IsRequirementEnabled(1839, this.Site))
                    {
                        // Add the debug information
                        Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1839");

                        // If code can reach here, it indicates that block AUX_ENDPOINT_CAPABILITIES can be returned from server.
                        Site.CaptureRequirement(
                            1839,
                            @"[In Appendix B: Product Behavior] Implementation does support the AUX_ENDPOINT_CAPABILITIES auxiliary buffer. (Microsoft Exchange Server 2013 and above follow this behavior.)");
                    }
                }

                // Check whether block AUX_SERVER_SESSION_INFO exists in the rgbAuxOut
                // According to the Open Specification, version 0x01 and type 0x4B indicate that block AUX_SERVER_SESSION_INFO exists in rgbAuxOut payload.
                if (rgbAuxOut1.Header.Version == 0x01 && rgbAuxOut1.Header.Type == 0x4B)
                {
                    // Get the OffsetServerSessionContextInfo field.
                    short offsetServerSessionContextInfo = BitConverter.ToInt16(rgbAuxOut1.Payload, 0);

                    byte[] serverSessionContextInfo = new byte[rgbAuxOut1.Payload.Length - 2];
                    Array.Copy(rgbAuxOut1.Payload, 2, serverSessionContextInfo, 0, serverSessionContextInfo.Length);

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1535");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1535
                    // The OffsetServerSessionContextInfo should is 6 when ServerSessionContextInfo is not null or empty, because
                    // 1. AUX_HEADER is 4 bytes and OffsetServerSessionContextInfo field is 2 bytes.
                    // 2. Block AUX_SERVER_SESSION_INFO follows AUX_HEADER.
                    Site.CaptureRequirementIfAreEqual<short>(
                        6,
                        offsetServerSessionContextInfo,
                        1535,
                        @"[In AUX_SERVER_SESSION_INFO Auxiliary Block Structure] OffsetServerSessionContextInfo (2 bytes): The offset from the beginning of the AUX_HEADER structure to the ServerSessionContextInfo field.");

                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1540");

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1540
                    Site.CaptureRequirementIfAreEqual<short>(
                        6,
                        offsetServerSessionContextInfo,
                        1540,
                        @"[In AUX_SERVER_SESSION_INFO Auxiliary Block Structure] [ServerSessionContextInfo (variable)] This field is offset from the beginning of the AUX_HEADER structure by the OffsetServerSessionContextInfo field value.");

                    if (serverSessionContextInfo[serverSessionContextInfo.Length - 1] == (byte)0)
                    {
                        // The last character of ServerSessionContextInfo is null indicates that ServerSessionContextInfo is a null-terminated Unicode string, and then R1539 can be verified.
                        Site.CaptureRequirement(
                            1539,
                            @"[In AUX_SERVER_SESSION_INFO Auxiliary Block Structure] ServerSessionContextInfo (variable): A null-terminated Unicode string that contains opaque server session context information to be logged by the client.");
                    }
                }

                // The version and type got in the above code have ensured that the payload contains an AUX_HEADER structure, and followed by an auxiliary block structure according to the Open Specification.
                this.Site.CaptureRequirement(
                    4847,
                    @"[In rgbAuxOut Output Buffer] The payload contains an AUX_HEADER structure, as specified in section 2.2.2.2, followed by an auxiliary block structure as specified in the auxiliary block structure table.");
            }

            if (Common.IsRequirementEnabled(1739, this.Site))
            {
                // Add the debug information
                this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1739");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1739
                // The AUX_ENDPOINT_CAPABILITIES structure cannot be obtained. So R1739 will be verified.
                this.Site.CaptureRequirementIfIsFalse(
                    isAuxEndpointCapabilitiesExist,
                    1739,
                    @"[In Appendix B: Product Behavior]  Implementation does not support the AUX_ENDPOINT_CAPABILITIES auxiliary buffer. (<2> Section 2.2.2.2: Exchange 2007 and Exchange 2010 do not return the AUX_ENDPOINT_CAPABILITIES auxiliary buffer.)");
            }
        }

        /// <summary>
        /// Verify the auxiliary buffer returned by calling the method EcDoRpcExt2
        /// </summary>
        /// <param name="rgbAuxOutValue">Each item in this list contains the server topology information.</param>
        private void VerifyRgbAuxOutOnEcDoRpcExt2(List<AUX_SERVER_TOPOLOGY_STRUCTURE> rgbAuxOutValue)
        {
            if (Common.IsRequirementEnabled(1403, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R688");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R688
                // When pcbAuxOut is not 0(i.e., rgbAuxOut is not null) and rgbAuxOutValue returned from method ParserRgbAuxOut is not null, 
                // then the format of rgbAuxOut is consistent with Open Specification.
                // According to the implementation of ParserRgbAuxOut method, this requirement will be covered if rgbAuxOutValue count is not 0(i.e., this method succeeds).
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0,
                    rgbAuxOutValue.Count,
                    688,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [pcbAuxOut] On output, this parameter contains the size of the data to be returned in the rgbAuxOut parameter.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1073");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1073
                // According to the implementation of ParserRgbAuxOut method, if rgbAuxOutValue count is not 0(i.e., this method succeeds), this requirement will be covered.
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0,
                    rgbAuxOutValue.Count,
                    1073,
                    @"[In Server Topology Information] Each of these auxiliary blocks [AUX_CLIENT_CONTROL] MUST be preceded by a properly formatted AUX_HEADER structure, as specified in section 2.2.2.2.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1403");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1403
                // According to the implementation of ParserRgbAuxOut method, this requirement will be covered if rgbAuxOutValue count is not 0(i.e., this method succeeds).
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0,
                    rgbAuxOutValue.Count,
                    1403,
                    @"[In Appendix B: Product Behavior] Implementation does return auxiliary payload data to the client. <24> Section 3.1.4.2: Exchange 2007, Exchange 2013, and Exchange 2016 support returning data in rgbAuxOut parameter.");
            }
        }

        /// <summary>
        /// Verify whether server returns auxiliary payload data to the client when calling method EcDoRpcExt2
        /// </summary>
        /// <param name="pcbAuxOut">The size of the data to be returned in the rgbAuxOut buffer</param>
        private void VerifyIsRgbSupportedOnEcDoRpcExt2(uint pcbAuxOut)
        {
            if (Common.IsRequirementEnabled(1926, this.Site))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1926");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1926
                // According to the implementation of ParserRgbAuxOut method, this requirement will be covered if rgbAuxOutValue count is 0(i.e., this method does not succeeds).
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    pcbAuxOut,
                    1926,
                    @"[In Appendix B: Product Behavior] Implementation does not return auxiliary payload data to the client. (Microsoft Exchange Server 2010 follows this behavior.)");
            }
        }

        /// <summary>
        /// Verify whether the payload in rgbOut contains ROP response
        /// </summary>
        /// <param name="isDeserialized">A boolean value that indicates whether rgbOut can be deserialized to ROP response</param>
        /// <param name="ropCount">The count of extended buffer</param>
        private void VerifyIsRopResponse(bool isDeserialized, int ropCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R668, rgbOut can {0} be deserialized to ROP response.", isDeserialized ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R668
            // Because rgbOut is determined by pcbOut, if the rgbOut can be deserialized to ROP response, this requirement will be covered.
            Site.CaptureRequirementIfIsTrue(
                isDeserialized,
                668,
                @"[In EcDoRpcExt2 Method (opnum 11)] [pcbOut] On output, this parameter contains the size of the ROP response payload, including the size of the RPC_HEADER_EXT header in the rgbOut parameter.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1320, rgbOut can {0} be deserialized to ROP response.", isDeserialized ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1320
            // If the requirement 668 is verified, this requirement will be verified. The logic is the similar.
            Site.CaptureRequirementIfIsTrue(
                isDeserialized,
                1320,
                @"[In EcDoRpcExt2 Method (opnum 11)] [rgbOut] The size of the payload is specified in the pcbOut parameter.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R660, rgbOut can {0} be deserialized to ROP response.", isDeserialized ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R660
            // If the requirement 668 is verified, this requirement will be covered, so the logic is same.
            Site.CaptureRequirementIfIsTrue(
                isDeserialized,
                660,
                @"[In EcDoRpcExt2 Method (opnum 11)] [rgbOut] The size of the ROP response payload plus the RPC_HEADER_EXT header is returned in the pcbOut parameter.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R657, rgbOut can {0} be deserialized to ROP response.", isDeserialized ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R657
            // If the rgbOut can be deserialized to ROP response, this requirement will be covered.
            Site.CaptureRequirementIfIsTrue(
                isDeserialized,
                657,
                @"[In EcDoRpcExt2 Method (opnum 11)] rgbOut: The ROP response payload.");

            // The ropCount > 1 means that there is more than one additional buffer in rgbOut.
            if (ropCount > 1)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R967, more than one additional buffer exist, and rgbOut can {0} be deserialized to ROP response.", isDeserialized ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R967
                // If the rgbOut can be deserialized to ROP response and the count is great than 1, it means that more than one additional buffer exists in rgbOut and this requirement will be covered.
                Site.CaptureRequirementIfIsTrue(
                    isDeserialized,
                    967,
                    @"[In Extended Buffer Packing] If there is still more residual room remains in the rgbOut parameter, the server can continue to produce more response data until the rgbOut parameter does not have enough room to hold another response.");
            }
        }

        /// <summary>
        /// Verify whether the single payload exceeds 32 KB in size
        /// </summary>
        /// <param name="payloadArray">An array of single payload</param>
        private void VerifyPayloadLengthResponse(byte[][] payloadArray)
        {
            for (int i = 0; i < payloadArray.Length; i++)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1628, the size of payload {0} is {1}.", i, payloadArray[i].Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1628
                Site.CaptureRequirementIfIsTrue(
                    payloadArray[i].Length <= 32768,
                    1628,
                    @"[In Extended Buffer Handling] A single payload MUST NOT exceed 32 KB in size.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R961, the size of payload {0} is {1}.", i, payloadArray[i].Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R961
                Site.CaptureRequirementIfIsTrue(
                    payloadArray[i].Length <= 32768,
                    961,
                    @"[In Extended Buffer Packing] When processing ROP requests, the server MUST NOT produce more than 32 KB worth of response data for all ROP requests.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R964, the length of payloadArray is: {0}", payloadArray[i].Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R964
                // If payload of additional response data is less than 32 KB in size, then R964 will be verified.
                Site.CaptureRequirementIfIsTrue(
                    payloadArray[i].Length <= 32768,
                    964,
                    @"[In Extended Buffer Packing]The additional response data is also limited to 32 KB in size.");
            }
        }

        /// <summary>
        /// Verify the response payload by sending command RopQueryRowsResponse.
        /// </summary>
        /// <param name="rowCount">The rowCount field contained in the RopQueryRowsResponse payload buffer indicates the number of Row returned .</param>
        /// <param name="secRowCount">The rowCount field contained in another RopQueryRowsResponse payload buffer indicates the number of Row returned.</param>
        /// <param name="ropQueryCount"> The number of RopQueryRows </param>
        private void VerifyRopQueryRowsResponse(int rowCount, int secRowCount, int ropQueryCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "The number of Row which contained in the RopQueryRowsResponse payload buffer returned is {0}, the number of Row which contained in another RopQueryRowsResponse payload buffer is {1}.", rowCount, secRowCount);

            // If the RowCount field contained in the RopQueryRowsResponse buffer is not equal to another one.
            if (rowCount != secRowCount)
            {
                // If rowCount isn't equal to secRowCount, it indicates that there are two payloads with different RowCount, if code can arrive here, this requirement will be verified.
                Site.CaptureRequirement(
                    973,
                    @"[In Extended Buffer Packing] For the RopQueryRows ROP, the server MUST adjust the row count when adding additional response data.");
            }

            // Check for more than one additional response data.
            this.VerifyRopFastTransferSourceGetBufferResponse(ropQueryCount);
        }

        /// <summary>
        /// Verify the response payload by sending command RopReadStreamResponse.
        /// </summary>
        /// <param name="dataSize">The DataSize field contained in the RopReadStreamResponse payload buffer.</param>
        /// <param name="secDataSize">The DataSize field contained in another RopReadStreamResponse payload buffer.</param>
        /// <param name="ropReadCount"> The number of RopReadStream </param>
        private void VerifyRopReadStreamResponse(ushort dataSize, ushort secDataSize, int ropReadCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "The DataSize field contained in the RopReadStreamResponse payload buffer is {0}, the DataSize field contained in another RopReadStreamResponse payload buffer is {1}.", dataSize, secDataSize);

            // If the DataSize field contained in the RopReadStreamResponse buffer is not equal to another one.
            if (dataSize != secDataSize)
            {
                // If dataSize isn't equal to secDataSize, it indicates there are two payloads with different DataSize, if code can arrive here, this requirement will be verified.
                Site.CaptureRequirement(
                    1280,
                    @"[In Extended Buffer Packing] For the RopReadStream ROP, the server MUST adjust the number of bytes to read when adding additional response data.");
            }

            // Check for more than one additional response data.
            this.VerifyRopFastTransferSourceGetBufferResponse(ropReadCount);
        }

        /// <summary>
        /// Verify the response payload by sending command RopFastTransferSourceGetBuffer
        /// </summary>
        /// <param name="ropFastCount"> The number of RopReadStream </param>
        private void VerifyRopFastTransferSourceGetBufferResponse(int ropFastCount)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "The number of RopReadStream is {0}.", ropFastCount);

            // ropFastCount > 1 means that there are additional response date
            if (ropFastCount > 1)
            {
                // If additional response date exists and code can arrive here, this requirement will be verified.
                Site.CaptureRequirement(
                    963,
                    @"[In Extended Buffer Packing] For the server to produce additional response data, it [Extended Buffer Packing] MUST build a response as if the client sent another request with only a RopQueryRows ROP, RopReadStream ROP, or RopFastTransferSourceGetBuffer ROP.");
            }
        }

        #endregion Verify response buffer

        #region Verify referenced methods

        /// <summary>
        /// Verify the Session Context Handle returned from server by calling EcDoDisconnect.
        /// </summary>
        /// <param name="pcxh">The value of the Session Context Handle returned from server after closing the Session Context.</param>
        /// <param name="retValue">The result returned from server after calling the method EcDoDisconnect</param>
        private void VerifyEcDoDisconnect(int pcxh, uint retValue)
        {
            // The return value 0 means that the method EcDoDisconnect is called successfully.
            if (retValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R406");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R406
                Site.CaptureRequirementIfAreEqual<int>(
                    0,
                    pcxh,
                    406,
                    @"[In EcDoDisconnect Method (opnum 1)] [pcxh] On output, the server MUST set the pcxh parameter to a zero value.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1268");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1268
                // The method EcDoDisconnect belongs to the interface EMSMDB, server will return a zero value to indicate whether the Session Context is destroyed.
                Site.CaptureRequirementIfAreEqual<int>(
                    0,
                    pcxh,
                    1268,
                    @"[In Abstract Data Model] A Session Context can be destroyed through the EMSMDB interface.");
            }
        }

        /// <summary>
        /// Verify the result value returned from server by calling method EcDummyRpc.
        /// </summary>
        /// <param name="retValue">The value returned from server after calling the method EcDummyRpc.</param>
        private void VerifyEcDummyRpc(uint retValue)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R810");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R810
            Site.CaptureRequirementIfAreEqual<long>(
                0,
                retValue,
                810,
                @"[In EcDummyRpc Method (opnum 6)] Return Values: The function MUST always succeed and return 0.");
        }

        /// <summary>
        /// Verify Session Context after calling method EcDoConnectEx
        /// </summary>
        /// <param name="pcxh">The Session Context Handle returned from server after calling method EcDoConnectEx.</param>
        /// <param name="retValue">The result returned by calling method EcDoConnectEx.</param>
        private void VerifySessionContext(IntPtr pcxh, uint retValue)
        {
            // The returnValue equals to 0 indicates that method EcDoConnectEx is called successfully.
            if (retValue == 0)
            {
                #region Verify parameter pcxh in Abstract Data Model
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1212, the Session Context Handle returned from server after calling method EcDoConnectEx is {0}.", pcxh);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1212
                // The method EcDoConnect belongs to the interface EMSMDB. Server will return a handle if a Session Context is create.
                // If server returns a handle, then this handle will be not null.
                Site.CaptureRequirementIfIsNotNull(
                    pcxh,
                    1212,
                    @"[In Abstract Data Model] A Session Context can be created and destroyed through the EMSMDB interface.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1266, the Session Context Handle returned from server after calling method EcDoConnectEx is {0}.", pcxh);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1266
                // Server returns a handle responding to the session Context on the server, if the handle is not null, this requirement will be verified.
                // According to the Open Specification, the Session Context handle is returned in the response by calling the method EcDoConnectEx.
                Site.CaptureRequirementIfIsNotNull(
                    pcxh,
                    1266,
                    @"[In Global.Handle] The server provides a session context handle (the Global.Handle ADM element) to the client to make subsequent interface calls by using this same session context information.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R477, the Session Context Handle returned from server after calling method EcDoConnectEx is {0}.", pcxh);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R477
                // If server returns a handle successfully, then this handle will not be null.
                Site.CaptureRequirementIfIsNotNull(
                    pcxh,
                    477,
                    @"[In EcDoConnectEx Method (Opnum 10)] This method [method EcDoConnectEx] returns a session context handle to be used by a client in subsequent calls.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1147, the Session Context Handle returned from server after calling method EcDoConnectEx is {0}.", pcxh);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1147
                // If server returns a handle successfully, then this handle will not be null.
                Site.CaptureRequirementIfIsNotNull(
                    pcxh,
                    1147,
                    @"[In Sending the EcDoConnectEx Method] pcxh: On success, this parameter will contain the session context handle.");

                #endregion Verify parameter pcxh in Abstract Data Model
            }
        }

        /// <summary>
        /// Verify the data returned from server by calling method EcDoConnectEx.
        /// </summary>
        /// <param name="pcxh">The Session Context Handle returned from server after calling method EcDoConnectEx.</param>
        /// <param name="valueOfszDisplayName">The display name of the user associated with the szUserDN parameter.</param>
        /// <param name="rgwClientVersion">The Client Version supported by client, passed as an input parameter.</param>
        /// <param name="rgwServerVersion"> The server protocol version the client uses to determine what protocol functionality the server supports.</param>
        /// <param name="rgwBestVersion"> The minimum client protocol version the server supports.</param>
        /// <param name="picxr"> The session index value that is associated with the CXH returned from this call. </param>
        /// <param name="pcbAuxOut">The size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="pulTimeStamp">The server has to return a time stamp in which the new Session Context was created.</param>
        /// <param name="retValue">The result returned by calling this method.</param>
        private void VerifyEcDoConnectEx(
            IntPtr pcxh,
            string valueOfszDisplayName,
            ushort[] rgwClientVersion,
            ushort[] rgwServerVersion,
            ushort[] rgwBestVersion,
            ushort picxr,
            uint pcbAuxOut,
            uint pulTimeStamp,
            uint retValue)
        {
            // The returnValue equals to 0 indicates that method EcDoConnectEx is called successfully.
            if (retValue == 0)
            {
                #region Verify a single payload MUST NOT exceed 32 KB in size.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R726, the size of pcbAuxOut is {0}.", pcbAuxOut);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R726
                // The first 8 bytes are block RPC_HEADER_EXT in pcbAuxOut.
                Site.CaptureRequirementIfIsTrue(
                    (int)pcbAuxOut - 8 <= 32768,
                    726,
                    @"[In Extended Buffer Handling] A single payload MUST NOT exceed 32 KB in size.");
                #endregion Verify a single payload MUST NOT exceed 32 KB in size.

                #region Verify parameter pcxh
                this.VerifySessionContext(pcxh, retValue);
                #endregion Verify parameter pcxh

                #region Verify any user can establish an authenticated connection to the RPC server.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "The return value of calling method EcDoConnectEx is {0}.", retValue);

                // The returnValue equals to 0 indicates that method EcDoConnectEx is called successfully, and this protocol allows any user to establish an authenticated connection to the RPC server.
                Site.CaptureRequirement(
                    12,
                    @"[In Transport] This protocol allows any user to establish an authenticated connection to the RPC server by using an authentication method as specified in [MS-RPCE].");

                #endregion Verify any user can establish an authenticated connection to the RPC server.

                #region Verify parameter picxr
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R535");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R535
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    picxr.GetType(),
                    535,
                    @"[In EcDoConnectEx Method (Opnum 10)] [piCxr] The server is free to return any 16-bit value for the session index.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1172");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1172
                Site.CaptureRequirementIfAreEqual<Type>(
                    typeof(ushort),
                    picxr.GetType(),
                    1172,
                    @"[In Sending the EcDoConnectEx Method] piCxr: On success, this value is a 16-bit session index that can be used in conjunction with the value returned in the pulTimeStamp parameter to link two Session Contexts on the server.");

                #endregion Verify parameter picxr

                #region Verify szDisplayName

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R540");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R540
                string userDn = Common.GetConfigurationPropertyValue("AdminUserEssdn", this.Site);
                string expectDisplayName = userDn.Substring(
                    userDn.LastIndexOf("/") + 4,
                    userDn.Length - userDn.LastIndexOf("/") - 4).ToLower(CultureInfo.InvariantCulture);

                Site.CaptureRequirementIfAreNotEqual<string>(
                    expectDisplayName,
                    valueOfszDisplayName.ToLower(CultureInfo.InvariantCulture),
                    540,
                    @"[In EcDoConnectEx Method (Opnum 10)] szDisplayName: The display name of the user associated with the szUserDN parameter.");

                #endregion Verify szDisplayName

                #region  Verify rgwBestVersion, rgwClientVersion and rgwServerVersion when EcDoConnectEx called successfully

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCRPC_R1086, the size of the rgwSeverVersion is: {0}, the size of the rgwBestVersion is: {1}, and the size of the rgwClientVersion is : {2}.",
                    rgwServerVersion.Length,
                    rgwBestVersion.Length,
                    rgwClientVersion.Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1086
                bool isVerifyR1086 = (rgwBestVersion.Length == 3) &&
                    (rgwServerVersion.Length == 3) &&
                    (rgwClientVersion.Length == 3);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1086,
                    1086,
                    @"[In Version Number Comparison] In the EcDoConnectEx method, as specified in section 3.1.4.1, the rgwClientVersion, rgwServerVersion, and rgwBestVersion parameters are each passed as three WORD values.");

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCRPC_R1097, The size of the rgwSeverVersion is: {0}, the size of the rgwBestVersion is: {1}, and the size of the rgwClientVersion is : {2}.",
                    rgwServerVersion.Length,
                    rgwBestVersion.Length,
                    rgwClientVersion.Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1097
                bool isVerifyR1097 = isVerifyR1086;

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1097,
                    1097,
                    @"[In Version Number Comparison] In order to make version comparisons, a three-WORD value version number (as transmitted over the wire) is converted into a four-WORD value version number.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R543. The size of rgwServerVersion is {0}.", rgwServerVersion.Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R543
                // The rgwServerVersion.Length! = 0 indicates that server returns the server protocol version to the client, so R543 can be verified.
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0,
                    rgwServerVersion.Length,
                    543,
                    @"[In EcDoConnectEx Method (Opnum 10)] rgwServerVersion: The server protocol version that the client uses to determine what protocol functionality the server supports.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1180. The size of rgwServerVersion is {0}.", rgwServerVersion.Length);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1180
                Site.CaptureRequirementIfAreNotEqual<int>(
                    0,
                    rgwServerVersion.Length,
                    1180,
                    @"[In Sending the EcDoConnectEx Method] rgwServerVersion: On success, this value is the server protocol version that the client uses to determine what protocol functionality the server supports.");

                ushort[] bestVersion = new ushort[4];
                ushort[] clientVersion = new ushort[4];

                AdapterHelper.ConvertVersion(rgwBestVersion, out bestVersion);
                AdapterHelper.ConvertVersion(rgwClientVersion, out clientVersion);

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCRPC_R1303, The three Values of Best Version is: {0}, {1},{2}, and Client Version is: {3},{4},{5}.",
                    rgwBestVersion[0],
                    rgwBestVersion[1],
                    rgwBestVersion[2],
                    rgwClientVersion[0],
                    rgwClientVersion[1],
                    rgwClientVersion[2]);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1303
                // According to the Open Specification, server will set the ClientVersion passed in as the BestVersion value if version matches.
                bool isVerifyR1303 = (bestVersion[0] == clientVersion[0]) &&
                    (bestVersion[1] == clientVersion[1]) &&
                    (bestVersion[2] == clientVersion[2]) &&
                    (bestVersion[3] == clientVersion[3]);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR1303,
                    1303,
                    @"[In EcDoConnectEx Method (Opnum 10)] [rgwBestVersion] [If the client passes a client protocol version the server supports,] The server can allow the client and return the value passed by the client in the rgwClientVersion parameter.");

                // Add the debug information
                Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verify MS-OXCRPC_R547, The three Values of Best Version is: {0}, {1},{2}, and Client Version is: {3},{4},{5}.",
                    rgwBestVersion[0],
                    rgwBestVersion[1],
                    rgwBestVersion[2],
                    rgwClientVersion[0],
                    rgwClientVersion[1],
                    rgwClientVersion[2]);

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R547
                // According to the Open Specification, server will set the ClientVersion passed in as the BestVersion value if version matches.
                bool isVerifyR547 =
                    (bestVersion[0] == clientVersion[0]) &&
                    (bestVersion[1] == clientVersion[1]) &&
                    (bestVersion[2] == clientVersion[2]) &&
                    (bestVersion[3] == clientVersion[3]);

                Site.CaptureRequirementIfIsTrue(
                    isVerifyR547,
                    547,
                    @"[In EcDoConnectEx Method (Opnum 10)] [rgwBestVersion] On success, the server returns the value passed in the rgwClientVersion parameter by the client.");

                #endregion Verify rgwBestVersion, rgwClientVersion and rgwServerVersion when EcDoConnectEx called successfully

                #region Verify parameter pulTimeStamp

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R562, he server does {0} return a time stamp in which the new Session Context was created.", Convert.ToBoolean(pulTimeStamp) ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R562
                // Parameter pulTimeStamp is not null indicates that server has returned a time stamp to the client, so R562 can be verified.
                Site.CaptureRequirementIfIsNotNull(
                    pulTimeStamp,
                    562,
                    @"[In EcDoConnectEx Method (Opnum 10)] [pulTimeStamp] On output, the server has to return a time stamp in which the new Session Context was created.");

                #endregion Verify parameter pulTimeStamp
            }
            else
            {
                #region Verify pcxh and Error Code
                // Method EcDoConnectEx is called failed.
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1485");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1485
                Site.CaptureRequirementIfAreEqual<IntPtr>(
                    IntPtr.Zero,
                    pcxh,
                    1485,
                    @"[In Sending the EcDoConnectEx Method] [pcxh] On failure, this value is NULL.");

                #endregion Verify pcxh

                #region Verify parameter Version when EcDoConncextEx failed with error code ecVersionMismatch(0x80040110)

                // Check if calling method failed for version mismatch (0x80040110) reason.
                if (retValue == 0x80040110)
                {
                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-OXCRPC_R549, The three Values of Best Version is: {0}, {1},{2}, and Client Version is: {3},{4},{5}.",
                        rgwBestVersion[0],
                        rgwBestVersion[1],
                        rgwBestVersion[2],
                        rgwClientVersion[0],
                        rgwClientVersion[1],
                        rgwClientVersion[2]);

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R549
                    bool isVerifyR549 = false;

                    // If the product major version of Best Version is more than the one of Client tVersion, then the Client Version is less than the Best Version.
                    if (rgwBestVersion[0] > rgwClientVersion[0])
                    {
                        isVerifyR549 = true;
                    }
                    else if (rgwBestVersion[1] > rgwClientVersion[1])
                    {
                        // If the build major version of Best Version is more than the one of Client Version,
                        // then the Client Version is less than the Best Version.
                        isVerifyR549 = true;
                    }
                    else if (rgwBestVersion[2] > rgwClientVersion[2])
                    {
                        // If the build minor version of Best Version is more than the one of Client Version,
                        // then the Client Version is less than the Best Version.
                        isVerifyR549 = true;
                    }
                    else
                    {
                        // The Client Version is more than the Best Version.
                        isVerifyR549 = false;
                    }

                    // If server returns error code ecVersionMismatch, then server will return a Best Version that the minimum version client supports, 
                    // and the Client Version that client passes in must be more than the value Best Version.
                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR549,
                        549,
                        @"[In EcDoConnectEx Method (Opnum 10)] [rgwBestVersion] [If the client passes a client protocol version the server doesn't support,] The server can return the minimum client protocol version required to access the server and fail the call with ecVersionMismatch (0x80040110).");

                    // Add the debug information
                    Site.Log.Add(
                        LogEntryKind.Debug,
                        "Verify MS-OXCRPC_R545, The three bytes regwBestVerision is: {0}, {1}, {2}",
                        rgwBestVersion[0],
                        rgwBestVersion[1],
                        rgwBestVersion[2]);

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R545
                    // According to the Open Specification, if the client protocol version is older than that required by the server,
                    // server will fail with ecVersionMismatch(0x80040110) and return the minimum client protocol version in rgwBestVersion.
                    bool isVerifyR545 = isVerifyR549;

                    Site.CaptureRequirementIfIsTrue(
                        isVerifyR545,
                        545,
                        @"[In EcDoConnectEx Method (Opnum 10)] rgwBestVersion: The minimum client protocol version that the server supports.");
                }
                #endregion Verify parameter Version when EcDoConncextEx failed with error code ecVersionMismatch(0x80040110)
            }
        }

        /// <summary>
        /// Verify the data returned from server by calling method EcDoRpcExt2.
        /// </summary>
        /// <param name="valueOfinPcxh">The valid CXH that was created by calling EcDoConnectEx passed by Client.</param>
        /// <param name="outPcxh">The Session Context Handle returned from server.</param>
        /// <param name="pcbOut">The size of the ROP response payload.</param>
        /// <param name="pcbAuxOut">The size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="inputPcbOut">The maximum size of the ROP response payload.</param>
        /// <param name="rgbOut"> The ROP response payload.</param>
        /// <param name="inputPcbAuxOut">The maximum size of the data to be returned in the rgbAuxOut buffer.</param>
        /// <param name="rgbAuxOut">The auxiliary payload buffer.</param>
        /// <param name="retValue">The result returned from server after calling the method EcDoRpcExt2.</param>
        /// <param name="inputPulFlags">The pulFlags value on input.</param>
        /// <param name="outputPulFlags">The pulFlags value on output.</param>
        private void VerifyEcDoRpcExt2(
            int valueOfinPcxh,
            int outPcxh,
            uint pcbOut,
            uint pcbAuxOut,
            uint inputPcbOut,
            byte[] rgbOut,
            uint inputPcbAuxOut,
            byte[] rgbAuxOut,
            uint retValue,
            uint inputPulFlags,
            uint outputPulFlags)
        {
            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4852. The size of rgbOut field is {0}, the value of pcbOut field is {1}.", rgbOut.Length, inputPcbOut);
        
            // Verify MS-OXCRPC requirement: MS-OXCRPC_R4852
            this.Site.CaptureRequirementIfIsTrue(
                rgbOut.Length <= inputPcbOut,
                4852,
                @"[In EcDoRpcExt2 Method (opnum 11)] The size of the rgbOut parameter is less than or equal to pcbOut field.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R4856. The length of rgbAuxOut field is {0}, the value of pcbAuxOut field is {1}.", rgbAuxOut.Length, inputPcbAuxOut);
        
            // Verify MS-OXCRPC requirement: MS-OXCRPC_R4856
            this.Site.CaptureRequirementIfIsTrue(
                rgbAuxOut.Length <= inputPcbAuxOut,
                4856,
                @"[In EcDoRpcExt2 Method (opnum 11)] The length of the auxiliary payload buffer is less than or equal to pcbAuxOut field.");

            // The returnValue is zero indicates server returns data on success.
            if (retValue == 0)
            {
                #region Verify parameter pcxh on success

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R625");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R625
                Site.CaptureRequirementIfAreEqual<int>(
                    valueOfinPcxh,
                    outPcxh,
                    625,
                    @"[In EcDoRpcExt2 Method (opnum 11)] [pcxh] On output, the server MUST return the same session context handle on success.");

                this.VerifySimpleDataBig(pcbOut);
                this.VerifySimpleDataSmall(pcbAuxOut);
                #endregion Verify parameter pcxh on success

                if (Common.IsRequirementEnabled(1899, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1899, on output, the value of pulFlags is {0}", outputPulFlags);

                    // Verify MS-OXCRPC requirement: R1899
                    // The input of pulFlags is PulFlags.NoCompression & PulFlags.NoXorMagic, its value is 0x00000003, so the output value should be equal with this value.
                    Site.CaptureRequirementIfAreEqual<uint>(
                        inputPulFlags,
                        outputPulFlags,
                        1899,
                        @"[In Appendix B: Product Behavior] Implementation does return the same value on output as was input, not 0x00000000. (<18> Section 3.1.4.2: Exchange 2010 returns the same value on output as was input, not 0x00000000.)");
                }

                if (Common.IsRequirementEnabled(1934, this.Site))
                {
                    // Add the debug information
                    Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1934, on output, the value of pulFlags is {0}", outputPulFlags);

                    // Verify MS-OXCRPC requirement: MS-OXCRPC_R1304
                    Site.CaptureRequirementIfAreEqual<uint>(
                        0x00000000,
                        outputPulFlags,
                        1934,
                        @"[In Appendix B: Product Behavior] On output, implementation does set field pulFlags to 0x00000000. (Microsoft Exchange Server 2007, Microsoft Exchange Server 2013 and above follow this behavior.)");
                }
            }
        }

        /// <summary>
        /// Verify ACXH and result returned by calling method EcDoAsyncWaitEx.
        /// </summary>
        /// <param name="acxh">The ACXH that is associated with the Session Context passed in parameter CXH.</param>
        /// <param name="retValue">The return value by calling method EcDoAsyncWaitEx.</param>
        private void VerifyEcDoAsyncWaitEx(IntPtr acxh, uint retValue)
        {
            // The returnValue is zero indicates server returns data on success.
            if (retValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1288");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R1288
                // If the Asynchronous Context Handle pacxh returned from server is not null, then it indicates that server uses this handle 
                // in the method EcDoAsyncWaitEx which is defined in the interface AsyncEMSMEB, and this requirement will be verified.
                Site.CaptureRequirementIfAreNotEqual<IntPtr>(
                    IntPtr.Zero,
                    acxh,
                    1288,
                    @"[In Abstract Data Model] The AsyncEMSMDB uses an asynchronous Global.Handle ADM element, as defined in section 3.1.1.1.");
            }
        }

        /// <summary>
        /// Verify the data returned from server by calling method EcDoAsyncConnectEx.
        /// </summary>
        /// <param name="retValue">The return value by calling method EcDoAsyncConnectEx.</param>
        /// <param name="pacxh">An ACXH that is associated with the Session Context passed in parameter CXH.</param>
        private void VerifyEcDoAsyncConnectEx(uint retValue, IntPtr pacxh)
        {
            // The pacxh! = IntPtr.Zero indicates that establishes an asynchronous context handle successfully.
            if (pacxh != IntPtr.Zero)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R309");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R309
                Site.CaptureRequirementIfAreEqual<uint>(
                    0,
                    retValue,
                    309,
                    @"[In Protocol Details] To establish an asynchronous context handle, a call to the EcDoAsyncConnectEx method on the EMSMDB interface MUST be successful.");
            }
        }

        /// <summary>
        /// Verify EcDoAsyncWaitEx (opnum 0) pulFlagsOut
        /// </summary>
        /// <param name="pulFlagsOut">Output flags for the client.</param>
        private void VerifyEcDoAsyncWaitExpulFlagsOut(uint pulFlagsOut)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1240");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1240
            Site.CaptureRequirementIfAreEqual<uint>(
                0x00000001,
                pulFlagsOut,
                1240,
                @"[In EcDoAsyncWaitEx Method (opnum 0)] pulFlagsOut: The value of flag ""NotificationPending"" is ""0x00000001"".");
        }

        /// <summary>
        /// Verify EcRRegisterPushNotification (opnum 4)
        /// </summary>
        /// <param name="pcxh"> The Session Context Handle returned from server.</param>
        /// <param name="retValue">The result returned from server after calling the method EcRRegisterPushNotification</param>
        /// <param name="valueOfinPcxh">The valid CXH that was created by calling EcDoConnectEx passed by Client.</param>
        private void VerifyEcRRegisterPushNotification(IntPtr pcxh, uint retValue, IntPtr valueOfinPcxh)
        {
            // returnValue equals to 0 indicates that method EcRRegisterPushNotification is called successfully.
            if (retValue == 0)
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R427");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R427
                Site.CaptureRequirementIfAreEqual<IntPtr>(
                    valueOfinPcxh,
                    pcxh,
                    427,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] [pcxh] On output, the server MUST return the same session context handle on success.");

                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R451");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R451
                Site.CaptureRequirementIfAreNotEqual<IntPtr>(
                    IntPtr.Zero,
                    pcxh,
                    451,
                    @"[In EcRRegisterPushNotification Method (opnum 4)] hNotification: If the call completes successfully, the hNotification parameter contains a handle to the notification callback on the server.");
            }
        }

        #endregion Verify referenced methods

        #region Verify Compression Algorithm

        /// <summary>
        /// Verify Compression Algorithm
        /// </summary>
        /// <param name="isUseLz77">A Boolean value to verify whether the payload is compressed  by the server by using the LZ77 compression algorithm and the DIRECT2 encoding algorithm or not</param>
        /// <param name="isExOrExt2">A Boolean value, true means use method EcDoConnectEx, false means use method EcDoRpcExt2</param>
        /// <param name="isAuxOutOrOut">A Boolean value, true means field is rgbAuxOut, false means field is rgbOut</param>
        /// <param name="flag">Flag that specify how data follows RPC_HEADER_EXT header </param>
        private void VerifyCompressionAlgorithm(bool isUseLz77, bool isExOrExt2, bool isAuxOutOrOut, short flag)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R831, the payload is {0} compressed  by the server by using the LZ77 compression algorithm and the DIRECT2 encoding algorithm.", isUseLz77 ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R831
            Site.CaptureRequirementIfIsTrue(
                isUseLz77,
                831,
                @"[In Compression Algorithm] Based on the Compressed flag in the Flags field that is passed in RPC_HEADER_EXT structure, as specified in section 2.2.2.1, of the extended buffer, the payload data is compressed by the server by using the Lempel-Ziv 1977 (LZ77) compression algorithm and the DIRECT2 encoding algorithm.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R36");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R36
            Site.CaptureRequirementIfAreEqual<short>(
                0x0001,
                (short)(flag & (short)RpcHeaderExtFlags.Compressed),
                36,
                @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] [Compressed 0x0001] The data that follows the RPC_HEADER_EXT structure is compressed.");

            // Use EcDoRpcExt2 method, field is rgbOut and check if the input parameter flag contains Compressed (0x0001).
            if ((!isAuxOutOrOut) && (!isExOrExt2))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R780, the payload is {0} compressed  by the server by using the LZ77 compression algorithm and the DIRECT2 encoding algorithm.", isUseLz77 ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R780
                // According to Open Specification, rgbOut is only used by method EcDoRpcExt2 
                Site.CaptureRequirementIfIsTrue(
                    isUseLz77,
                    780,
                    @"[In rgbOut Output Buffer] If the Compressed flag is present in the Flags field of an RPC_HEADER_EXT structure, the payload data associated with that RPC_HEADER_EXT structure MUST be compressed by the server.");
            }
        }

        /// <summary>
        /// Verify DIRECT2 Encoding Algorithm
        /// </summary>
        /// <param name="isDirect2Encoding">A Boolean value to verify whether the decoder identify the presence of metadata in the compressed and encoded data stream or not</param>
        private void VerifyDIRECT2EncodingAlgorithm(bool isDirect2Encoding)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R881, the decoder does {0} identify the presence of metadata in the compressed and encoded data stream.", isDirect2Encoding ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R881
            // The DIRECT2 Encoding Algorithm in the implementation of method LZ77Decompress is consistent with Open Specification, this requirement will be covered if LZ77Decompress succeeds.
            Site.CaptureRequirementIfIsTrue(
                isDirect2Encoding,
                881,
                @"[In DIRECT2 Encoding Algorithm] The decoder MUST be able to identify the presence of metadata in the compressed and encoded data stream.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R887, the decoder does {0} identify the presence of metadata in the compressed and encoded data stream.", isDirect2Encoding ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R887
            // The DIRECT2 Encoding Algorithm in the implementation of method LZ77Decompress is consistent with Open Specification, bit "0" included in DIRECT2 Encoding Algorithm, and this requirement will be covered if LZ77Decompress succeeds.
            Site.CaptureRequirementIfIsTrue(
                isDirect2Encoding,
                887,
                @"[In Bitmask]If a ""0"" bit is encountered, the next byte in the input stream is the next byte in the output stream.");

            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R888, the decoder does {0} identify the presence of metadata in the compressed and encoded data stream.", isDirect2Encoding ? string.Empty : "not");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R888
            // The DIRECT2 Encoding Algorithm in the implementation of method LZ77Decompress is consistent with Open Specification, bit "1" included in DIRECT2 Encoding Algorithm, and this requirement will be covered if LZ77Decompress succeeds.
            Site.CaptureRequirementIfIsTrue(
                isDirect2Encoding,
                888,
                @"[In Bitmask]If a ""1"" bit is encountered, the next byte or series of bytes is metadata that MUST be interpreted further.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1873");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1873
            // The DIRECT2 Encoding Algorithm in the implementation of method LZ77Decompress is consistent with Open Specification, bit "1" included in DIRECT2 Encoding Algorithm, and this requirement will be covered if LZ77Decompress succeeds.
            this.Site.CaptureRequirementIfIsTrue(
                isDirect2Encoding,
                1873,
                @"[In Bitmask] The bitmask also contains a ""1"" in the bit following the last encoded element to indicate the end of the compressed data.");

            // Add the debug information
            this.Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R1877");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R1877
            // The DIRECT2 Encoding Algorithm in the implementation of method LZ77Decompress is consistent with Open Specification, bit "1" included in DIRECT2 Encoding Algorithm, and this requirement will be covered if LZ77Decompress succeeds.
            this.Site.CaptureRequirementIfIsTrue(
                isDirect2Encoding,
                1877,
                @"[In Bitmask] The final end bit is always necessary, even if an additional bitmask has to be allocated.");
        }

        /// <summary>
        /// Verify obfuscation algorithm
        /// </summary>
        /// <param name="isObfuscation">A Boolean value that indicates whether the obfuscation succeeds</param>
        /// <param name="isExOrExt2">A Boolean value, true means using method EcDoConnectEx, false means using method EcDoRpcExt2</param>
        /// <param name="isAuxOutOrOut">A Boolean value, true means field is rgbAuxOut, false means field is rgbOut</param>
        /// <param name="flag">Flag that specify how data that follows RPC_HEADER_EXT header</param>
        private void VerifyObfuscationAlgorithm(bool isObfuscation, bool isExOrExt2, bool isAuxOutOrOut, short flag)
        {
            // Add the debug information
            Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R39");

            // Verify MS-OXCRPC requirement: MS-OXCRPC_R39
            Site.CaptureRequirementIfAreEqual<short>(
                0x0002,
                (short)(flag & (short)RpcHeaderExtFlags.XorMagic),
                39,
                @"[In RPC_HEADER_EXT Structure] [Flags (2 bytes)] [XorMagic 0x0002] The data following the RPC_HEADER_EXT structure has been obfuscated.");

            // Use EcDoRpcExt2 method, field is rgbOut and check if the input parameter flag contains XorMagic(0x0002)
            if ((!isAuxOutOrOut) && (!isExOrExt2))
            {
                // Add the debug information
                Site.Log.Add(LogEntryKind.Debug, "Verify MS-OXCRPC_R783, the obfuscation does {0} succeed.", isObfuscation ? string.Empty : "not");

                // Verify MS-OXCRPC requirement: MS-OXCRPC_R783
                // According to Open Specification, rgbOut is only used by method EcDoRpcExt2 
                Site.CaptureRequirementIfIsTrue(
                    isObfuscation,
                    783,
                    @"[In rgbOut Output Buffer] If the XorMagic flag is present in the Flags field of an RPC_HEADER_EXT structure, the payload data associated with that RPC_HEADER_EXT structure MUST be obfuscated by the server.");
            }
        }
        #endregion Verify Compression Algorithm
    }
}