//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    /// <summary>
    /// Organizes the const values
    /// </summary>
    public static class ConstValues
    {
        /// <summary>
        /// Size of Guid in byte
        /// </summary>
        public const int GuidByteSize = 16;

        /// <summary>
        /// Size of GlobalCounter field in byte
        /// </summary>
        public const int GlobalCounterByteSize = 6;

        /// <summary>
        /// MaximumByteCount field is present when ByteCount is equal to 0xBABE(47806)
        /// </summary>
        public const ushort MaximumByteCountIndicator = 0xbabe;

        /// <summary>
        /// The size of AUX_HEADER.
        /// </summary>
        public const int AuxHeaderByteSize = 4;

        /// <summary>
        /// The maximum value of RowCount field in RopQueryRows request
        /// </summary>
        public const ushort QueryRowsRequestRowCountMax = 0xffff;

        /// <summary>
        /// The size of sample tags.
        /// </summary>
        public const int SampleTagsByteSize = 6;

        /// <summary>
        /// Indicate the MessageId field in Message ROPs with value 0x00000000000000000.
        /// </summary>
        public const ulong MessageIdZeroValue = 0x00000000000000000;

        /// <summary>
        /// The payload length in ROP input buffer for null ROP request
        /// </summary>
        public const int RequestBufferPayloadLength = 2;

        /// <summary>
        /// The RopSize value in ROP input buffer for null ROP request
        /// </summary>
        public const ushort RequestBufferRopSize = 2;

        /// <summary>
        /// Size of property which type is PytpTime in byte
        /// </summary>
        public const int PtypTimeByteSize = 8;

        /// <summary>
        /// Size of Version field in RPC_HEADER_EXT
        /// </summary>
        public const int RpcHeaderExtVersionByteSize = 2;

        /// <summary>
        /// Size of Flags field in RPC_HEADER_EXT
        /// </summary>
        public const int RpcHeaderExtFlagsByteSize = 2;

        /// <summary>
        /// Size of Size field in RPC_HEADER_EXT
        /// </summary>
        public const int RpcHeaderExtSizeByteSize = 2;

        /// <summary>
        /// A Default value for Locale. 1033 indicates "en-us".
        /// </summary>
        public const uint DefaultLocale = 0x00000409;

        /// <summary>
        /// The base number of product major version in three WORD version.
        /// </summary>
        public const long OffsetProductMajorVersion = 0x100000000;
        
        /// <summary>
        /// The base number of build major number in three WORD version.
        /// </summary>
        public const int OffsetBuildMajorNumber = 0x10000;
        
        /// <summary>
        /// The size of RopSize in ROP input and output buffer structure.
        /// </summary>
        public const int RopSizeInRopInputOutputBufferSize = 2;
        
        /// <summary>
        /// 96 is an arbitrary initial size for Buffer
        /// </summary>
        public const int ArbitraryInitialSizeForBuffer = 96;
        
        /// <summary>
        /// The size of Size field in AUX_HEADER.
        /// </summary>
        public const int AuxHeaderSizeByteSize = 2;
        
        /// <summary>
        /// The size of Version field in AUX_HEADER.
        /// </summary>
        public const int AuxHeaderVersionByteSize = 1;
        
        /// <summary>
        /// The size of Type field in AUX_HEADER.
        /// </summary>
        public const int AuxHeaderTypeByteSize = 1;
        
        /// <summary>
        /// The size of AUX_HEADER.
        /// </summary>
        public const int AuxHeaderSize = AuxHeaderSizeByteSize + AuxHeaderVersionByteSize + AuxHeaderTypeByteSize;
        
        /// <summary>
        /// The size of normalized version.
        /// </summary>
        public const int NormalizedVersionSize = 4;

        /// <summary>
        /// The mask to get the high-bit of WORD.
        /// </summary>
        public const ushort HighBitMask = 0x8000;

        /// <summary>
        /// The high-order BYTE of the first WORD in Version.
        /// </summary>
        public const ushort HighByteMask = 0xFF00;
        
        /// <summary>
        /// The low-order BYTE of the first WORD in Version.
        /// </summary>
        public const ushort LowByteMask = 0x00FF;
        
        /// <summary>
        /// The step distance of one byte.
        /// </summary>
        public const int StepDistanceOfOneByte = 8;
        
        /// <summary>
        /// TableFlags of ConversationMembers.
        /// </summary>
        public const int ConversationMemberTableFlag = 0x80;

        /// <summary>
        /// The GID length.
        /// </summary>
        public const int GidLength = 22;
        
        /// <summary>
        /// The size of Pad field.
        /// </summary>
        public const int PadFieldByteSize = 2;

        /// <summary>
        /// Value of Version field in RPC_HEADER_EXT
        /// </summary>
        public const ushort RpcHeaderExtVersionValue = 0x0000;

        /// <summary>
        /// Code page id used in RopLogon
        /// </summary>
        public const ushort LogonCodePageId = 0x0FFF;

        /// <summary>
        /// The endmark of a string.
        /// </summary>
        public const char StringEndMark = '\0';

        /// <summary>
        /// The default value for server output object handle.
        /// </summary>
        public const uint DefaultOutputHandle = 0xFFFFFFFF;

        /// <summary>
        /// Indicate the OutputHandleIndex field.
        /// </summary>
        public const uint OutputHandleIndexForOneRop = 1;

        /// <summary>
        /// Indicate a value that is not zero.
        /// </summary>
        public const bool NonZero = true;

        /// <summary>
        /// An unsigned integer indicates an invalid value of pcbAuxOut that that should be larger than 4104 (0x00001008), as specified by EcDoConnectEx method and EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooBigpcbAuxOut = 4105;

        /// <summary>
        /// An unsigned integer indicates a valid value of pcbAuxOut, as specified by EcDoConnectEx method and EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint ValidpcbAuxOut = 4104;

        /// <summary>
        /// An unsigned integer indicates an invalid value of cbAuxIn that should be larger than 4104 (0x00001008), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooBigcbAuxIn = 4105;

        /// <summary>
        /// An unsigned integer indicates an invalid value of cbAuxIn that should be larger than 0 and less than 8 (0x00000008), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooSmallcbAuxIn = 7;

        /// <summary>
        /// An unsigned integer indicates an invalid value of cbIn that should be larger than 32775 (0x00008007) and less than 262144 (0x00040000), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint BigcbIn = 32776;

        /// <summary>
        /// An unsigned integer indicates an invalid value of cbIn that should be larger than 262144 (0x00040000) for Exchange 2010 and 2013, as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooBigcbIn = 262145;

        /// <summary>
        /// An unsigned integer indicates an invalid value of cbIn that should be less than 8 (0x00000008), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooSmallcbIn = 7;

        /// <summary>
        /// An unsigned integer indicates an invalid value of pcbOut that should be less than 8 (0x00000008), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooSmallpcbOut = 7;

        /// <summary>
        /// An unsigned integer indicates an invalid value of pcbOut that should be larger than 8 (0x00000008) and less than 32775 (0x00008007), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint SmallpcbOut = 32768;

        /// <summary>
        /// An unsigned integer indicates a valid value of pcbOut, as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint ValidpcbOut = 262144;

        /// <summary>
        /// An integer indicates an invalid value of pcbOut that should be larger than 262144 (0x00040000), as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const uint TooBigpcbOut = 262145;

        /// <summary>
        /// InvalidPcxh specify an invalid session context handle, as specified by EcDoRpcExt2 method in [MS-OXCRPC].
        /// </summary>
        public const int InvalidPcxh = 1;

        /// <summary>
        /// An unsigned integer indicates a client-derived 32-bit hash value of the User's distinguished name, as specified by EcDoConnectEx method in [MS-OXCRPC].
        /// </summary>
        public const uint ConnectionMod = 3409270;

        /// <summary>
        /// An unsigned integer indicates the code page in which text data is sent if Unicode format is not requested by the client on subsequent calls using this Session Context, as specified by EcDoConnectEx method in [MS-OXCRPC].
        /// </summary>
        public const uint CodePageId = 1252;

        /// <summary>
        /// An unsigned short indicates an arbitrary invalid value for picxr, as specified by EcDoConnectEx method in [MS-OXCRPC].
        /// </summary>
        public const ushort Invalidpicxr = 171;

        /// <summary>
        /// An unsigned long indicates a value of RowCount field in RopQueryRows request, as specified by RopQueryRows ROP in [MS-OXCROPS].
        /// </summary>
        public const ulong MaximumRowCount = 65535;

        /// <summary>
        /// An unsigned long indicates a small value of MaximumByteCount field in RopReadStream request, as specified by RopReadStream ROP in [MS-OXCROPS].
        /// </summary>
        public const ulong RequestedByteCount = 32;

        /// <summary>
        /// An integer indicates a value of MaximumBufferSize field in RopFastTransferSourceGetBuffer request, as specified by RopFastTransferSourceGetBuffer ROP in [MS-OXCROPS].
        /// </summary>
        public const ulong MaximumBufferSize = 65535;
    }
}