namespace Microsoft.Protocols.TestSuites.MS_OXCRPC
{
    using System;
    using System.Runtime.InteropServices;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The class provides methods to assist MS-OXCRPCAdapter.
    /// </summary>
    public static class AdapterHelper
    {
        /// <summary>
        /// An integer indicates the length of RPC_HEADER_EXT.
        /// </summary>
        public static readonly int RPCHeaderExtlength = Marshal.SizeOf(typeof(RPC_HEADER_EXT));

        /// <summary>
        /// Transfer ITestSite into adapter, make adapter can use ITestSite's function.
        /// </summary>
        private static ITestSite site;

        /// <summary>
        /// Initialize the AdapterHelper class.
        /// </summary>
        /// <param name="testSite">The instance of ITestSite.</param>
        public static void Initialize(ITestSite testSite)
        {
            site = testSite;
        }

        /// <summary>
        /// A method used to compose ROP command to execute ROP call.
        /// </summary>
        /// <param name="ropType">A parameter represents the ROP Type refers to a specific ROP call.</param>
        /// <param name="inputServerObjectHandle">A parameter represents an input Server object handle.</param>
        /// <param name="propertyValue">A parameter represents a TaggedPropertyValue structure.</param>
        /// <returns>A return value represents a byte array of serialized ROP command.</returns>
        public static byte[] ComposeRgbIn(ROPCommandType ropType, uint inputServerObjectHandle, TaggedPropertyValue propertyValue)
        {
            ComposeROPCommand composeROPCommand = new ComposeROPCommand();

            switch (ropType)
            {
                case ROPCommandType.RopSynchronizationImportDeletes:
                    return composeROPCommand.ComposeRopSynchronizationImportDeletes(inputServerObjectHandle, propertyValue);

                default:
                    throw new InvalidCastException("Invalid ROP command type");
            }
        }

        /// <summary>
        /// A method used to compose ROP command to execute ROP call.
        /// </summary>
        /// <param name="ropType">A parameter represents the ROP Type refers to a specific ROP call.</param>
        /// <param name="inputServerObjectHandle">A parameter represents an input Server object handle.</param>
        /// <param name="auxInfo">The auxInfo point to the object executed by this ROP call. 
        /// It could be a folder handle, a message handle or an attachment handle or the count to be operated.</param>
        /// <returns>A return value represents a byte array of serialized ROP command.</returns>
        public static byte[] ComposeRgbIn(ROPCommandType ropType, uint inputServerObjectHandle, ulong auxInfo)
        {
            ComposeROPCommand composeROPCommand = new ComposeROPCommand();
            
            // UserDN used for logon request.
            string userDN = null;
            
            switch (ropType)
            {
                case ROPCommandType.RopLogon:
                    userDN = Common.GetConfigurationPropertyValue("AdminUserEssdn", site) + "\0";
                    return composeROPCommand.ComposeLogonRequest(userDN, (uint)auxInfo);
                case ROPCommandType.RopLogonNormalUser:
                    userDN = Common.GetConfigurationPropertyValue("NormalUserEssdn", site) + "\0";
                    return composeROPCommand.ComposeLogonRequest(userDN, (uint)auxInfo);
                case ROPCommandType.RopCreateMessage:
                    return composeROPCommand.ComposeCreateMessageRequest(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopOpenStream:
                    return composeROPCommand.ComposeOpenStreamRequest(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopReadStream:
                    return composeROPCommand.ComposeReadStreamRequest(inputServerObjectHandle, (uint)auxInfo);
                case ROPCommandType.RopCommitStream:
                    return composeROPCommand.ComposeCommitStreamRequest(inputServerObjectHandle);
                case ROPCommandType.RopWriteStream:
                    return composeROPCommand.ComposeWriteStreamRequest(inputServerObjectHandle);
                case ROPCommandType.RopOpenFolder:
                    return composeROPCommand.ComposeOpenFolderRequest(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopGetHierarchyTable:
                    return composeROPCommand.ComposeGetHierarchyTableRequest(inputServerObjectHandle);
                case ROPCommandType.RopSaveChangesMessage:
                    return composeROPCommand.ComposeSaveChangesMessageRequest(inputServerObjectHandle);
                case ROPCommandType.RopQueryRows:
                    return composeROPCommand.ComposeQueryRowsRequest(inputServerObjectHandle, (ushort)auxInfo);
                case ROPCommandType.RopFastTransferSourceCopyMessages:
                    return composeROPCommand.ComposeFastTransferSourceCopyMessagesRequest(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopFastTransferSourceGetBuffer:
                    return composeROPCommand.ComposeRopFastTransferSourceGetBufferRequest(inputServerObjectHandle, (ushort)auxInfo);
                case ROPCommandType.RopSetColumns:
                    return composeROPCommand.ComposeSetColumnsRequest(inputServerObjectHandle);
                case ROPCommandType.RopRegisterNotification:
                    return composeROPCommand.ComposeRegisterNotifyEvents(inputServerObjectHandle, auxInfo);
                case ROPCommandType.WithoutRops:
                    return composeROPCommand.BuildRequestBuffer();
                case ROPCommandType.MultipleRops:
                    return composeROPCommand.ComposeMultipleRops(inputServerObjectHandle);
                case ROPCommandType.RopGetContentsTable:
                    return composeROPCommand.ComposeGetContentsTableRequest(inputServerObjectHandle, (byte)auxInfo);
                case ROPCommandType.RopLongTermIdFromId:
                    return composeROPCommand.ComposeRopLongTermIdFromId(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopSynchronizationOpenCollector:
                    return composeROPCommand.ComposeRopSynchronizationOpenCollector(inputServerObjectHandle);
                case ROPCommandType.RopSynchronizationImportMessageChange:
                    return composeROPCommand.ComposeRopSynchronizationImportMessageChange(inputServerObjectHandle, (byte)auxInfo);
                case ROPCommandType.RopCreateFolder:
                    return composeROPCommand.ComposeCreateFolderRequest(inputServerObjectHandle, site);
                case ROPCommandType.RopDeleteFolder:
                    return composeROPCommand.ComposeDeleteFolderRequest(inputServerObjectHandle, auxInfo);
                case ROPCommandType.RopRelease:
                    return composeROPCommand.ComposeRopReleaseRequest(inputServerObjectHandle);
                case ROPCommandType.RopHardDeleteMessagesAndSubfolders:
                    return composeROPCommand.ComposeRopHardDeleteMessagesAndSubfolders(inputServerObjectHandle);
                default:
                    throw new InvalidCastException("Invalid ROP command type");
            }
        }

        /// <summary>
        /// Composes rgbAuxIn buffer
        /// </summary>
        /// <param name="rgbAuxIn">Enum value of rgbAuxIn</param>
        /// <param name="rgbAuxInPayload">The payload of rgbAuxIn buffer.</param>
        /// <returns>Returns a buffer composed according to the parameters</returns>
        public static byte[] ComposeRgbAuxIn(RgbAuxInEnum rgbAuxIn, byte[] rgbAuxInPayload)
        {
            byte version;
            byte type;
            const string ErrorMessage = "This input structure name is not a valid AUX structure type";

            // Version set to 0x01 means the Version field of AUX_HEADER is AUX_VERSION_1
            // Version set to 0x02 means the Version field of AUX_HEADER is AUX_VERSION_2
            switch (rgbAuxIn)
            {
                case RgbAuxInEnum.AUX_PERF_SESSIONINFO:
                     version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_SESSIONINFO;
                    break;
                case RgbAuxInEnum.AUX_PERF_SESSIONINFO_V2:
                    version = (byte)AuxVersions.AUX_VERSION_2;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_SESSIONINFO;
                    break;
                case RgbAuxInEnum.AUX_PERF_CLIENTINFO:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_CLIENTINFO;
                    break;
                case RgbAuxInEnum.AUX_PERF_PROCESSINFO:
                    version = (byte)AuxVersions.AUX_VERSION_2;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_PROCESSINFO;
                    break;
                case RgbAuxInEnum.AUX_PERF_DEFMDB_SUCCESS:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_DEFMDB_SUCCESS;
                    break;
                case RgbAuxInEnum.AUX_PERF_DEFGC_SUCCESS:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_DEFGC_SUCCESS;
                    break;
                case RgbAuxInEnum.AUX_PERF_MDB_SUCCESS_V2:
                    version = (byte)AuxVersions.AUX_VERSION_2;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_MDB_SUCCESS;
                    break;
                case RgbAuxInEnum.AUX_PERF_GC_SUCCESS:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_GC_SUCCESS;
                    break;
                case RgbAuxInEnum.AUX_PERF_GC_SUCCESS_V2:
                    version = (byte)AuxVersions.AUX_VERSION_2;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_GC_SUCCESS;
                    break;
                case RgbAuxInEnum.AUX_PERF_FAILURE:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_FAILURE;
                    break;
                case RgbAuxInEnum.AUX_PERF_FAILURE_V2:
                     version = (byte)AuxVersions.AUX_VERSION_2;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_BG_FAILURE;
                    break;
                case RgbAuxInEnum.AUX_PERF_ACCOUNTINFO:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_TYPE_PERF_ACCOUNTINFO;
                    break;
                case RgbAuxInEnum.AUX_CLIENT_CONNECTION_INFO:
                    version = (byte)AuxVersions.AUX_VERSION_1;
                    type = (byte)AuxTypes.AUX_CLIENT_CONNECTION_INFO;
                    break;
                default:
                    throw new NotImplementedException(ErrorMessage);
            }

            RPC_HEADER_EXT rpcHeaderExt = new RPC_HEADER_EXT
            {
                Version = 0x0000, // This value MUST be set to 0x0000.
                Flags = (ushort)RpcHeaderExtFlags.Last, // No Compression and no obfuscation
                Size = (ushort)(rgbAuxInPayload.Length + ConstValues.AuxHeaderByteSize) // The total length of the payload data that follows the RPC_HEADER_EXT structure. 
            };

            // The length of the payload data after it has been uncompressed.
            rpcHeaderExt.SizeActual = rpcHeaderExt.Size;

            // The length of AuxHeader is the length of AUX_HEADER structure plus the length of any additional payload data.
            AUX_HEADER auxHeader = new AUX_HEADER
            {
                Size = (ushort)(rgbAuxInPayload.Length + ConstValues.AuxHeaderByteSize),
                Version = version,
                Type = type
            };
            int returnSize = rgbAuxInPayload.Length + Marshal.SizeOf(rpcHeaderExt) + Marshal.SizeOf(auxHeader);
            byte[] returnByte = new byte[returnSize];

            // RgbAuxIn serialization 
            // Serializes the RPC_HEADER_EXT
            int index = 0;
            Array.Copy(BitConverter.GetBytes(rpcHeaderExt.Version), 0, returnByte, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(rpcHeaderExt.Flags), 0, returnByte, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(rpcHeaderExt.Size), 0, returnByte, index, sizeof(short));
            index += sizeof(short);
            Array.Copy(BitConverter.GetBytes(rpcHeaderExt.SizeActual), 0, returnByte, index, sizeof(short));
            index += sizeof(short);

            // Serializes the AUX_HEADER
            Array.Copy(BitConverter.GetBytes(auxHeader.Size), 0, returnByte, index, sizeof(ushort));
            index += sizeof(ushort);
            returnByte[index] = auxHeader.Version;
            index++;
            returnByte[index] = auxHeader.Type;
            index++;

            Array.Copy(rgbAuxInPayload, 0, returnByte, index, rgbAuxInPayload.Length);

            return returnByte;
        }

        #region Buffer composer
        /// <summary>
        /// Composes buffer for AUX_PERF_SESSIONINFO, refer to the definition of AUX_PERF_SESSIONINFO in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_SESSIONINFO</param>
        /// <returns>The byte array contains the AUX_PERF_SESSIONINFO</returns>
        public static byte[] Compose_AUX_PERF_SESSIONINFO(int reserved)
        {
            AUX_PERF_SESSIONINFO aux_perf_sessioninfo = new AUX_PERF_SESSIONINFO
            {
                SessionID = 0,
                Reserved = (short)reserved,
                SessionGuid = new byte[ConstValues.GuidByteSize]
            };
            return aux_perf_sessioninfo.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_SESSIONINFO_V2, refer to the definition of AUX_PERF_SESSIONINFO_V2 in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_SESSIONINFO_V2</param>
        /// <returns>The byte array contains the AUX_PERF_SESSIONINFO_V2</returns>
        public static byte[] Compose_AUX_PERF_SESSIONINFO_V2(int reserved)
        {
            AUX_PERF_SESSIONINFO_v2 aux_perf_sessioninfo_v2 = new AUX_PERF_SESSIONINFO_v2
            {
                SessionID = 0,
                Reserved = (short)reserved,
                SessionGuid = new byte[ConstValues.GuidByteSize],
                ConnectionID = 0
            };
            return aux_perf_sessioninfo_v2.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_CLIENTINFO, refer to the definition of AUX_PERF_CLIENTINFO in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_CLIENTINFO</param>
        /// <param name="clientMode">The value of ClientMode that determines the mode in which the client is running</param>
        /// <returns>The byte array contains the AUX_PERF_CLIENTINFO</returns>
        public static byte[] Compose_AUX_PERF_CLIENTINFO(int reserved, short clientMode)
        {
            AUX_PERF_CLIENTINFO aux_perf_clientinfo = new AUX_PERF_CLIENTINFO
            {
                ClientMode = clientMode,
                Reserved = (short)reserved
            };
            return aux_perf_clientinfo.Serialize();
        }

        /// <summary>
        ///  Composes buffer for AUX_PERF_PROCESSINFO, refer to the definition of AUX_PERF_PROCESSINFO in this Open Specification
        /// </summary>
        /// <param name="reserved1">Value for Reserved_1 field in AUX_PERF_PROCESSINFO</param>
        /// <param name="reserved2">Value for Reserved_2 field in AUX_PERF_PROCESSINFO</param>
        /// <returns>The byte array contains the AUX_PERF_PROCESSINFO</returns>
        public static byte[] Compose_AUX_PERF_PROCESSINFO(int reserved1, int reserved2)
        {
            AUX_PERF_PROCESSINFO aux_perf_processinfo = new AUX_PERF_PROCESSINFO
            {
                ProcessID = 0,
                ProcessNameOffset = 0,
                ProcessGuid = new byte[ConstValues.GuidByteSize],
                Reserved1 = (short)reserved1,
                Reserved2 = (short)reserved2
            };
            return aux_perf_processinfo.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_DEFMDB_SUCCESS, refer to the definition of AUX_PERF_DEFMDB_SUCCESS in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_DEFMDB_SUCCESS</param>
        /// <returns>The byte array contains the AUX_PERF_DEFMDB_SUCCESS</returns>
        public static byte[] Compose_AUX_PERF_DEFMDB_SUCCESS(int reserved)
        {
            AUX_PERF_DEFMDB_SUCCESS aux_perf_defmdbsuccess = new AUX_PERF_DEFMDB_SUCCESS
            {
                Reserved = (short)reserved
            };
            return aux_perf_defmdbsuccess.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_DEFGC_SUCCESS, refer to the definition of AUX_PERF_DEFGC_SUCCESS in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_DEFGC_SUCCESS</param>
        /// <returns>The byte array contains the AUX_PERF_DEFGC_SUCCESS</returns>
        public static byte[] Compose_AUX_PERF_DEFGC_SUCCESS(int reserved)
        {
            AUX_PERF_DEFGC_SUCCESS aux_perf_defgcsuccess = new AUX_PERF_DEFGC_SUCCESS
            {
                Reserved = reserved
            };
            return aux_perf_defgcsuccess.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_MDB_SUCCESS_V2, refer to the definition of AUX_PERF_MDB_SUCCESS_V2 in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_MDB_SUCCESS_V2</param>
        /// <returns>The byte array contains the AUX_PERF_MDB_SUCCESS_V2</returns>
        public static byte[] Compose_AUX_PERF_MDB_SUCCESS_V2(int reserved)
        {
            AUX_PERF_MDB_SUCCESS_V2 aux_perf_mdbsuccessv2 = new AUX_PERF_MDB_SUCCESS_V2
            {
                Reserved = (short)reserved
            };
            return aux_perf_mdbsuccessv2.Serialize();
        }

        /// <summary>
        ///  Composes buffer for AUX_PERF_GC_SUCCESS, refer to the definition of AUX_PERF_GC_SUCCESS in this Open Specification
        /// </summary>
        /// <param name="reserved1">Value for Reserved_1 field in AUX_PERF_GC_SUCCESS</param>
        /// <param name="reserved2">Value for Reserved_2 field in AUX_PERF_GC_SUCCESS</param>
        /// <returns>The byte array contains the AUX_PERF_GC_SUCCESS</returns>
        public static byte[] Compose_AUX_PERF_GC_SUCCESS(int reserved1, int reserved2)
        {
            AUX_PERF_GC_SUCCESS aux_perf_gc_success = new AUX_PERF_GC_SUCCESS
            {
                Reserved1 = (short)reserved1,
                Reserved2 = reserved2
            };
            return aux_perf_gc_success.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_GC_SUCCESS_V2, refer to the definition of AUX_PERF_GC_SUCCESS_V2 in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_GC_SUCCESS_V2</param>
        /// <returns>The byte array contains the AUX_PERF_GC_SUCCESS_V2</returns>
        public static byte[] Compose_AUX_PERF_GC_SUCCESS_V2(int reserved)
        {
            AUX_PERF_GC_SUCCESS_V2 aux_perf_gc_success_v2 = new AUX_PERF_GC_SUCCESS_V2
            {
                Reserved = reserved
            };
            return aux_perf_gc_success_v2.Serialize();
        }

        /// <summary>
        /// Composes buffer for AUX_PERF_FAILURE, refer to the definition of AUX_PERF_FAILURE in this Open Specification
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_FAILURE</param>
        /// <returns>The byte array contains the AUX_PERF_FAILURE</returns>
        public static byte[] Compose_AUX_PERF_FAILURE(int reserved)
        {
            AUX_PERF_FAILURE aux_perf_failure = new AUX_PERF_FAILURE
            {
                Reserved = reserved
            };
            return aux_perf_failure.Serialize();
        }

        /// <summary>
        ///  Composes buffer for AUX_PERF_FAILURE_V2, refer to definition of AUX_PERF_FAILURE_V2 in this Open Specification
        /// </summary>
        /// <param name="reserved1">Value for Reserved_1 field in AUX_PERF_FAILURE_V2</param>
        /// <param name="reserved2">Value for Reserved_2 field in AUX_PERF_FAILURE_V2</param>
        /// <returns>The byte array contains the AUX_PERF_FAILURE_V2</returns>
        public static byte[] Compose_AUX_PERF_FAILURE_V2(int reserved1, int reserved2)
        {
            AUX_PERF_FAILURE_V2 aux_perf_failure_v2 = new AUX_PERF_FAILURE_V2
            {
                Reserved1 = (short)reserved1,
                Reserved2 = reserved2
            };
            return aux_perf_failure_v2.Serialize();
        }

        /// <summary>
        /// Compose buffer for AUX_PERF_ACCOUNTINFO, refer to the definition of AUX_PERF_ACCOUNTINFO 
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_PERF_ACCOUNTINFO</param>
        /// <returns>The byte array contains the AUX_PERF_ACCOUNTINFO</returns>
        public static byte[] Compose_AUX_PERF_ACCOUNTINFO(int reserved)
        {
            AUX_PERF_ACCOUNTINFO aux_perf_accountinfo = new AUX_PERF_ACCOUNTINFO
            {
                Reserved = (short)reserved
            };
            return aux_perf_accountinfo.Serialize();
        }

        /// <summary>
        /// Compose buffer for AUX_CLIENT_CONNECTION_INFO
        /// </summary>
        /// <param name="reserved">Value for Reserved field in AUX_CLIENT_CONNECTION_INFO</param>
        /// <param name="connectionFlags">The value of ConnectionFlags</param>
        /// <returns>The byte array contains the AUX_CLIENT_CONNECTION_INFO</returns>
        public static byte[] Compose_AUX_CLIENT_CONNECTION_INFO(int reserved, int connectionFlags)
        {
            AUX_CLIENT_CONNECTION_INFO aux_client_connectionInfo = new AUX_CLIENT_CONNECTION_INFO
            {
                ConnectionFlags = connectionFlags,
                Reserved = (short)reserved
            };
            return aux_client_connectionInfo.Serialize();
        }
        #endregion

        /// <summary>
        /// Convert the three DWORD version into a normalized version.
        /// </summary>
        /// <param name="version">The three DWORD version.</param>
        /// <param name="normalizeVersion">The returned normalized version.</param>
        public static void ConvertVersion(ushort[] version, out ushort[] normalizeVersion)
        {
            normalizeVersion = new ushort[ConstValues.NormalizedVersionSize];

            if ((version[1] & ConstValues.HighBitMask) != 0)
            {
                normalizeVersion[0] = (ushort)((version[0] & ConstValues.HighByteMask) >> ConstValues.StepDistanceOfOneByte);
                normalizeVersion[1] = (ushort)(version[0] & ConstValues.LowByteMask);
                normalizeVersion[2] = (ushort)(version[1] & (~ConstValues.HighBitMask));
                normalizeVersion[3] = version[2];
            }
            else
            {
                normalizeVersion[0] = version[0];
                normalizeVersion[1] = 0;
                normalizeVersion[2] = version[1];
                normalizeVersion[3] = version[2];
            }
        }
    }
}