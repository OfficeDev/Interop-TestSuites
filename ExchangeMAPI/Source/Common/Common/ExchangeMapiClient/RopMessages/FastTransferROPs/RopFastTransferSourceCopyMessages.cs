namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopFastTransferSourceCopyMessages request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceCopyMessagesRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4B.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the output Server Object will be stored.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the number of identifiers in the MessageIds field.
        /// </summary>
        public ushort MessageIdCount;

        /// <summary>
        /// The number of identifiers contained in this field is specified by the MessageIdCount field. 
        /// This list specifies the messages to copy.
        /// </summary>
        public ulong[] MessageIds;

        /// <summary>
        /// This flag defines the parameter to control the type of FastTransfer download operation. The possible values are specified in [MS-OXCFXICS].
        /// </summary>
        public byte CopyFlags;

        /// <summary>
        /// This flag defines the data representation parameters of the download operation. The possible values are specified in [MS-OXCFXICS].
        /// </summary>
        public byte SendOptions;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            // Set start index is 0
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];

            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.InputHandleIndex;
            serializeBuffer[index++] = this.OutputHandleIndex;

            // 0 indicates start index
            Array.Copy(BitConverter.GetBytes((ushort)this.MessageIdCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.MessageIdCount; i++)
            {
                // 0 indicates start index
                Array.Copy(BitConverter.GetBytes((ulong)this.MessageIds[i]), 0, serializeBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
            }

            serializeBuffer[index++] = this.CopyFlags;
            serializeBuffer[index++] = this.SendOptions;

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of RopFastTransferSourceCopyMessages request buffer structure.
        /// </summary>
        /// <returns>The size of RopFastTransferSourceCopyMessages request buffer structure.</returns>
        public int Size()
        {
            // 8 indicates sizeof (byte) * 6 + sizeof (Uint16)
            int size = sizeof(byte) * 8;
            for (int i = 0; i < this.MessageIdCount; i++)
            {
                size += sizeof(ulong);
            }
                                                                                                                                                                                                
            return size;
        }
    }

    /// <summary>
    /// RopFastTransferSourceCopyMessages response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceCopyMessagesResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4B.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            // Get the responseBuffer
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopFastTransferSourceCopyMessagesResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopFastTransferSourceCopyMessagesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}