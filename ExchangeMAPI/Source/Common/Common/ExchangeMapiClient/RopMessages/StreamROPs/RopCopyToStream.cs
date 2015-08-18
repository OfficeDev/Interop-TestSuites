namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopCopyToStream request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyToStreamRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x3A.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the source Server Object is stored.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the destination Server Object is stored.
        /// </summary>
        public byte DestHandleIndex;

        /// <summary>
        /// Unsigned 64-bit integer. This value specifies the number of bytes to be copied.
        /// </summary>
        public ulong ByteCount;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializeBuffer = new byte[Marshal.SizeOf(this)];
            IntPtr requestBuffer = new IntPtr();
            requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.StructureToPtr(this, requestBuffer, true);
                Marshal.Copy(requestBuffer, serializeBuffer, 0, Marshal.SizeOf(this));
                return serializeBuffer;
            }
            finally
            {
                Marshal.FreeHGlobal(requestBuffer);
            }
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// RopCopyToStream response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyToStreamResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x3A.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the SourceHandleIndex specified in the request.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For this response, this field is set to a value other than 0x00000503.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Unsigned 64-bit integer. This value specifies the number of bytes read from the source object.
        /// </summary>
        public ulong ReadByteCount;

        /// <summary>
        /// Unsigned 64-bit integer. This value specifies the number of bytes written to the destination object.
        /// </summary>
        public ulong WrittenByteCount;

        /// <summary>
        /// Unsigned 32-bit integer. This index MUST be set to the DestHandleIndex specified in the request.
        /// </summary>
        public uint DestHandleIndex;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.SourceHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Check return value with 0x00000503.
            if (this.ReturnValue != 0x00000503)
            {
                this.ReadByteCount = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
                this.WrittenByteCount = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
            }
            else
            {
                this.DestHandleIndex = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
                this.ReadByteCount = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
                this.WrittenByteCount = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong); 
            }

            return index - startIndex;
        }
    }
}