namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopCreateMessage request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCreateMessageRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x06.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored. 
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the output Server Object will be stored. 
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// 16-bit identifier. This value specifies the code page for the message.
        /// </summary>
        public ushort CodePageId;

        /// <summary>
        /// 64-bit identifier. This value identifies the parent folder.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// 8-bit bool. This value specifies whether the message is a Folder Associated Information message.
        /// </summary>
        public byte AssociatedFlag;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] resultBytes = new byte[Marshal.SizeOf(this)];
            IntPtr ptr = new IntPtr();
            ptr = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.StructureToPtr(this, ptr, true);
                Marshal.Copy(ptr, resultBytes, 0, Marshal.SizeOf(this));
                return resultBytes;
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
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
    /// RopCreateMessage response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCreateMessageResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x06.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit bool. This value specifies whether the MessageId field is present.
        /// </summary>
        public byte HasMessageId;

        /// <summary>
        /// 64-bit identifier. This field is present if HasMessageId is non-zero and is not present if it is zero. 
        /// This value is an identifier that is associated with the created message.
        /// </summary>
        public ulong? MessageId;

        /// <summary>
        /// De-serialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.OutputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.HasMessageId = ropBytes[index++];
                if (this.HasMessageId != 0)
                {
                    this.MessageId = (ulong)BitConverter.ToInt64(ropBytes, index);
                    index += sizeof(ulong);
                }
            }

            return index - startIndex;
        }
    }
}