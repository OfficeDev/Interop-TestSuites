namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopWritePerUserInformation request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopWritePerUserInformationRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x64.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// LongTermId structure. The structure specifies the folder for which to set per-user information. 
        /// The format of the LongTermId structure is specified in [MS-OXCDATA] section 2.2.1.3.1.
        /// </summary>
        public LongTermId FolderId;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether this operation specifies the end of the per-user information stream.
        /// </summary>
        public byte HasFinished;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the location in the per-user information stream to start writing.
        /// </summary>
        public uint DataOffset;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the Data field in bytes.
        /// </summary>
        public ushort DataSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the DataSize field. 
        /// This array is the per-user data to write.
        /// </summary>
        public byte[] Data;

        /// <summary>
        /// GUID. This field is present when the DataOffset field is 0x00000000 and the logon 
        /// associated with LogonId was created with the Private flag set (see [MS-OXCSTOR] for more information) 
        /// and is not present otherwise.
        /// </summary>
        public byte[] ReplGuid;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];

            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.InputHandleIndex;
            if (this.FolderId.Size() > 0)
            {
                Array.Copy(this.FolderId.Serialize(), 0, serializeBuffer, index, this.FolderId.Size());
                index += this.FolderId.Size();
            }

            serializeBuffer[index++] = this.HasFinished;

            Array.Copy(BitConverter.GetBytes((uint)this.DataOffset), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            Array.Copy(BitConverter.GetBytes((ushort)this.DataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            if (this.DataSize > 0)
            {
                Array.Copy(this.Data, 0, serializeBuffer, index, this.DataSize);
                index += this.DataSize;
            }

            if (this.DataOffset == 0) 
            {
                if (this.ReplGuid != null)
                {
                    // ReplGuid require 16 bytes.
                    Array.Copy(this.ReplGuid, 0, serializeBuffer, index, 16);
                    index += 16;
                }
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 10;
            if (this.FolderId.Size() > 0)
            {
                size += this.FolderId.Size();
            }

            if (this.DataSize > 0)
            {
                size += this.DataSize;
            }

            if (this.DataOffset == 0 && this.ReplGuid != null) 
            {
                size += 16;
            }

            return size;
        }
    }

    /// <summary>
    /// RopWritePerUserInformation response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopWritePerUserInformationResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x64.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation.
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
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopWritePerUserInformationResponse)Marshal.PtrToStructure(
                    responseBuffer,
                    typeof(RopWritePerUserInformationResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}