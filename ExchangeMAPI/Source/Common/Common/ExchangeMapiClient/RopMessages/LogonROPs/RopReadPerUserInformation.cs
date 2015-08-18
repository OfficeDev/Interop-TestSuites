namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// The structure of LongTermId
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct LongTermId : ISerializable, IDeserializable
    {
        /// <summary>
        /// A 128-bit unsigned integer identifying a Store object.
        /// </summary>
        public byte[] DatabaseGuid;

        /// <summary>
        /// An unsigned 48-bit integer identifying the folder within its Store object.
        /// </summary>
        public byte[] GlobalCounter;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];
            Array.Copy(this.DatabaseGuid, 0, serializeBuffer, index, sizeof(byte) * 16);
            index += sizeof(byte) * 16;
            Array.Copy(this.GlobalCounter, 0, serializeBuffer, index, sizeof(byte) * 6);
            index += sizeof(byte) * 6;
            Array.Copy(BitConverter.GetBytes(0), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            return sizeof(byte) * 24;
        }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.DatabaseGuid = new byte[16];
            Array.Copy(ropBytes, index, this.DatabaseGuid, 0, sizeof(byte) * 16);
            index += sizeof(byte) * 16;
            this.GlobalCounter = new byte[6];
            Array.Copy(ropBytes, index, this.GlobalCounter, 0, sizeof(byte) * 6);
            index += sizeof(byte) * 6;

            // 0 would hold 2 bytes.
            index += 2;
               
            return index - startIndex;
        }
    }

    /// <summary>
    /// RopReadPerUserInformation request buffer structure.
    /// </summary>
    public struct RopReadPerUserInformationRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x63.
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
        /// LongTermId structure. The structure specifies the folder for which to get per-user information. 
        /// The format of the LongTermId structure is specified in [MS-OXCDATA] section 2.2.1.3.1.
        /// </summary>
        public LongTermId FolderId;

        /// <summary>
        /// Reserved. This field is not used and is ignored by the server. This field MUST be set to 0x00.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the location at which to start reading within the per-user information stream.
        /// </summary>
        public uint DataOffset;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the maximum number of bytes of per-user information to be retrieved.
        /// </summary>
        public ushort MaxDataSize;

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

            serializeBuffer[index++] = this.Reserved;

            Array.Copy(BitConverter.GetBytes((uint)this.DataOffset), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            Array.Copy(BitConverter.GetBytes((ushort)this.MaxDataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 10;
            size += this.FolderId.Size();
            return size;
        }
    }

    /// <summary>
    /// RopReadPerUserInformation response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopReadPerUserInformationResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x63.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether this operation reached the end of the per-user information stream.
        /// </summary>
        public byte HasFinished;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the Data field.
        /// </summary>
        public ushort DataSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the DataSize field. This field contains the per-user data that is returned.
        /// </summary>
        public byte[] Data;

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
            this.InputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.HasFinished = ropBytes[index++];
                this.DataSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.DataSize > 0)
                {
                    this.Data = new byte[this.DataSize];
                    Array.Copy(ropBytes, index, this.Data, 0, this.DataSize);
                    index += this.DataSize;
                }
            }

            return index - startIndex;
        }
    }
}