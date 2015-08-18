namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopReadStream request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopReadStreamRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x2C.
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
        /// Unsigned 16-bit integer. The value of this field specifies the maximum number of bytes to read 
        /// if the value is not equal to 0xBABE; the MaximumByteCount field specifies the maximum number of bytes 
        /// to read if the value of ByteCount is equal to 0xBABE.
        /// </summary>
        public ushort ByteCount;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the maximum number of bytes to read if the 
        /// value of the ByteCount field is equal to 0xBABE. The MaximumByteCount field is present when 
        /// ByteCount is equal to 0xBABE and is not present otherwise. If MaximumByteCount is greater than 0x80000000, 
        /// then the RPC call SHOULD fail with error code 0x000004B6.
        /// </summary>
        public uint MaximumByteCount;

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

            Array.Copy(BitConverter.GetBytes((ushort)this.ByteCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.ByteCount == 0xBABE)
            {
                Array.Copy(BitConverter.GetBytes((uint)this.MaximumByteCount), 0, serializeBuffer, index, sizeof(uint));
                index += sizeof(uint);
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 5 indicates sizeof(byte) * 3 + sizeof(UInt16)
            int size = sizeof(byte) * 5;
            if (this.ByteCount == 0xBABE)
            {
                size += 4;
            }

            return size;
        }
    }

    /// <summary>
    /// RopReadStream response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopReadStreamResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x2C.
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
        /// Unsigned 16-bit integer. This value specifies the size, in bytes, of the Data field.
        /// </summary>
        public ushort DataSize;

        /// <summary>
        /// Array of bytes. These values are the bytes read from the stream. The size of this field, in bytes, 
        /// is specified by the DataSize field.
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
            this.DataSize = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            if (this.DataSize >= 0)
            {
                this.Data = new byte[this.DataSize];
                Array.Copy(ropBytes, index, this.Data, 0, this.DataSize);
                index += this.DataSize;
            }

            return index - startIndex;
        }
    }
}