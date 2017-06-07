
namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopWriteStreamExtended request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopWriteStreamExtendedRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0xA3.
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
        /// Unsigned 16-bit integer. This value specifies the size of the Data field.
        /// </summary>
        public ushort DataSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the DataSize field. 
        /// These values specify the bytes to be written to the stream.
        /// </summary>
        public byte[] Data;

        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];

            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.InputHandleIndex;

            Array.Copy(BitConverter.GetBytes((ushort)this.DataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.DataSize > 0)
            {
                Array.Copy(this.Data, 0, serializeBuffer, index, this.DataSize);
                index += this.DataSize;
            }

            return serializeBuffer;
        }

        public int Size()
        {
            // 5 indicates sizeof(byte) * 3 + sizeof(UInt16)
            int size = sizeof(byte) * 5;
            if (this.DataSize > 0)
            {
                size += this.DataSize;
            }

            return size;
        }
    }

    /// <summary>
    /// RopWriteStreamExtended response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopWriteStreamExtendedResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0xA3.
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
        /// Unsigned 32-bit integer. This value specifies the number of bytes actually written.
        /// </summary>
        public uint WrittenSize;
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.InputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);
            this.WrittenSize = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            return index - startIndex;
        }
    }
}
