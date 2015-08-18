namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopMoveCopyMessages request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopMoveCopyMessagesRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x33.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the source Server Object is stored.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the destination Server Object is stored.
        /// </summary>
        public byte DestHandleIndex;

        /// <summary>
        /// This value specifies the size of the MessageIds field.
        /// </summary>
        public ushort MessageIdCount;

        /// <summary>
        /// These identifiers specify which messages to move or copy.
        /// </summary>
        public ulong[] MessageIds;

        /// <summary>
        /// This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.
        /// </summary>
        public byte WantAsynchronous;

        /// <summary>
        /// This value specifies whether the operation is a copy or a move.
        /// </summary>
        public byte WantCopy;

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
            serializeBuffer[index++] = this.SourceHandleIndex;
            serializeBuffer[index++] = this.DestHandleIndex;

            Array.Copy(BitConverter.GetBytes((ushort)this.MessageIdCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.MessageIdCount; i++)
            {
                Array.Copy(BitConverter.GetBytes((ulong)this.MessageIds[i]), 0, serializeBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
            }

            serializeBuffer[index++] = this.WantAsynchronous;
            serializeBuffer[index++] = this.WantCopy;

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 8;
            for (int i = 0; i < this.MessageIdCount; i++)
            {
                size += sizeof(ulong);
            }
                                                                                                                                                                                                
            return size;
        }
    }

    /// <summary>
    /// RopMoveCopyMessages response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopMoveCopyMessagesResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x33.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the SourceHandleIndex specified in the request.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. 
        /// For successful response, this field is set to a value other than 0x00000503.
        /// For Null Destination Failure  response, this field is set to 0x00000503. 
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Unsigned 32-bit integer. This index MUST be set to the DestHandleIndex specified in the request.
        /// </summary>
        public uint DestHandleIndex;

        /// <summary>
        /// This value indicates whether the operation was only partially completed.
        /// </summary>
        public byte PartialCompletion;

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

            // Success response doesn't contain DestHandleIndex field
            // 0x00000503 indicates NullDestinationObject(MS-OXCDATA section 2.4)
            if (this.ReturnValue != 0x00000503) 
            {
                this.PartialCompletion = ropBytes[index++];
            }
            else
            {
                this.DestHandleIndex = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
                this.PartialCompletion = ropBytes[index++];
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}