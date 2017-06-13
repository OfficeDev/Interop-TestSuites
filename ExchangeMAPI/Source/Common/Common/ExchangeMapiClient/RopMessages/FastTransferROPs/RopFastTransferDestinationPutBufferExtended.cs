namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// RopWriteStreamExtended request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferDestinationPutBufferExtendedRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x9D.
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
        /// This value specifies the size of the TransferData field.
        /// </summary>
        public ushort TransferDataSize;

        /// <summary>
        /// This array contains the data to be uploaded to the destination fast transfer object.
        /// </summary>
        public byte[] TransferData;

        public byte[] Serialize()
        {
            // Set start index is 0
            int index = 0;

            // Get the serialized ROP request buffer
            byte[] serializeBuffer = new byte[this.Size()];

            serializeBuffer[index++] = this.RopId;
            serializeBuffer[index++] = this.LogonId;
            serializeBuffer[index++] = this.InputHandleIndex;

            // 0 indicates start index
            Array.Copy(BitConverter.GetBytes((ushort)this.TransferDataSize), 0, serializeBuffer, index, sizeof(ushort));

            // 2 indicates 16 bit occupies 2 bytes location  
            index += 2;

            // 0 indicates minimum size of TransferDataSize
            if (this.TransferDataSize > 0)
            {
                // 0 indicates start index
                Array.Copy(this.TransferData, 0, serializeBuffer, index, this.TransferDataSize);
                index += this.TransferDataSize;
            }

            return serializeBuffer;
        }

        public int Size()
        {
            // 5 indicates sizeof(byte) * 3 + sizeof(UInt16)
            int size = sizeof(byte) * 5;

            // 0 indicates minimum size of TransferDataSize
            if (this.TransferDataSize > 0)
            {
                size += this.TransferDataSize;
            }

            return size;
        }
    }

    /// <summary>
    /// RopFastTransferDestinationPutBufferExtended response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferDestinationPutBufferExtendedResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x9D.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// The possible values for this enumeration are specified in [MS-OXCFXICS]. 
        /// This value specifies the current status of the transfer.
        /// </summary>
        public ushort TransferStatus;

        /// <summary>
        /// This value specifies the number of steps that have been completed in the current operation.
        /// </summary>
        public uint InProgressCount;

        /// <summary>
        /// This value specifies the approximate total number of steps to be completed in the current operation.
        /// </summary>
        public uint TotalStepCount;

        /// <summary>
        /// Reserved. The server MUST set this field to 0x00.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// This value is the buffer size that was used.
        /// </summary>
        public ushort BufferUsedSize;

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
            this.TransferStatus = (ushort)BitConverter.ToUInt16(ropBytes, index);
            index += sizeof(ushort);
            this.InProgressCount = (uint)BitConverter.ToUInt32(ropBytes, index);
            index += sizeof(uint);
            this.TotalStepCount=(uint)BitConverter.ToUInt32(ropBytes, index);
            index += sizeof(uint);
            this.Reserved= ropBytes[index++];
            this.BufferUsedSize=(ushort)BitConverter.ToUInt16(ropBytes, index);
            index += sizeof(ushort);

            return index - startIndex;
        }
    }
}
