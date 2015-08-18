namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopFastTransferSourceGetBuffer request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceGetBufferRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4E.
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
        /// This value specifies the buffer size requested.
        /// </summary>
        public ushort BufferSize;

        /// <summary>
        /// This field is present when the BufferSize field is set to 0xBABE. 
        /// This value specifies the maximum size limit when the server determines the buffer size.
        /// </summary>
        public ushort? MaximumBufferSize;

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

            // 0 indicates start index
            Array.Copy(BitConverter.GetBytes((ushort)this.BufferSize), 0, serializeBuffer, index, sizeof(ushort));
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location
            index += 2;
                                                                                                                                                                             
            // The field MaximumBufferSize is present when the BufferSize is set to 0xBABE.
            if (this.BufferSize == 0xBABE)
            {
                Array.Copy(BitConverter.GetBytes((ushort)this.MaximumBufferSize), 0, serializeBuffer, index, sizeof(ushort));
                                                                                                                                                                             
                // 2 indicates UInt16 bit occupies 2 bytes location
                index += 2;
            }
                                                                                                                                                                                                
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of  RopFastTransferSourceGetBuffer request buffer structure.
        /// </summary>
        /// <returns>The size of RopFastTransferSourceGetBuffer request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (byte) * 3 + 2*sizeof (Uint16)
            int size = sizeof(byte) * 5;
            if (this.BufferSize == 0xBABE)
            {
                size += 2;
            }

            return size;
        }
    }

    /// <summary>
    /// RopFastTransferSourceGetBuffer response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceGetBufferResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4E.
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
        public ushort InProgressCount;

        /// <summary>
        /// This value specifies the approximate number of steps to be completed in the current operation.
        /// </summary>
        public ushort TotalStepCount;

        /// <summary>
        /// Reserved. The server MUST set this field to 0x00.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// This value specifies the size of the TransferBuffer field.
        /// </summary>
        public ushort TransferBufferSize;

        /// <summary>
        /// This field is present if the ReturnValue is not 0x00000480 and is not present otherwise.
        /// </summary>
        public byte[] TransferBuffer;

        /// <summary>
        /// This field is present if the ReturnValue is 0x00000480 and is not present otherwise. 
        /// This value specifies the number of milliseconds for the client to wait before trying this operation again.
        /// </summary>
        public uint? BackoffTime;

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
                                                                                                                                                                             
            // 4 indicates UInt32 bit occupies 4 bytes location
            index += 4;
            this.TransferStatus = (ushort)BitConverter.ToInt16(ropBytes, index);
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location
            index += 2;
            this.InProgressCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location
            index += 2;
            this.TotalStepCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location
            index += 2;
            this.Reserved = ropBytes[index++];
            this.TransferBufferSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location
            index += 2;

            this.TransferBuffer = new byte[this.TransferBufferSize];
                                                                                                                                                                             
            // 0 indicates start index
            Array.Copy(ropBytes, index, this.TransferBuffer, 0, this.TransferBufferSize);
            index += this.TransferBufferSize;
                                                                                                                                                                             
            // The field BackOffTime is present if the ReturnValue is 0x480
            if (this.ReturnValue == 0x00000480) 
            {
                this.BackoffTime = (uint)BitConverter.ToInt32(ropBytes, index);
                                                                                                                                                                             
                // 4 indicates UInt32 bit occupies 4 bytes location
                index += 4;
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}