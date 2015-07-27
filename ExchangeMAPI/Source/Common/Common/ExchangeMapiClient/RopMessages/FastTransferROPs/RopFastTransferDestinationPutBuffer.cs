//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopFastTransferDestinationPutBuffer request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferDestinationPutBufferRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x54.
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
        /// This value specifies the size of the TransferData field.
        /// </summary>
        public ushort TransferDataSize;

        /// <summary>
        /// This array contains the data to be uploaded to the destination fast transfer object.
        /// </summary>
        public byte[] TransferData;
        
        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
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

        /// <summary>
        /// Return the size of RopFastTransferDestinationPutBuffer request buffer structure.
        /// </summary>
        /// <returns>The size of RopFastTransferDestinationPutBuffer request buffer structure.</returns>
        public int Size()
        {
            // 5 indicates sizeof (byte) * 3 + sizeof (Uint16)
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
    /// RopFastTransferDestinationPutBuffer response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferDestinationPutBufferResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x54.
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
        /// This value specifies the approximate total number of steps to be completed in the current operation.
        /// </summary>
        public ushort TotalStepCount;

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
            // Get the responseBuffer
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopFastTransferDestinationPutBufferResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopFastTransferDestinationPutBufferResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}