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
    /// RopFastTransferSourceCopyTo request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceCopyToRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4D.
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
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the output Server Object will be stored.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the level at which the copy is occurring.
        /// </summary>
        public byte Level;

        /// <summary>
        /// This flag defines the parameter to control the type of FastTransfer download operation. The possible values are specified in [MS-OXCFXICS].
        /// </summary>
        public uint CopyFlags;

        /// <summary>
        /// This flag defines the data representation parameters of the download operation. The possible values are specified in [MS-OXCFXICS].
        /// </summary>
        public byte SendOptions;

        /// <summary>
        /// This value specifies the number of structures in the PropertyTags field.
        /// </summary>
        public ushort PropertyTagCount;

        /// <summary>
        /// This array specifies the properties to exclude during the copy.
        /// </summary>
        public PropertyTag[] PropertyTags;

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
            serializeBuffer[index++] = this.OutputHandleIndex;
            serializeBuffer[index++] = this.Level;
                                                                                                                                                                             
            // 0 indicates start index
            Array.Copy(BitConverter.GetBytes((uint)this.CopyFlags), 0, serializeBuffer, index, sizeof(uint));
                                                                                                                                                                             
            // 4 indicates UInt32 bit occupies 4 bytes location  
            index += 4;
            serializeBuffer[index++] = this.SendOptions;

            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyTagCount), 0, serializeBuffer, index, sizeof(ushort));
                                                                                                                                                                             
            // 2 indicates UInt16 bit occupies 2 bytes location  
            index += 2;

            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                // 0 indicates start index
                Array.Copy(this.PropertyTags[i].Serialize(), 0, serializeBuffer, index, this.PropertyTags[i].Size());
                index += this.PropertyTags[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of RopFastTransferSourceCopyTo request buffer structure.
        /// </summary>
        /// <returns>The size of RopFastTransferSourceCopyTo request buffer structure.</returns>
        public int Size()
        {
            // 12 indicates sizeof (byte) * 6 + sizeof (Uint16)+sizeof(Uint32)
            int size = sizeof(byte) * 12;
            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                size += this.PropertyTags[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopFastTransferSourceCopyTo response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFastTransferSourceCopyToResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x4D.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation.
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
            // Get the responseBuffer
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopFastTransferSourceCopyToResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopFastTransferSourceCopyToResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}