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
    /// RopTransportNewMail request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopTransportNewMailRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x51.
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
        /// 64-bit identifier. This value identifies the new Message object.
        /// </summary>
        public ulong MessageId;

        /// <summary>
        /// 64-bit identifier. This value identifies the folder of the new Message object.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// Null-terminated ASCII string. This string specifies the message class of the new Message object.
        /// </summary>
        public byte[] MessageClass;

        /// <summary>
        /// 32-bit flags structure. The possible values are specified in [MS-OXCMSG]. 
        /// This field contains the message flags of the new message object.
        /// </summary>
        public uint MessageFlags;

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

            Array.Copy(BitConverter.GetBytes((ulong)this.MessageId), 0, serializeBuffer, index, sizeof(ulong));
            index += sizeof(ulong);
            Array.Copy(BitConverter.GetBytes((ulong)this.FolderId), 0, serializeBuffer, index, sizeof(ulong));
            index += sizeof(ulong);

            if (this.MessageClass != null)
            {
                Array.Copy(this.MessageClass, 0, serializeBuffer, index, this.MessageClass.Length);
                index += this.MessageClass.Length;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.MessageFlags), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 23 indicates sizeof(byte) * 3 + sizeof(UInt64) * 2 + sizeof(UInt32)
            int size = sizeof(byte) * 23;
            if (this.MessageClass != null)
            {
                size += this.MessageClass.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopTransportNewMail response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopTransportNewMailResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x51.
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
                this = (RopTransportNewMailResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopTransportNewMailResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}