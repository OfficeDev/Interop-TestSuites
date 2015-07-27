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
    /// RopSetReadFlags request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetReadFlagsRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x66.
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
        /// 8-bit Boolean. This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.
        /// </summary>
        public byte WantAsynchronous; 

        /// <summary>
        /// 8-bit flags structure. The possible values for these flags are specified in [MS-OXCMSG]. These flags specify the flags to set.
        /// </summary>
        public byte ReadFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of identifiers in the MessageIds field.
        /// </summary>
        public ushort MessageIdCount;

        /// <summary>
        /// Array of 64-bit identifiers. The number of identifiers contained in this field is specified by the MessageIdCount field. 
        /// These IDs specify the messages that are to have their read flags changed.
        /// </summary>
        public ulong[] MessageIds;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            serializedBuffer[index++] = this.RopId;
            serializedBuffer[index++] = this.LogonId;
            serializedBuffer[index++] = this.InputHandleIndex;
            serializedBuffer[index++] = this.WantAsynchronous;
            serializedBuffer[index++] = this.ReadFlags;

            Array.Copy(BitConverter.GetBytes((short)this.MessageIdCount), 0, serializedBuffer, index, sizeof(ushort));
            index += 2;

            for (int i = 0; i < this.MessageIdCount; i++)
            {
                Array.Copy(BitConverter.GetBytes((ulong)this.MessageIds[i]), 0, serializedBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
            }
           
            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopSetReadFlags request buffer structure.
        /// </summary>
        /// <returns>The size of RopSetReadFlags request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (byte) * 5 + sizeof (UInt16) 
            int size = sizeof(byte) * 7;
            for (int i = 0; i < this.MessageIdCount; i++)
            {
                size += sizeof(ulong);
            }

            return size;
        }
    }

    /// <summary>
    /// RopSetReadFlags response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetReadFlagsResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x66.
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
        /// 8-bit Boolean. This value indicates whether the operation was only partially completed. 
        /// The operation is partially completed if the server was unable to modify one or more of the Message objects that are 
        /// specified in the MessageIds field of the request.
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
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopSetReadFlagsResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopSetReadFlagsResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}