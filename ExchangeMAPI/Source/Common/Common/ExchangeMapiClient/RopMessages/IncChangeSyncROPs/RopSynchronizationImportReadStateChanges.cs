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
    /// This structure specifies the messages and associated read states to be changed.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct MessageReadState : ISerializable
    {
        /// <summary>
        /// This value specifies the size of the MessageId field.
        /// </summary>
        public ushort MessageIdSize;

        /// <summary>
        /// This value identifies the message to be marked as read or unread.
        /// </summary>
        public byte[] MessageId;

        /// <summary>
        /// This value specifies whether to mark the message as read or not.
        /// </summary>
        public byte MarkAsRead;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            // 0 indicates start index
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            Array.Copy(BitConverter.GetBytes((ushort)this.MessageIdSize), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.MessageIdSize > 0)
            {
                Array.Copy(this.MessageId, 0, serializedBuffer, index, this.MessageIdSize);
                index += this.MessageIdSize;
            }

            serializedBuffer[index++] = this.MarkAsRead;
            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of this MessageReadState structure.
        /// </summary>
        /// <returns>The size of this MessageReadState structure.</returns>
        public int Size()
        {
            // 3 indicates sizeof (byte) + sizeof (UInt16)
            int size = sizeof(byte) * 3;
            size += this.MessageIdSize;
            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationImportReadStateChanges request buffer structure.
    /// </summary>
    public struct RopSynchronizationImportReadStateChangesRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x80.
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
        /// This value specifies the size, in bytes, of the MessageReadStates field.
        /// </summary>
        public ushort MessageReadStateSize;

        /// <summary>
        /// List of MessageReadState structures. These values specify the messages and associated read states to be changed.
        /// </summary>
        public MessageReadState[] MessageReadStates;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            // 0 indicates start index
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            serializedBuffer[index++] = this.RopId;
            serializedBuffer[index++] = this.LogonId;
            serializedBuffer[index++] = this.InputHandleIndex;

            Array.Copy(BitConverter.GetBytes((ushort)this.MessageReadStateSize), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.MessageReadStates.Length; i++)
            {
                Array.Copy(this.MessageReadStates[i].Serialize(), 0, serializedBuffer, index, this.MessageReadStates[i].Size());
                index += this.MessageReadStates[i].Size();
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of this MessageReadState structure.
        /// </summary>
        /// <returns>The size of MessageReadState structure.</returns>
        public int Size()
        {
            // 5 indicates sizeof (byte) * 3 + sizeof (UInt16)
            int size = sizeof(byte) * 5;

            for (int i = 0; i < this.MessageReadStates.Length; i++)
            {
                size += this.MessageReadStates[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationImportReadStateChanges response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportReadStateChangesResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x80.
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
                this = (RopSynchronizationImportReadStateChangesResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopSynchronizationImportReadStateChangesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}