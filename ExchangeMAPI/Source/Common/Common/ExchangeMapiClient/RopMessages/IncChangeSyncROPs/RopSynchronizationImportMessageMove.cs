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
    /// RopSynchronizationImportMessageMove request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportMessageMoveRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x78.
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
        /// This value specifies the size of the SourceFolderId field.
        /// </summary>
        public uint SourceFolderIdSize;

        /// <summary>
        /// This value identifies the parent folder of the source message.
        /// </summary>
        public byte[] SourceFolderId;

        /// <summary>
        /// This value specifies the size of the SourceMessageId field.
        /// </summary>
        public uint SourceMessageIdSize;

        /// <summary>
        /// This value identifies the source message.
        /// </summary>
        public byte[] SourceMessageId;

        /// <summary>
        /// This value specifies the size of the PredecessorChangeList field.
        /// </summary>
        public uint PredecessorChangeListSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the PredecessorChangeListSize field.
        /// </summary>
        public byte[] PredecessorChangeList;

        /// <summary>
        /// This value specifies the size of the DestinationMessageId field.
        /// </summary>
        public uint DestinationMessageIdSize;

        /// <summary>
        /// This value identifies the destination message.
        /// </summary>
        public byte[] DestinationMessageId;

        /// <summary>
        /// This value specifies the size of the ChangeNumber field.
        /// </summary>
        public uint ChangeNumberSize;

        /// <summary>
        /// This field specifies the change number of the message.
        /// </summary>
        public byte[] ChangeNumber;

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

            Array.Copy(BitConverter.GetBytes((uint)this.SourceFolderIdSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.SourceFolderIdSize > 0)
            {
                Array.Copy(this.SourceFolderId, 0, serializedBuffer, index, this.SourceFolderIdSize);
                index += (int)this.SourceFolderIdSize;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.SourceMessageIdSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.SourceMessageIdSize > 0)
            {
                Array.Copy(this.SourceMessageId, 0, serializedBuffer, index, this.SourceMessageIdSize);
                index += (int)this.SourceMessageIdSize;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.PredecessorChangeListSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.PredecessorChangeListSize > 0)
            {
                Array.Copy(this.PredecessorChangeList, 0, serializedBuffer, index, this.PredecessorChangeListSize);
                index += (int)this.PredecessorChangeListSize;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.DestinationMessageIdSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.DestinationMessageIdSize > 0)
            {
                Array.Copy(this.DestinationMessageId, 0, serializedBuffer, index, this.DestinationMessageIdSize);
                index += (int)this.DestinationMessageIdSize;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.ChangeNumberSize), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            if (this.ChangeNumberSize > 0)
            {
                Array.Copy(this.ChangeNumber, 0, serializedBuffer, index, this.ChangeNumberSize);
                index += (int)this.ChangeNumberSize;
            }
           
            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopSynchronizationImportMessageMove request buffer structure.
        /// </summary>
        /// <returns>The size of RopSynchronizationImportMessageMove request buffer structure.</returns>
        public int Size()
        {
            // 23 indicates sizeof (byte) * 3 + sizeof (UInt32)*5 
            int size = sizeof(byte) * 23;
            size += (int)this.SourceFolderIdSize;
            size += (int)this.SourceMessageIdSize;
            size += (int)this.PredecessorChangeListSize;
            size += (int)this.DestinationMessageIdSize;
            size += (int)this.ChangeNumberSize;
            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationImportMessageMove response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportMessageMoveResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x78.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. 
        /// For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 64-bit identifier. This field should be set by server if success.
        /// </summary>
        public ulong? MessageId;

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
                                                                                                                                                                             
            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.MessageId = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}