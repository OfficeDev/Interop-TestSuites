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
    /// RopCopyFolder request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyFolderRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x36.
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
        /// This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.
        /// </summary>
        public byte WantAsynchronous;

        /// <summary>
        /// This value specifies that the copy is recursive.
        /// </summary>
        public byte WantRecursive;

        /// <summary>
        /// This value specifies whether the NewFolderName field is specified in Unicode or ASCII.
        /// </summary>
        public byte UseUnicode;

        /// <summary>
        /// This value identifies the folder to be copied.
        /// </summary>
        public ulong FolderId;

        /// <summary>
        /// If the UseUnicode field is non-zero, then the string is composed of Unicode characters. 
        /// Otherwise, the string is composed of ASCII characters. This string specifies the name for the new copied folder.
        /// </summary>
        public byte[] NewFolderName;

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
            serializeBuffer[index++] = this.SourceHandleIndex;
            serializeBuffer[index++] = this.DestHandleIndex;
            serializeBuffer[index++] = this.WantAsynchronous;
            serializeBuffer[index++] = this.WantRecursive;
            serializeBuffer[index++] = this.UseUnicode;

            Array.Copy(BitConverter.GetBytes((ulong)this.FolderId), 0, serializeBuffer, index, sizeof(ulong));
            index += sizeof(ulong);

            if (this.NewFolderName != null)
            {
                Array.Copy(this.NewFolderName, 0, serializeBuffer, index, this.NewFolderName.Length);
                index += this.NewFolderName.Length;
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 15 indicates sizeof (byte) * 7 + sizeof (Uint64)
            int size = sizeof(byte) * 15;
            if (this.NewFolderName != null)
            {
                size += this.NewFolderName.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopCopyFolder response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopCopyFolderResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x36.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the SourceHandleIndex specified in the request.
        /// </summary>
        public byte SourceHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. 
        /// For this response, this field is set to a value other than 0x00000503.
        /// For Null Destination Failure response, this field is set to 0x00000503.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table 
        /// where the handle for the destination Server Object is stored.
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
            index += 4;

            // Success response doesn't contain DestHandleIndex field
            // 0x00000503 indicates NullDestinationObject(MS-OXCDATA section 2.4)
            if (this.ReturnValue != 0x00000503) 
            {
                this.PartialCompletion = ropBytes[index++];
            }
            else
            {
                this.DestHandleIndex = (uint)BitConverter.ToInt32(ropBytes, index);
                                                                                                                                                                             
                // 4 indicates sizeof(UInt32)
                index += 4;
                this.PartialCompletion = ropBytes[index++];
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}