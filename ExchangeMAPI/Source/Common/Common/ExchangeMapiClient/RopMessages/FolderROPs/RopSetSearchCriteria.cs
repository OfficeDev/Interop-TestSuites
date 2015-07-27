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
    /// RopSetSearchCriteria request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetSearchCriteriaRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x30.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the length of the RestrictionData field.
        /// </summary>
        public ushort RestrictionDataSize;

        /// <summary>
        /// This field contains a restriction packet, as specified in [MS-OXCDATA] section 2.13. The restriction specifies the filter for this search folder.
        /// </summary>
        public byte[] RestrictionData;

        /// <summary>
        /// This value specifies the number of IDs in the FolderIds field.
        /// </summary>
        public ushort FolderIdCount;

        /// <summary>
        /// This field contains identifiers that specify which folders are searched.
        /// </summary>
        public ulong[] FolderIds;

        /// <summary>
        /// These flags control the search for a search folder.
        /// </summary>
        public uint SearchFlags;

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

            // Serialize RestrictionDataSize
            Array.Copy(BitConverter.GetBytes((ushort)this.RestrictionDataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            if (this.RestrictionDataSize > 0)
            {
                // Serialize RestrictionData
                Array.Copy(this.RestrictionData, 0, serializeBuffer, index, this.RestrictionDataSize);
                index += this.RestrictionDataSize;
            }

            // Serialize FolderIdCount
            Array.Copy(BitConverter.GetBytes((ushort)this.FolderIdCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.FolderIdCount; i++)
            {
                // Serialize FolderIds
                Array.Copy(BitConverter.GetBytes((ulong)this.FolderIds[i]), 0, serializeBuffer, index, sizeof(ulong));
                index += sizeof(ulong);
            }

            // Serialize SearchFlags
            Array.Copy(BitConverter.GetBytes((uint)this.SearchFlags), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 11 indicates sizeof(byte) * 3 + sizeof(UInt16) * 2 + sizeof(UInt32)
            int size = sizeof(byte) * 11;
            size += this.RestrictionDataSize;
            for (int i = 0; i < this.FolderIdCount; i++)
            {
                size += sizeof(ulong);
            }

            return size;
        }
    }

    /// <summary>
    /// RopSetSearchCriteria response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetSearchCriteriaResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x30.
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
            // Get the responseBuffer
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopSetSearchCriteriaResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopSetSearchCriteriaResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}