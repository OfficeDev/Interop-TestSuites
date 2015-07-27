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
    /// RopSynchronizationImportDeletes request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportDeletesRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x74.
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
        /// This value specifies whether this operation consists of hierarchy or content deletions.
        /// </summary>
        public byte IsHierarchy;

        /// <summary>
        /// This value specifies the number of structures in the PropertyValues field.
        /// </summary>
        public ushort PropertyValueCount;

        /// <summary>
        /// The format of the TaggedPropertyValue structure is specified in [MS-OXCDATA] 
        /// and possible properties to be set are specified in [MS-OXCFXICS]. 
        /// These values are used to specify the folders or messages to delete.
        /// </summary>
        public TaggedPropertyValue[] PropertyValues;

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
            serializedBuffer[index++] = this.IsHierarchy;

            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyValueCount), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            for (int i = 0; i < this.PropertyValueCount; i++)
            {
                if (this.PropertyValues[i].Value != null)
                {
                    Array.Copy(this.PropertyValues[i].Serialize(), 0, serializedBuffer, index, this.PropertyValues[i].Size());
                    index += this.PropertyValues[i].Size();
                }
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopSynchronizationImportDeletes request buffer structure.
        /// </summary>
        /// <returns>The size of RopSynchronizationImportDeletes request buffer structure.</returns>
        public int Size()
        {
            // 6 indicates sizeof (byte) * 4 + sizeof (UInt16) * 2 
            int size = sizeof(byte) * 6;
            for (int i = 0; i < this.PropertyValueCount; i++)
            {
                if (this.PropertyValues[i].Value != null)
                {
                    size += this.PropertyValues[i].Size();
                }
            }

            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationImportDeletes response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportDeletesResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x74.
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
                this = (RopSynchronizationImportDeletesResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopSynchronizationImportDeletesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}