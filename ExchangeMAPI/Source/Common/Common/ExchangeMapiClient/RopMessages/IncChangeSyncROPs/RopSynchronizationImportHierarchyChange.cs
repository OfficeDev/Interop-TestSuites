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
    /// RopSynchronizationImportHierarchyChange request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportHierarchyChangeRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x73.
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
        /// This value specifies the number of structures in the HierarchyValues field.
        /// </summary>
        public ushort HierarchyValueCount;

        /// <summary>
        /// The format of the TaggedPropertyValue structure is specified in [MS-OXCDATA] 
        /// and possible properties to be set are specified in [MS-OXCFXICS]. 
        /// These values are used to specify some hierarchy related properties of the folder.
        /// </summary>
        public TaggedPropertyValue[] HierarchyValues;

        /// <summary>
        /// This value specifies the number of structures in the PropertyValues field.
        /// </summary>
        public ushort PropertyValueCount;

        /// <summary>
        /// The format of the TaggedPropertyValue structure is specified in [MS-OXCDATA]. 
        /// These values are used to specify folder properties.
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

            Array.Copy(BitConverter.GetBytes((ushort)this.HierarchyValueCount), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            for (int i = 0; i < this.HierarchyValueCount; i++)
            {
                if (this.HierarchyValues[i].Value != null)
                {
                    Array.Copy(this.HierarchyValues[i].Serialize(), 0, serializedBuffer, index, this.HierarchyValues[i].Size());
                    index += this.HierarchyValues[i].Size();
                }
            }

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
        /// Return the size of RopSynchronizationImportHierarchyChange request buffer structure.
        /// </summary>
        /// <returns>The size of RopSynchronizationImportHierarchyChange request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (byte) * 3 + sizeof (UInt16) * 2
            int size = sizeof(byte) * 7;
            for (int i = 0; i < this.HierarchyValueCount; i++)
            {
                if (this.HierarchyValues[i].Value != null)
                {
                    size += this.HierarchyValues[i].Size();
                }
            }

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
    /// RopSynchronizationImportHierarchyChange response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationImportHierarchyChangeResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x73.
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
        public ulong FolderId;

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
                this.FolderId = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += sizeof(ulong);
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}