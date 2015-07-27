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
    /// RopSynchronizationConfigure request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationConfigureRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x70.
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
        /// The possible values for this enumeration are specified in [MS-OXCFXICS]. 
        /// This value controls the type of synchronization.
        /// </summary>
        public byte SynchronizationType;
        
        /// <summary>
        /// The possible values are specified in [MS-OXCFXICS]. These values control the behavior of the operation.
        /// </summary>
        public byte SendOptions;

        /// <summary>
        /// The possible values are specified in [MS-OXCFXICS]. These flags control the behavior of the synchronization.
        /// </summary>
        public ushort SynchronizationFlags;

        /// <summary>
        /// This value specifies the length of the RestrictionData field.
        /// </summary>
        public ushort RestrictionDataSize;

        /// <summary>
        /// This field contains a restriction packet, as specified in [MS-OXCDATA] section 2.13. 
        /// The restriction specifies the filter for this synchronization object.
        /// </summary>
        public byte[] RestrictionData;

        /// <summary>
        /// The possible values are specified in [MS-OXCFXICS]. 
        /// These flags control the additional behavior of the synchronization.
        /// </summary>
        public uint SynchronizationExtraFlags;

        /// <summary>
        /// This value specifies how many tags are present in the PropertyTags field.
        /// </summary>
        public ushort PropertyTagCount;

        /// <summary>
        /// The format of the PropertyTag structure is specified in [MS-OXCDATA].
        /// This field specifies the property tags to be used for the synchronization process.
        /// </summary>
        public PropertyTag[] PropertyTags;

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
            serializedBuffer[index++] = this.OutputHandleIndex;
            serializedBuffer[index++] = this.SynchronizationType;
            serializedBuffer[index++] = this.SendOptions;

            Array.Copy(BitConverter.GetBytes((ushort)this.SynchronizationFlags), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.RestrictionDataSize), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.RestrictionDataSize > 0)
            {
                Array.Copy(this.RestrictionData, 0, serializedBuffer, index, this.RestrictionDataSize);
                index += this.RestrictionDataSize;
            }

            Array.Copy(BitConverter.GetBytes((uint)this.SynchronizationExtraFlags), 0, serializedBuffer, index, sizeof(uint));
            index += sizeof(uint);
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyTagCount), 0, serializedBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                Array.Copy(this.PropertyTags[i].Serialize(), 0, serializedBuffer, index, this.PropertyTags[i].Size());
                index += this.PropertyTags[i].Size();
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopSynchronizationConfigure request buffer structure.
        /// </summary>
        /// <returns>The size of RopSynchronizationConfigure request buffer structure.</returns>
        public int Size()
        {
            // 16 indicates sizeof (byte) * 6 + sizeof (UInt16) * 3 + sizeof (UInt32)
            int size = sizeof(byte) * 16;
            size += this.RestrictionDataSize;
            for (int i = 0; i < this.PropertyTagCount; i++)
            {
                size += this.PropertyTags[i].Size();
            }
                                                                                                                                                                                                
            return size;
        }
    }

    /// <summary>
    /// RopSynchronizationConfigure response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSynchronizationConfigureResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x70.
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
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopSynchronizationConfigureResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopSynchronizationConfigureResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}