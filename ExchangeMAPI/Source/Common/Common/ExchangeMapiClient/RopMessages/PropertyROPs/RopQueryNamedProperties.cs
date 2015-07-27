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
    /// RopQueryNamedProperties request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopQueryNamedPropertiesRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x5F.
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
        /// 8-bit flags structure. The possible values are specified in [MS-OXCPRPT]. 
        /// These flags control how this remote operation behaves.
        /// </summary>
        public byte QueryFlags;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the PropertyGuid field is present.
        /// </summary>
        public byte HasGuid;

        /// <summary>
        /// 128-bit GUID. This field is present if HasGuid is non-zero and is not present if the value 
        /// of the HasGuid field is zero. This value specifies the subset of named properties to be returned.
        /// </summary>
        public byte[] PropertyGuid;

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
            serializeBuffer[index++] = this.QueryFlags;
            serializeBuffer[index++] = this.HasGuid;
            if (this.HasGuid != 0)
            {
                // PropertyGuid holds 16 bytes.
                Array.Copy(this.PropertyGuid, 0, serializeBuffer, index, 16);
                index += 16;
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 5;
            if (this.HasGuid != 0)
            {
                // PropertyGuid holds 16 bytes.
                size += 16;
            }

            return size;
        }
    }

    /// <summary>
    /// RopQueryNamedProperties response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopQueryNamedPropertiesResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x5F.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For this response, this field is set to 0x00000000 or 0x00040380.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of Ids and names in the PropertyIds and PropertyNames fields, respectively.
        /// </summary>
        public ushort IdCount;

        /// <summary>
        /// Array of PropertyId structures. The number of structures contained in this field is specified by the IdCount field. 
        /// The format of the PropertyId structure is specified in [MS-OXCDATA]. 
        /// This field lists the property IDs for which property names are given.
        /// </summary>
        public PropertyId[] PropertyIds;

        /// <summary>
        /// List of PropertyName structures. The number of structures contained in this field is specified by the IdCount field. 
        /// The format of the PropertyName structure is specified in [MS-OXCDATA]. 
        /// This field lists the property names for the property IDs specified in the PropertyIds field.
        /// </summary>
        public PropertyName[] PropertyNames;

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
                this.IdCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.IdCount > 0)
                {
                    this.PropertyIds = new PropertyId[this.IdCount];
                    for (int i = 0; i < this.IdCount; i++)
                    {
                        index += this.PropertyIds[i].Deserialize(ropBytes, index);
                    }

                    this.PropertyNames = new PropertyName[this.IdCount];
                    for (int i = 0; i < this.IdCount; i++)
                    {
                        index += this.PropertyNames[i].Deserialize(ropBytes, index);
                    }
                }
            }

            return index - startIndex;
        }
    }
}