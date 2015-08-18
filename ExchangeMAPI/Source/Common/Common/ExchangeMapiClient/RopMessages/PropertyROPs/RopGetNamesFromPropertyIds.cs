namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopGetNamesFromPerpertyIds request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetNamesFromPropertyIdsRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x55.
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
        /// Unsigned 16-bit integer. This value specifies the number of PropertyId structures in the PropertyIds field.
        /// </summary>
        public ushort PropertyIdCount;

        /// <summary>
        /// Array of PropertyId structures. The number of structures contained in this field is specified 
        /// by the PropertyIdCount field. The format of the PropertyId structure is specified in [MS-OXCDATA]. 
        /// This field specifies which property IDs for which to retrieve names.
        /// </summary>
        public PropertyId[] PropertyIds;

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
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyIdCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.PropertyIdCount; i++)
            {
                Array.Copy(this.PropertyIds[i].Serialize(), 0, serializeBuffer, index, this.PropertyIds[i].Size());
                index += this.PropertyIds[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 5 indicates sizeof(byte) * 3 + sizeof(UInt16)
            int size = sizeof(byte) * 5;
            if (this.PropertyIdCount > 0)
            {
                for (int i = 0; i < this.PropertyIdCount; i++)
                {
                    size += this.PropertyIds[i].Size();
                }
            }

            return size;
        }
    }

    /// <summary>
    /// RopGetNamesFromPropertyIds response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetNamesFromPropertyIdsResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x55.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of structures in the PropertyNames field.
        /// </summary>
        public ushort PropertyNameCount;

        /// <summary>
        /// List of PropertyName structures. The number of structures contained in this field is specified 
        /// by the PropertyNameCount field. The format of the PropertyName structure is specified in [MS-OXCDATA]. 
        /// This field lists the property names for the property IDs requested.
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
                this.PropertyNameCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.PropertyNameCount > 0)
                {
                    this.PropertyNames = new PropertyName[this.PropertyNameCount];
                    for (int i = 0; i < this.PropertyNameCount; i++)
                    {
                        index += this.PropertyNames[i].Deserialize(ropBytes, index);
                    }
                }
            }

            return index - startIndex;
        }
    }
}