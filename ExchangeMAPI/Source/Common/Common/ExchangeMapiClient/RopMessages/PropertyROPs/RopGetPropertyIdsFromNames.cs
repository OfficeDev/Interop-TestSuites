namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// This field specifies the property name requested.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct PropertyName : ISerializable, IDeserializable
    {
        /// <summary>
        ///  The possible values are listed in [MS-OXCDATA] 2.6.1.
        /// </summary>
        public byte Kind;

        /// <summary>
        /// The GUID that identifies the property set for the named property.
        /// </summary>
        public byte[] Guid;

        /// <summary>
        /// Present only if Kind = 0x00. An unsigned 32-bit integer that identifies the named property within its property set.
        /// </summary>
        public uint? LID;

        /// <summary>
        /// Present only if Kind = 0x01. A single byte giving the number of bytes in the Name string that follows it.
        /// </summary>
        public byte? NameSize;

        /// <summary>
        /// Present only if Kind = 0x01. A Unicode (UTF-16) string, followed by two zero bytes as a null terminator, 
        /// that identifies the property within its property set.
        /// </summary>
        public byte[] Name;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] resultBytes = new byte[this.Size()];
            resultBytes[index++] = this.Kind;
            if (this.Guid != null)
            {
                // Guid holds 16 bytes.
                Array.Copy(this.Guid, 0, resultBytes, index, 16);
                index += 16;
            }

            if (this.Kind == (byte)TestSuites.Common.Kind.LidField)
            {
                Array.Copy(BitConverter.GetBytes((uint)this.LID), 0, resultBytes, index, sizeof(uint));
                index += sizeof(uint);
            }
            else if (this.Kind == (byte)TestSuites.Common.Kind.NameField)
            {
                resultBytes[index++] = (byte)this.NameSize;
                if (this.NameSize > 0)
                {
                    Array.Copy(this.Name, 0, resultBytes, index, (byte)this.NameSize);
                    index += (byte)this.NameSize;
                }
            }

            return resultBytes;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 17 indicate sizeof(byte) + (16 of Guid)
            int size = sizeof(byte) * 17;
            if (this.Kind == (byte)TestSuites.Common.Kind.LidField)
            {
                size += sizeof(byte) * 4;
            }
            else if (this.Kind == (byte)TestSuites.Common.Kind.NameField)
            {
                size += sizeof(byte);
                size += (byte)this.NameSize;
            }

            return size;
        }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.Kind = ropBytes[index++];

            // Guid holds 16 bytes.
            this.Guid = new byte[16];
            Array.Copy(ropBytes, index, this.Guid, 0, 16);
            index += 16;
            if (this.Kind == (byte)TestSuites.Common.Kind.LidField)
            {
                this.LID = (uint)BitConverter.ToInt32(ropBytes, index);
                index += 4;
            }
            else if (this.Kind == (byte)TestSuites.Common.Kind.NameField)
            {
                this.NameSize = ropBytes[index++];
                if (this.NameSize > 0)
                {
                    this.Name = new byte[(int)this.NameSize];
                    Array.Copy(ropBytes, index, this.Name, 0, (byte)this.NameSize);
                    index += (byte)this.NameSize;
                }
            }

            return index - startIndex;
        }
    }

    /// <summary>
    /// RopGetPropertyIdsFromNames request buffer structure.
    /// </summary>
    public struct RopGetPropertyIdsFromNamesRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x56.
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
        /// 8-bit flags structure. These flags control the behavior of this operation. 
        /// The possible values are specified in [MS-OXCPRPT].
        /// </summary>
        public byte Flags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of structures in the PropertyNames field.
        /// </summary>
        public ushort PropertyNameCount;

        /// <summary>
        /// List of PropertyName structures. The number of structures contained in this field is specified 
        /// by the PropertyNameCount field.
        /// </summary>
        public PropertyName[] PropertyNames;

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
            serializeBuffer[index++] = this.Flags;
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyNameCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            for (int i = 0; i < this.PropertyNameCount; i++)
            {
                Array.Copy(this.PropertyNames[i].Serialize(), 0, serializeBuffer, index, this.PropertyNames[i].Size());
                index += this.PropertyNames[i].Size();
            }
            
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 6 indicates sizeof(byte) * 4 + sizeof(UInt16)
            int size = sizeof(byte) * 6;
            if (this.PropertyNameCount > 0)
            {
                for (int i = 0; i < this.PropertyNameCount; i++)
                {
                    size += this.PropertyNames[i].Size();
                }
            }

            return size;
        }
    }

    /// <summary>
    /// Property Id
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct PropertyId : IDeserializable, ISerializable
    {
        /// <summary>
        /// Property Id.
        /// </summary>
        public ushort ID;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes((ushort)this.ID), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            return sizeof(ushort);
        }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.ID = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            return index - startIndex;
        }
    }

    /// <summary>
    /// RopGetPropertyIdsFromNames response buffer structure.
    /// </summary>
    public struct RopGetPropertyIdsFromNamesResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x56.
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
        /// Unsigned 16-bit integer. This value specifies the number of PropertyId structures in the PropertyIds field.
        /// </summary>
        public ushort PropertyIdCount;

        /// <summary>
        /// Array of PropertyId structures. The number of structures contained in this field is 
        /// specified by the PropertyIdCount field. The format of the PropertyId structure is specified in [MS-OXCDATA]. 
        /// This field lists the property IDs associated with the requested property names, 
        /// which are specified in the PropertyNames field of the request.
        /// </summary>
        public PropertyId[] PropertyIds;

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
                this.PropertyIdCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.PropertyIdCount > 0)
                {
                    this.PropertyIds = new PropertyId[this.PropertyIdCount];
                    for (int i = 0; i < this.PropertyIdCount; i++)
                    {
                        index += this.PropertyIds[i].Deserialize(ropBytes, index);
                    }
                }
            }

            return index - startIndex;
        }
    }
}