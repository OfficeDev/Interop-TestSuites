namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopRestrict request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopRestrictRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x14.
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
        /// 8-bit flags structure. The possible values are specified in [MS-OXCTABL]. These flags control this operation.
        /// </summary>
        public byte RestrictFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the length of the RestrictionData field.
        /// </summary>
        public ushort RestrictionDataSize;

        /// <summary>
        /// Restriction packet. The size of this field is RestrictionDataSize bytes. 
        /// This field contains a restriction packet, as specified in [MS-OXCDATA] section 2.13. 
        /// The restriction specifies the filter for this table.
        /// </summary>
        public byte[] RestrictionData;

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
            serializeBuffer[index++] = this.RestrictFlags;

            // Add this serialization for RestrictionDataSize.
            Array.Copy(BitConverter.GetBytes(this.RestrictionDataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            if (null != this.RestrictionData)
            {
                Array.Copy(this.RestrictionData, 0, serializeBuffer, index, this.RestrictionData.Length);
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = (sizeof(byte) * 4) + sizeof(ushort);
            if (null != this.RestrictionData)
            {
                size += this.RestrictionData.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopRestrict response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopRestrictResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x14.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For success response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit enumeration. The possible values for this enumeration are specified in [MS-OXCTABL]. 
        /// This value specifies the status of the table.
        /// </summary>
        public byte TableStatus;

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
                this.TableStatus = ropBytes[index++];
            }

            return index - startIndex;
        }
    }
}