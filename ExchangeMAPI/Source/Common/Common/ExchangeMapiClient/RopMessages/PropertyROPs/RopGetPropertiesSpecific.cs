namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopGetPropertiesSpecific request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetPropertiesSpecificRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x07.
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
        /// Unsigned 16-bit integer. This value specifies the maximum size allowed for a property value returned.
        /// </summary>
        public ushort PropertySizeLimit;

        /// <summary>
        /// 16-bit Boolean. This value specifies whether to return string properties in Unicode.
        /// </summary>
        public ushort WantUnicode;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies how many tags are present in PropertyTags.
        /// </summary>
        public ushort PropertyTagCount;

        /// <summary>
        /// Array of PropertyTag structures. The number of structures contained in this field is specified 
        /// by the PropertyTagCount field. The format of the PropertyTag structure is specified in [MS-OXCDATA]. 
        /// This field specifies the properties requested.
        /// </summary>
        public PropertyTag[] PropertyTags;

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
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertySizeLimit), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.WantUnicode), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.PropertyTagCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.PropertyTagCount > 0)
            {
                IntPtr requestBuffer = new IntPtr();
                requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(PropertyTag)));
                try
                {
                    Context.Instance.Init();
                    foreach (PropertyTag propTag in this.PropertyTags)
                    {
                        Marshal.StructureToPtr(propTag, requestBuffer, true);
                        Marshal.Copy(requestBuffer, serializeBuffer, index, Marshal.SizeOf(typeof(PropertyTag)));
                        index += Marshal.SizeOf(typeof(PropertyTag));

                        // Insert properties into Context
                        Context.Instance.Properties.Add(new Property((PropertyType)propTag.PropertyType));
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(requestBuffer);
                }
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = (sizeof(byte) * 3) + (sizeof(ushort) * 3);
            if (this.PropertyTagCount > 0)
            {
                size += this.PropertyTags.Length * Marshal.SizeOf(typeof(PropertyTag));
            }

            return size;
        }
    }

    /// <summary>
    /// RopGetPropertiesSpecific response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetPropertiesSpecificResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x07.
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
        /// PropertyRow structure. The format of the PropertyRow structure is specified in [MS-OXCDATA] 
        /// and the columns used for these rows were those specified in the PropertyTags field in the request.
        /// </summary>
        public PropertyRow RowData;

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
            this.ReturnValue = BitConverter.ToUInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                // Parse RowData
                // Add RowData bytes into Context
                Context.Instance.PropertyBytes = ropBytes;
                Context.Instance.CurIndex = index;

                // Allocate PropretyRowSetNode to parse RowData
                this.RowData = new PropertyRow();

                // Set row count
                this.RowData.Parse(Context.Instance);

                // Context.Instance.CurIndex indicates the already deserialized bytes' index
                index = Context.Instance.CurIndex;
            }
            else
            {
                this.RowData = null;
            }

            return index - startIndex;
        }
    }
}