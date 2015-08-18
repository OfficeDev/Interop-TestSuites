namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopQueryRows request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopQueryRowsRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x15.
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
        public byte QueryRowsFlags;

        /// <summary>
        /// 8-bit Boolean. This value specifies the direction to read rows.
        /// </summary>
        public byte ForwardRead;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of requested rows.
        /// </summary>
        public ushort RowCount;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializeBuffer = new byte[Marshal.SizeOf(this)];
            IntPtr requestBuffer = new IntPtr();
            requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.StructureToPtr(this, requestBuffer, true);
                Marshal.Copy(requestBuffer, serializeBuffer, 0, Marshal.SizeOf(this));
                return serializeBuffer;
            }
            finally
            {
                Marshal.FreeHGlobal(requestBuffer);
            }
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// RopQueryRows response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopQueryRowsResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x15.
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
        /// This value specifies current location of the cursor.
        /// </summary>
        public byte Origin;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of structures in the RowData field.
        /// </summary>
        public ushort RowCount;

        /// <summary>
        /// List of PropertyRow structures. The number of structures contained in this field is specified 
        /// by the RowCount field. The format of the PropertyRow structure is specified in [MS-OXCDATA] and
        /// the columns used for these rows were those previously set on this table by a RopSetColumns request.
        /// </summary>
        public PropertyRowSet RowData;

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
                this.Origin = ropBytes[index++];
                this.RowCount = BitConverter.ToUInt16(ropBytes, index);
                index += sizeof(ushort);

                // Parse RowData
                // Add RowData bytes into Context
                Context.Instance.PropertyBytes = ropBytes;
                Context.Instance.CurIndex = index;

                // Allocate PropretyRowSetNode to parse RowData
                this.RowData = new PropertyRowSet
                {
                    // Set row count
                    Count = this.RowCount
                };
                
                this.RowData.Parse(Context.Instance);

                // Context.Instance.CurIndex indicates the already deserialized bytes' index
                index = Context.Instance.CurIndex;
            }
            else
            {
                this.Origin = 0;
                this.RowCount = 0;
                this.RowData = null;
            }

            return index - startIndex;
        }
    }
}