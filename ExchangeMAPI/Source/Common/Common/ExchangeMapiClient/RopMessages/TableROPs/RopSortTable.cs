namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopSortTable request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSortTableRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x13.
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
        public byte SortTableFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies how many SortOrder structures are present in the SortOrders field.
        /// </summary>
        public ushort SortOrderCount;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of category SortOrder structures in the SortOrders field.
        /// </summary>
        public ushort CategoryCount;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of expanded categories in the SortOrders field.
        /// </summary>
        public ushort ExpandedCount;

        /// <summary>
        /// Array of SortOrder structures. The number of structures contained in this field is specified by the SortOrderCount field.
        /// The format of the SortOrder structure is specified in [MS-OXCDATA]. 
        /// This field specifies the sort order for the rows in the table.
        /// </summary>
        public SortOrder[] SortOrders;

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
            serializeBuffer[index++] = this.SortTableFlags;
            Array.Copy(BitConverter.GetBytes((ushort)this.SortOrderCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.CategoryCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes((ushort)this.ExpandedCount), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            for (int i = 0; i < this.SortOrderCount; i++)
            {
                Array.Copy(this.SortOrders[i].Serialize(), 0, serializeBuffer, index, this.SortOrders[i].Size());
                index += this.SortOrders[i].Size();
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 10 indicates sizeof(byte) * 4 + sizeof(ushort) * 3
            int size = sizeof(byte) * 10;
            for (int i = 0; i < this.SortOrderCount; i++)
            {
                size += this.SortOrders[i].Size();
            }

            return size;
        }
    }

    /// <summary>
    /// The SortOrder structure describes one column that is part of a sort key for sorting rows of a table. 
    /// It gives both the column and the direction of the sort.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct SortOrder : ISerializable 
    {
        /// <summary>
        /// Identifies the data type of the column to sort on. If the property is multi-valued, 
        /// for example, the MultivalueFlag bit (0x1000) is set in the PropertyType, 
        /// then clients MUST also set the MultivalueInstance bit (0x2000). 
        /// In this case the server MUST generate one row for each individual value of a multi-value column, 
        /// and sort the table by individual values of that column.
        /// </summary>
        public ushort PropertyType;

        /// <summary>
        /// Identifies the column to sort on.
        /// </summary>
        public ushort PropertyId;

        /// <summary>
        /// MUST be one of the values listed in [MS-OXCDATA] 2.14.1.
        /// </summary>
        public byte Order;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializeBuffer = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes(this.PropertyType), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            Array.Copy(BitConverter.GetBytes(this.PropertyId), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            serializeBuffer[index++] = this.Order;
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 5;
            return size;
        }
    }

    /// <summary>
    /// RopSortTable response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSortTableResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x13.
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