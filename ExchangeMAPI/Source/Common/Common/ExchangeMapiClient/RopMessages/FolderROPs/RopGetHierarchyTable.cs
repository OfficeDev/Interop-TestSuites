namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopGetHierarchyTable request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetHierarchyTableRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x04.
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
        /// The possible values are specified in [MS-OXCFOLD]. These flags control the type of table.
        /// </summary>
        public byte TableFlags;

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
    /// RopGetHierarchyTable response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetHierarchyTableResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x04.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request.
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. 
        /// For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This value represents the number of rows in the hierarchy table.
        /// </summary>
        public uint RowCount;

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
            this.OutputHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);
                                                                                                                                                                             
            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.RowCount = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
            }
                                                                                                                                                                                                
            return index - startIndex;
        }
    }
}