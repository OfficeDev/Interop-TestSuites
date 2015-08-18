namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopTransportSend request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopTransportSendRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x4A.
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
    /// RopTransportSend response buffer structure.
    /// </summary>
    public struct RopTransportSendResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x4A.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether property values are returned.
        /// </summary>
        public byte NoPropertiesReturned;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the number of structures returned in the PropertyValues field.
        /// </summary>
        public ushort PropertyValueCount;

        /// <summary>
        /// Array of TaggedPropertyValue structures. The format of the PropertyValue structure is specified in [MS-OXCDATA]. 
        /// The number of structures contained in this field is specified by the PropertyValueCount field. 
        /// This array specifies the properties to copy.
        /// </summary>
        public TaggedPropertyValue[] PropertyValues;

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
                this.NoPropertiesReturned = ropBytes[index++];
                this.PropertyValueCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.PropertyValueCount > 0)
                {
                    this.PropertyValues = new TaggedPropertyValue[this.PropertyValueCount];

                    Context.Instance.PropertyBytes = ropBytes;
                    Context.Instance.CurIndex = index;
                    Context.Instance.CurProperty = new Property(PropertyType.PtypUnspecified);
                    for (int i = 0; i < this.PropertyValueCount; i++)
                    {
                        this.PropertyValues[i] = new TaggedPropertyValue();
                        this.PropertyValues[i].Parse(Context.Instance);
                        index = Context.Instance.CurIndex;
                    }
                }
            }

            return index - startIndex;
        }
    }
}