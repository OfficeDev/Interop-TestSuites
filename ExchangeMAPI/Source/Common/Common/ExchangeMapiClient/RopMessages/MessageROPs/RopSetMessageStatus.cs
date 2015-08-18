namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopSetMessageStatus request buffer structure
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetMessageStatusRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x20.
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
        /// 64-bit identifier. This value identifies the message for which the status will be changed.
        /// </summary>
        public ulong MessageId;

        /// <summary>
        /// 32-bit flags structure. This field contains status flags to set on the message. The possible values are specified in [MS-OXCMSG].
        /// </summary>
        public uint MessageStatusFlags;

        /// <summary>
        /// 32-bit bitmask. This bitmask specifies which bits in the MessageStatusFlags field are to be changed.
        /// </summary>
        public uint MessageStatusMask;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            byte[] serializedBuffer = new byte[Marshal.SizeOf(this)];
            IntPtr requestBuffer = new IntPtr();
            requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.StructureToPtr(this, requestBuffer, true);
                Marshal.Copy(requestBuffer, serializedBuffer, 0, Marshal.SizeOf(this));
                return serializedBuffer;
            }
            finally
            {
                Marshal.FreeHGlobal(requestBuffer);
            }
        }

        /// <summary>
        /// Return the size of RopSetMessageStatus request buffer structure.
        /// </summary>
        /// <returns>The size of RopSetMessageStatus request buffer structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// RopSetMessageStatus response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetMessageStatusResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x20.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For successful response, this field is set to 0x00000000. For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;
        
        /// <summary>
        /// 32-bit flags structure. This field contains the status flags that were set on the message before this operation. 
        /// The possible values are specified in [MS-OXCMSG].
        /// </summary>
        public uint MessageStatusFlags;

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
            index += 4;

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.MessageStatusFlags = (uint)BitConverter.ToInt32(ropBytes, index);
                index += 4;
            }

            return index - startIndex;
        }
    }
}