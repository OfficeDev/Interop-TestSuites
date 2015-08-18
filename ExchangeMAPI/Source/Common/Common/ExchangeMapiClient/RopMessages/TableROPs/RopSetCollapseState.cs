namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopSetCollapseState request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetCollapseStateRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x6C.
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
        /// Unsigned 16-bit integer. This value specifies the size of the CollapseState field.
        /// </summary>
        public ushort CollapseStateSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the CollapseStateSize field. 
        /// This array specifies a collapse state for a categorized table.
        /// </summary>
        public byte[] CollapseState;

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

            Array.Copy(BitConverter.GetBytes((short)this.CollapseStateSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            if (this.CollapseState != null)
            {
                Array.Copy(this.CollapseState, 0, serializeBuffer, index, this.CollapseState.Length);
                index += this.CollapseState.Length;
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
            if (this.CollapseState != null)
            {
                size += this.CollapseState.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopSetCollapseState response buffer structure
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetCollapseStateResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x6C.
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
        /// Unsigned 16-bit integer. This value specifies the size of the Bookmark field.
        /// </summary>
        public ushort BookmarkSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the BookmarkSize field. 
        /// This array specifies the current cursor position.
        /// </summary>
        public byte[] Bookmark;

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
                this.BookmarkSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                this.Bookmark = new byte[this.BookmarkSize];
                Array.Copy(ropBytes, index, this.Bookmark, 0, this.BookmarkSize);
                index += this.BookmarkSize;
            }

            return index - startIndex;
        }
    }
}