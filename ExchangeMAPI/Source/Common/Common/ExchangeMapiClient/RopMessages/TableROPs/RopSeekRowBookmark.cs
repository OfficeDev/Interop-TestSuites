namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopSeekRowBookmark request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSeekRowBookmarkRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x19.
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
        /// Unsigned 16-bit integer. This value specifies the size of the Bookmark field.
        /// </summary>
        public ushort BookmarkSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the BookmarkSize field. 
        /// This array specifies the origin for the seek operation.
        /// </summary>
        public byte[] Bookmark;

        /// <summary>
        /// Signed 32-bit integer. This value specifies the direction and the number of rows to seek.
        /// </summary>
        public int RowCount;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the server returns the actual number of rows sought in the response.
        /// </summary>
        public byte WantRowMovedCount;

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
            Array.Copy(BitConverter.GetBytes((short)this.BookmarkSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.Bookmark != null)
            {
                Array.Copy(this.Bookmark, 0, serializeBuffer, index, this.Bookmark.Length);
                index += this.Bookmark.Length;
            }

            Array.Copy(BitConverter.GetBytes((int)this.RowCount), 0, serializeBuffer, index, sizeof(uint));
            index += sizeof(uint);
            serializeBuffer[index++] = this.WantRowMovedCount;
 
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        { 
            // 10 indicates sizeof(byte) * 4 + sizeof(UInt16) + sizeof(UInt32)
            int size = sizeof(byte) * 10;
            if (this.Bookmark != null)
            {
                size += this.Bookmark.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopSeekRowBookmark response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSeekRowBookmarkResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x19.
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
        /// 8-bit Boolean. This value specifies whether the bookmark target is no longer visible.
        /// </summary>
        public byte RowNoLongerVisible;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the full number of rows sought past was less than 
        /// the number that was requested.
        /// </summary>
        public byte HasSoughtLess;

        /// <summary>
        /// Signed 32-bit integer. This value specifies the direction and number of rows sought.
        /// </summary>
        public uint RowsSought;

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
                this.RowNoLongerVisible = ropBytes[index++];
                this.HasSoughtLess = ropBytes[index++];
                this.RowsSought = (uint)BitConverter.ToInt32(ropBytes, index);
                index += sizeof(uint);
            }

            return index - startIndex;
        }
    }
}