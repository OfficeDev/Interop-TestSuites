//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopFindRow request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFindRowRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x4F.
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
        public byte FindRowFlags;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the length of the RestrictionData field.
        /// </summary>
        public ushort RestrictionDataSize;

        /// <summary>
        /// Restriction packet. The size of this field, in bytes, is specified by the RestrictionDataSize field. 
        /// This field contains a restriction packet, as specified in [MS-OXCDATA] section 2.13. 
        /// The restriction specifies the filter for this operation.
        /// </summary>
        public byte[] RestrictionData;

        /// <summary>
        /// 8-bit enumeration. The possible values for this enumeration are specified in [MS-OXCTABL]. 
        /// This enumeration specifies where this operation begins its search.
        /// </summary>
        public byte Origin;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the Bookmark field.
        /// </summary>
        public ushort BookmarkSize;

        /// <summary>
        /// Array of bytes. The size of this field, in bytes, is specified by the BookmarkSize field. 
        /// This array specifies the bookmark to use as the origin.
        /// </summary>
        public byte[] Bookmark;

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
            serializeBuffer[index++] = this.FindRowFlags;

            Array.Copy(BitConverter.GetBytes((short)this.RestrictionDataSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);

            if (this.RestrictionData != null)
            {
                Array.Copy(this.RestrictionData, 0, serializeBuffer, index, this.RestrictionData.Length);
                index += this.RestrictionData.Length;
            }

            serializeBuffer[index++] = this.Origin;

            Array.Copy(BitConverter.GetBytes((short)this.BookmarkSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.Bookmark != null)
            {
                Array.Copy(this.Bookmark, 0, serializeBuffer, index, this.Bookmark.Length);
                index += this.Bookmark.Length;
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = (sizeof(byte) * 5) + (sizeof(ushort) * 2);
            if (this.RestrictionData != null)
            {
                size += this.RestrictionData.Length;
            }

            if (this.Bookmark != null)
            {
                size += this.Bookmark.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopFindRow response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopFindRowResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x4F.
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
        /// 8-bit Boolean. This value indicates whether the RowData field is present.
        /// </summary>
        public byte HasRowData;

        /// <summary>
        /// PropertyRow structure. This field is only present when the HasRowData field is set to a non-zero value. 
        /// The format of the PropertyRow structure is specified in [MS-OXCDATA] and the columns used for these 
        /// rows were those previously set on this table by a RopSetColumns.
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
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += sizeof(uint);

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.RowNoLongerVisible = ropBytes[index++];
                this.HasRowData = ropBytes[index++];
                if (this.HasRowData != 0)
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
            }

            return index - startIndex;
        }
    }
}