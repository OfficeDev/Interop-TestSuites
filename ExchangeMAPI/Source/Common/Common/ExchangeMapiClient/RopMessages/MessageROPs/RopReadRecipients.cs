namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopReadRecipients request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopReadRecipientsRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x0F.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// . This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored. 
        /// </summary>
        public byte InputHandleIndex;
 
        /// <summary>
        /// This value specifies the recipient to start reading.
        /// </summary>
        public uint RowId;

        /// <summary>
        /// This field MUST be set to 0x0000. Server behavior is undefined if this field is not set to 0x0000.
        /// </summary>
        public ushort Reserved;

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
        /// Return the size of RopReadRecipients request buffer structure.
        /// </summary>
        /// <returns>The size of RopReadRecipients request buffer structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// ReadRecipientRow structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct ReadRecipientRow : IDeserializable
    {
        /// <summary>
        /// This value specifies the ID of the recipient row.
        /// </summary>
        public uint RowId;

        /// <summary>
        /// 8-bit enumeration. The possible values for this enumeration are specified in [MS-OXCMSG]. 
        /// This enumeration specifies the type of recipient.
        /// </summary>
        public byte RecipientType;

        /// <summary>
        /// This value specifies the code page for the recipient.
        /// </summary>
        public ushort CodePageId;

        /// <summary>
        /// The server MUST set this field to 0x0000.
        /// </summary>
        public ushort Reserved;

        /// <summary>
        /// This value specifies the size of the RecipientRow field.
        /// </summary>
        public ushort RecipientRowSize;

        /// <summary>
        /// RecipientRow structure. The format of this structure is specified in [MS-OXCDATA]. 
        /// The size of this field, in bytes, is specified by the RecipientRowSize field.
        /// </summary>
        public RecipientRow RecipientRow;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
           int index = startIndex;
           
            this.RowId = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;
            this.RecipientType = ropBytes[index++];
            this.CodePageId = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += 2;
            this.Reserved = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += 2;
            this.RecipientRowSize = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += 2;
            index += this.RecipientRow.Deserialize(ropBytes, index);
            return index - startIndex;
        }
    }

    /// <summary>
    /// RopReadRecipients response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopReadRecipientsResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x0F.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This value specifies the number of structures in the RecipientRows field.
        /// </summary>
        public byte RowCount;

        /// <summary>
        /// List of ReadRecipientRow structures. The number of structures contained in this field is specified by the RowCount field. 
        /// </summary>
        public ReadRecipientRow[] RecipientRows;

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
                this.RowCount = ropBytes[index++];
                if (this.RowCount > 0)
                {
                    this.RecipientRows = new ReadRecipientRow[this.RowCount];
                    for (int i = 0; i < this.RowCount; i++)
                    {
                        index += this.RecipientRows[i].Deserialize(ropBytes, index);
                    }
                }
            }

            return index - startIndex;
        }
    }
}