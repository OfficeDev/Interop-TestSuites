namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopOpenEmbeddedMessage request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenEmbeddedMessageRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x46.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the input Server Object is stored. 
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This index specifies the location in the Server Object Handle Table where the handle for the output Server Object will be stored. 
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies which code page is used for string values associated with the message.
        /// </summary>
        public ushort CodePageId;

        /// <summary>
        /// The possible values are specified in [MS-OXCMSG]. These flags control the access to the message.
        /// </summary>
        public byte OpenModeFlags;

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
        /// Return the size of RopOpenEmbeddedMessage request buffer structure.
        /// </summary>
        /// <returns>The size of RopOpenEmbeddedMessage request buffer structure.</returns>
        public int Size()
        {
            return Marshal.SizeOf(this);
        }
    }

    /// <summary>
    /// RopOpenEmbeddedMessage response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOpenEmbeddedMessageResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x46.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the OutputHandleIndex specified in the request. 
        /// </summary>
        public byte OutputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For successful response, this field is set to 0x00000000.
        /// For failure response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This field MUST be set to 0x00.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// This value specifies the ID of the embedded message.
        /// </summary>
        public ulong MessageId;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the message has named properties.
        /// </summary>
        public byte HasNamedProperties;

        /// <summary>
        /// TypedString structure. The format of the TypedString structure is specified in [MS-OXCDATA]. 
        /// This structure specifies the subject prefix of the message.
        /// </summary>
        public TypedString SubjectPrefix;

        /// <summary>
        /// TypedString structure. The format of the TypedString structure is specified in [MS-OXCDATA]. 
        /// This structure specifies the normalized subject of the message.
        /// </summary>
        public TypedString NormalizedSubject;

        /// <summary>
        /// This value specifies the number of recipients on the message.
        /// </summary>
        public ushort RecipientCount;

        /// <summary>
        /// This value specifies the number of structures in the RecipientColumns field.
        /// </summary>
        public ushort ColumnCount;

        /// <summary>
        /// Array of Property Tag structures. The number of structures contained in this field is specified by the ColumnCount field. 
        /// The format of the Property Tag structure is specified in [MS-OXCDATA]. 
        /// This field specifies the property values that can be included for each recipient row.
        /// </summary>
        public PropertyTag[] RecipientColumns;

        /// <summary>
        /// This value specifies the number of rows in the RecipientRows field.
        /// </summary>
        public byte RowCount;

        /// <summary>
        /// List of OpenRecipientRow structures. The number of structures contained in this field is specified by the RowCount field. 
        /// </summary>
        public OpenRecipientRow[] RecipientRows;

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
            index += 4;

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.Reserved = ropBytes[index++];
                this.MessageId = (ulong)BitConverter.ToInt64(ropBytes, index);
                index += 8;
                this.HasNamedProperties = ropBytes[index++];

                index += this.SubjectPrefix.Deserialize(ropBytes, index);
                index += this.NormalizedSubject.Deserialize(ropBytes, index);
                this.RecipientCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += 2;
                this.ColumnCount = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += 2;

                // RecipientColumns
                if (this.ColumnCount >= 0)
                {
                    this.RecipientColumns = new PropertyTag[this.ColumnCount];
                    Context.Instance.Init();
                    for (int i = 0; i < this.ColumnCount; i++)
                    {
                        index += this.RecipientColumns[i].Deserialize(ropBytes, index);
                        Context.Instance.Properties.Add(new Property((PropertyType)this.RecipientColumns[i].PropertyType));
                    }

                    this.RowCount = ropBytes[index++];
                    if (this.RowCount >= 0)
                    {
                        this.RecipientRows = new OpenRecipientRow[this.RowCount];
                        for (int i = 0; i < this.RowCount; i++)
                        {
                            index += this.RecipientRows[i].Deserialize(ropBytes, index);
                        }
                    }
                }
            }

            return index - startIndex;
        }
    }
}