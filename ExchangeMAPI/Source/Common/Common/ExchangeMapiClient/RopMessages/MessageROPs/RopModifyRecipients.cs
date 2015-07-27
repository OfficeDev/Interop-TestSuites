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
    /// ModifyRecipientRow request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct ModifyRecipientRow : ISerializable
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
        /// This value specifies the size of the RecipientRow field.
        /// </summary>
        public ushort RecipientRowSize;

        /// <summary>
        /// RecipientRow structure. This field is present when the RecipientRowSize field is non-zero and is not present otherwise. 
        /// The format of the RecipientRow structure is specified in [MS-OXCDATA]. 
        /// The size of this field, in bytes, is specified by the RecipientRowSize field.
        /// </summary>
        public byte[] RecptRow;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];
            Array.Copy(BitConverter.GetBytes((uint)this.RowId), 0, serializedBuffer, index, sizeof(uint));
            index += 4;

            serializedBuffer[index++] = this.RecipientType;
            Array.Copy(BitConverter.GetBytes((ushort)this.RecipientRowSize), 0, serializedBuffer, index, sizeof(ushort));
            index += 2;

            Array.Copy(this.RecptRow, 0, serializedBuffer, index, this.RecipientRowSize);
            index += this.RecipientRowSize;
            return serializedBuffer;   
        }

        /// <summary>
        /// Return the size of ModifyRecipientRow request buffer structure.
        /// </summary>
        /// <returns>The size of ModifyRecipientRow request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (UInt16) + sizeof (UInt32) + sizeof (byte)
            int size = sizeof(byte) * 7;
            if (this.RecipientRowSize > 0)
            {
                size += this.RecipientRowSize;
            }

            return size;
        }
    }

    /// <summary>
    /// RopModifyRecipients request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopModifyRecipientsRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x0E.
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
        /// This value specifies the number of structures in the RecipientColumns field.
        /// </summary>
        public ushort ColumnCount;

        /// <summary>
        /// Array of PropertyTag structures. The number of structures contained in this field is specified by the ColumnCount field. 
        /// The format of the PropertyTag structure is specified in [MS-OXCDATA]. This field specifies the property values that can be included for each recipient row.
        /// </summary>
        public PropertyTag[] RecipientColumns;

        /// <summary>
        /// This value specifies the number of rows in the RecipientRows field.
        /// </summary>
        public ushort RowCount;

        /// <summary>
        /// List of ModifyRecipientRow structures. The number of structures contained in this field is specified by the RowCount field.
        /// </summary>
        public ModifyRecipientRow[] RecipientRows;

        /// <summary>
        /// Serialize the ROP request buffer.
        /// </summary>
        /// <returns>The ROP request buffer serialized.</returns>
        public byte[] Serialize()
        {
            int index = 0;
            byte[] serializedBuffer = new byte[this.Size()];

            serializedBuffer[index++] = this.RopId;
            serializedBuffer[index++] = this.LogonId;
            serializedBuffer[index++] = this.InputHandleIndex;

            Array.Copy(BitConverter.GetBytes((short)this.ColumnCount), 0, serializedBuffer, index, sizeof(ushort));
            index += 2;

            if (this.ColumnCount > 0)
            {
                IntPtr requestBuffer = new IntPtr();
                requestBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(PropertyTag)));
                try
                {
                    Context.Instance.Init();
                    foreach (PropertyTag propTag in this.RecipientColumns)
                    {
                        Marshal.StructureToPtr(propTag, requestBuffer, true);
                        Marshal.Copy(requestBuffer, serializedBuffer, index, Marshal.SizeOf(typeof(PropertyTag)));
                        index += Marshal.SizeOf(typeof(PropertyTag));

                        // Insert properties into Context
                        Context.Instance.Properties.Add(new Property((PropertyType)propTag.PropertyType));
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(requestBuffer);
                }
            }

            Array.Copy(BitConverter.GetBytes((short)this.RowCount), 0, serializedBuffer, index, sizeof(ushort));
            index += 2;
            for (int i = 0; i < this.RowCount; i++)
            {
                Array.Copy(this.RecipientRows[i].Serialize(), 0, serializedBuffer, index, this.RecipientRows[i].Size());
                index += this.RecipientRows[i].Size();
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopModifyRecipients request buffer structure.
        /// </summary>
        /// <returns>The size of RopModifyRecipients request buffer structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof (byte) * 3 + sizeof (UInt16)*2 
            int size = sizeof(byte) * 7;
            if (this.ColumnCount > 0)
            {
                size += this.RecipientColumns.Length * Marshal.SizeOf(typeof(PropertyTag));
            }

            if (this.RowCount > 0)
            {
                for (int i = 0; i < this.RowCount; i++)
                {
                    size += this.RecipientRows[i].Size();
                }
            }

            return size;
        }
    }

    /// <summary>
    /// RopModifyRecipients response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopModifyRecipientsResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x0E.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request. 
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer structure.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            IntPtr responseBuffer = new IntPtr();
            responseBuffer = Marshal.AllocHGlobal(Marshal.SizeOf(this));
            try
            {
                Marshal.Copy(ropBytes, startIndex, responseBuffer, Marshal.SizeOf(this));
                this = (RopModifyRecipientsResponse)Marshal.PtrToStructure(responseBuffer, typeof(RopModifyRecipientsResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}