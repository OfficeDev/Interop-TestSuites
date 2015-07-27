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
    using System.Collections.Generic;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopGetReceiveFolderTable request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetReceiveFolderTableRequest : ISerializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x68.
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
    /// RopGetReceiveFolderTable response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopGetReceiveFolderTableResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0x68.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This index MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// This value specifies the status of the remote operation. For this response, this field is set to 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// This value specifies the number of row structures contained in the Rows field.
        /// </summary>
        public uint RowCount;

        /// <summary>
        /// Array of row structures. This field contains the rows of the Receive folder table. 
        /// Each row is returned in either a StandardPropertyRow structure or a FlaggedPropertyRow structure, 
        /// both of which are specified in [MS-OXCDATA] sections 2.9.1.1 and 2.9.1.2. 
        /// The number of row structures contained in this field is specified by the RowCount field. 
        /// The ValueArray field of either StandardPropertyRow or FlaggedPropertyRow MUST include only the PidTagFolderId, 
        /// PidTagMessageClass, and PidTagLastModificationTime properties, in that order, and no other properties.
        /// </summary>
        public PropertyRowSet Rows;

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
                this.RowCount = (uint)BitConverter.ToUInt32(ropBytes, index);
                index += sizeof(uint);
             
                List<Property> properties = new List<Property>();
                Property pidTagFolderId = new Property(PropertyType.PtypInteger64)
                {
                    Name = "PidTagFolderId"
                };
                properties.Add(pidTagFolderId);

                Property pidTagMessageClass = new Property(PropertyType.PtypString8)
                {
                    Name = "PidTagMessageClass"
                };
                properties.Add(pidTagMessageClass);

                Property pidTagLastModificationTime = new Property(PropertyType.PtypTime)
                {
                    Name = "PidTagLastModificationTime"
                };
                properties.Add(pidTagLastModificationTime);

                Context.Instance.Properties = properties;
                Context.Instance.PropertyBytes = ropBytes;
                Context.Instance.CurIndex = index;
                this.Rows = new PropertyRowSet
                {
                    Count = (int)this.RowCount
                };

                // Set row count
                this.Rows.Parse(Context.Instance);

                // Context.Instance.CurIndex indicates the already deserialized bytes' index
                index = Context.Instance.CurIndex;
            }

            return index - startIndex;
        }
    } 
}