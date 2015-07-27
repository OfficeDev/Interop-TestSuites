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
    /// RopOptionsData request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOptionsDataRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x6F.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server object handle table 
        /// where the handle for the input Server object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Null-terminated ASCII string. This value specifies the address type for which options are to be returned.
        /// </summary>
        public byte[] AddressType;

        /// <summary>
        /// Unsigned 8-bit Boolean. This value specifies whether the help file data is to be returned in a format that is suited for 32-bit machines.
        /// </summary>
        public byte WantWin32;

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

            if (this.AddressType != null)
            {
                Array.Copy(this.AddressType, 0, serializeBuffer, index, this.AddressType.Length);
                index += this.AddressType.Length;
            }

            serializeBuffer[index++] = this.WantWin32;         
            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            int size = sizeof(byte) * 4;
            if (this.AddressType != null)
            {
                size += this.AddressType.Length;
            }

            return size;
        }
    }

    /// <summary>
    /// RopOptionsData response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopOptionsDataResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x6F.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This field MUST be set to the InputHandleIndex specified in the request.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// Reserved. This field MUST be set to 0x01.
        /// </summary>
        public byte Reserved;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the OptionsInfo field.
        /// </summary>
        public ushort OptionsInfoSize;

        /// <summary>
        /// Array of bytes. This field contains opaque data from the server. Clients SHOULD ignore this field. 
        /// Servers SHOULD return this field as an empty array.
        /// </summary>
        public byte[] OptionsInfo;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the HelpFile field.
        /// </summary>
        public ushort HelpFileSize;

        /// <summary>
        /// Array of bytes. This field contains the help file associated with the specified address type.
        /// </summary>
        public byte[] HelpFile;

        /// <summary>
        /// Null-terminated ASCII string. This string specifies the name of the help file 
        /// that is associated with the specified address type. This field is present if HelpFileSize is 
        /// non-zero and is not present otherwise.
        /// </summary>
        public byte[] HelpFileName;

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
                this.Reserved = ropBytes[index++];
                this.OptionsInfoSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.OptionsInfoSize > 0)
                {
                    this.OptionsInfo = new byte[this.OptionsInfoSize];
                    Array.Copy(ropBytes, index, this.OptionsInfo, 0, this.OptionsInfoSize);
                    index += this.OptionsInfoSize;
                }

                this.HelpFileSize = (ushort)BitConverter.ToInt16(ropBytes, index);
                index += sizeof(ushort);
                if (this.HelpFileSize >= 0)
                {
                    this.HelpFile = new byte[this.HelpFileSize];
                    Array.Copy(ropBytes, index, this.HelpFile, 0, this.HelpFileSize);
                    index += this.HelpFileSize;

                    if (this.HelpFileSize > 0)
                    {
                        this.HelpFileName = new byte[ropBytes.Length - index];
                        Array.Copy(ropBytes, index, this.HelpFileName, 0, ropBytes.Length - index);
                        index += ropBytes.Length - index - 1;
                    }
                }
            }

            return index - startIndex;
        }
    }
}