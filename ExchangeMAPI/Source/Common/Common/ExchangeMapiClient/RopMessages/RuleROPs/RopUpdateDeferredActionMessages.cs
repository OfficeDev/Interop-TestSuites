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
    /// RopUpdateDefferredActionMessages request buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopUpdateDeferredActionMessagesRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x57.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon on which the operation is performed.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Unsigned 8-bit integer. This index refers to the location in the Server Object Handle Table used to find the handle for this operation.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the ServerEntryId field.
        /// </summary>
        public ushort ServerEntryIdSize;

        /// <summary>
        /// Byte Array. The size of this field, in bytes, is specified by the ServerEntryIdSize field. This value specifies the ID of the message on the server.
        /// </summary>
        public byte[] ServerEntryId;

        /// <summary>
        /// Unsigned 16-bit integer. This value specifies the size of the ClientEntryId field.
        /// </summary>
        public ushort ClientEntryIdSize;

        /// <summary>
        /// Byte Array. The size of this field, in bytes, is specified by the ClientEntryIdSize field. This value specifies the ID of the downloaded message on the client.
        /// </summary>
        public byte[] ClientEntryId;

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

            Array.Copy(BitConverter.GetBytes((ushort)this.ServerEntryIdSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.ServerEntryIdSize > 0)
            {
                Array.Copy(this.ServerEntryId, 0, serializeBuffer, index, this.ServerEntryIdSize);
                index += this.ServerEntryIdSize;
            }

            Array.Copy(BitConverter.GetBytes((ushort)this.ClientEntryIdSize), 0, serializeBuffer, index, sizeof(ushort));
            index += sizeof(ushort);
            if (this.ClientEntryIdSize > 0)
            {
                Array.Copy(this.ClientEntryId, 0, serializeBuffer, index, this.ClientEntryIdSize);
                index += this.ClientEntryIdSize;
            }

            return serializeBuffer;
        }

        /// <summary>
        /// Return the size of this structure.
        /// </summary>
        /// <returns>The size of this structure.</returns>
        public int Size()
        {
            // 7 indicates sizeof(byte) * 3 + sizeof(UInt16) * 2
            int size = sizeof(byte) * 7;

            if (this.ServerEntryIdSize > 0)
            {
                size += this.ServerEntryIdSize;
            }

            if (this.ClientEntryIdSize > 0)
            {
                size += this.ClientEntryIdSize;
            }

            return size;
        }
    }

    /// <summary>
    /// RopUpdateDeferredActionMessages response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopUpdateDeferredActionMessagesResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. 
        /// For this operation, this field is set to 0x57.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index refers to the handle in the Server Object Handle 
        /// Table specified as the input handle.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation.
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
                this = (RopUpdateDeferredActionMessagesResponse)Marshal.PtrToStructure(
                    responseBuffer, 
                    typeof(RopUpdateDeferredActionMessagesResponse));
                return Marshal.SizeOf(this);
            }
            finally
            {
                Marshal.FreeHGlobal(responseBuffer);
            }
        }
    }
}