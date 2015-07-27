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
    /// RopSetMessageReadFlag request buffer structure
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetMessageReadFlagRequest : ISerializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x11.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the logon associated with this operation.
        /// </summary>
        public byte LogonId;
        
        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table that is referenced in the response.
        /// </summary>
        public byte ResponseHandleIndex;

        /// <summary>
        /// Unsigned 8-bit integer. This index specifies the location in the Server Object Handle Table 
        /// where the handle for the input Server Object is stored.
        /// </summary>
        public byte InputHandleIndex;

        /// <summary>
        /// 8-bit flags structure. The possible values for these flags are specified in [MS-OXCMSG].
        /// </summary>
        public byte ReadFlags;

        /// <summary>
        /// Array of 24 bytes. This field is present when the logon associated with LogonId was created with the Private flag 
        /// unset (see [MS-OXCSTOR] for more information) and is not present otherwise. 
        /// This value specifies the information that is returned to the client in a successful response.
        /// </summary>
        public byte[] ClientData; 

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
            serializedBuffer[index++] = this.ResponseHandleIndex;
            serializedBuffer[index++] = this.InputHandleIndex;
            serializedBuffer[index++] = this.ReadFlags;
            if (null != this.ClientData)
            {
                Array.Copy(this.ClientData, 0, serializedBuffer, index, this.ClientData.Length);
                index += this.ClientData.Length;
            }

            return serializedBuffer;
        }

        /// <summary>
        /// Return the size of RopSetMessageReadFlag request buffer structure
        /// </summary>
        /// <returns>The size of RopSetMessageReadFlag request buffer structure.</returns>
        public int Size()
        {
            // 5 indicates sizeof (byte) * 5
            int size = sizeof(byte) * 5;

            if (null != this.ClientData)
            {
                size += this.ClientData.Length;
            }
            
            return size;
        }
    }

    /// <summary>
    /// RopSetMessageReadFlag response buffer structure
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopSetMessageReadFlagResponse : IDeserializable
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies the type of remote operation. For this operation, this field is set to 0x11.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// Unsigned 8-bit integer. This index MUST be set to the ResponseHandleIndex specified in the request.
        /// </summary>
        public byte ResponseHandleIndex;

        /// <summary>
        /// Unsigned 32-bit integer. This value specifies the status of the remote operation. 
        /// For successful response, this field is set to 0x00000000. For failure response, this field is set to a value other than 0x00000000.
        /// </summary>
        public uint ReturnValue;

        /// <summary>
        /// 8-bit Boolean. This value specifies whether the read status of a public folder's message has changed.
        /// </summary>
        public byte ReadStatusChanged;

        /// <summary>
        /// Unsigned 8-bit integer. This field is present when the value in the ReadStatusChanged field is non-zero and is
        /// not present otherwise. This field MUST be set to the value of the LogonId field in the request.
        /// </summary>
        public byte? LogonId;

        /// <summary>
        /// Array of 24 bytes. This field is present when the value in the ReadStatusChanged field is non-zero and is not 
        /// present otherwise. This field MUST be set to the value of the ClientData field in the request.
        /// </summary>
        public byte[] ClientData;

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
            this.ResponseHandleIndex = ropBytes[index++];
            this.ReturnValue = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;

            // Only success response has below fields
            if (this.ReturnValue == 0)
            {
                this.ReadStatusChanged = ropBytes[index++];
                if (this.ReadStatusChanged != 0)
                {
                    this.LogonId = ropBytes[index++];
                    this.ClientData = new byte[24];
                    Array.Copy(ropBytes, index, this.ClientData, 0, 24);
                    index += 24;
                }
            }

            return index - startIndex;
        }
    }
}