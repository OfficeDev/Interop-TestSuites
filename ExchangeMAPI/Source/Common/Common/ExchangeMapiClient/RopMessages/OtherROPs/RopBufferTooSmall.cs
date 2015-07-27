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
    /// RopBufferTooSmall response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopBufferTooSmallResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the type of remote operation. For this operation, this field is set to 0xFF.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This value specifies the size requested for the output buffer for a subsequent request.
        /// </summary>
        public ushort SizeNeeded;

        /// <summary>
        /// Array of bytes. The size of this array is the remaining space in the output ROP response buffer. 
        /// This array specifies the section of the input buffer that was not executed because of this condition.
        /// </summary>
        public byte[] RequestBuffers;

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

            this.SizeNeeded = (ushort)BitConverter.ToInt16(ropBytes, index);
            index += sizeof(ushort);
            this.RequestBuffers = new byte[ropBytes.Length - index];
            Array.Copy(ropBytes, index, this.RequestBuffers, 0, ropBytes.Length - index);
            index += ropBytes.Length - index;
           
            return index - startIndex;
        }
    }
}