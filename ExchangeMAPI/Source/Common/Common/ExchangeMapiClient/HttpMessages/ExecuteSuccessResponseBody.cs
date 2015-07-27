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

    /// <summary>
    /// A class indicates the Execute request type success response body 
    /// </summary>
    public class ExecuteSuccessResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Gets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; private set; }

        /// <summary>
        /// Gets the Flags. Reserved. The server MUST set this field to 0x00000000 and the client MUST ignore this field.
        /// </summary>
        public uint Flags { get; private set; }

        /// <summary>
        /// Gets an unsigned integer that specifies the size, in bytes, of the RopBuffer field.
        /// </summary>
        public uint RopBufferSize { get; private set; }

        /// <summary>
        /// Gets an array of bytes that constitute the ROP responses payload. 
        /// </summary>
        public byte[] RopBuffer { get; private set; }

        /// <summary>
        /// Parse the Execute request type success response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of ExecuteSuccessResponseBody class.</returns>
        public static ExecuteSuccessResponseBody Parse(byte[] rawData)
        {
            ExecuteSuccessResponseBody responseBody = new ExecuteSuccessResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.Flags = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.RopBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.RopBuffer = new byte[responseBody.RopBufferSize];
            Array.Copy(rawData, index, responseBody.RopBuffer, 0, responseBody.RopBufferSize);
            index += (int)responseBody.RopBufferSize;

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);

            return responseBody;
        }
    }
}
