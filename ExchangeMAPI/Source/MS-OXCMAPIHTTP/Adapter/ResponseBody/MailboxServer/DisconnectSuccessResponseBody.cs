namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the Disconnect request type success response body 
    /// </summary>
    public class DisconnectSuccessResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Gets or private sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; private set; }

        /// <summary>
        /// Parse the Disconnect request type success response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of DisconnectSuccessResponseBody class.</returns>
        public static DisconnectSuccessResponseBody Parse(byte[] rawData)
        {
            DisconnectSuccessResponseBody responseBody = new DisconnectSuccessResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);

            return responseBody;
        }
    }
}