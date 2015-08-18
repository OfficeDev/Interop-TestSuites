namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the Disconnect request type failure response body.
    /// </summary>
    public class DisconnectFailedResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Parse the Disconnect request type Failed response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of DisconnectFailedResponseBody class.</returns>
        public static DisconnectFailedResponseBody Parse(byte[] rawData)
        {
            DisconnectFailedResponseBody responseBody = new DisconnectFailedResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);

            return responseBody;
        }
    }
}