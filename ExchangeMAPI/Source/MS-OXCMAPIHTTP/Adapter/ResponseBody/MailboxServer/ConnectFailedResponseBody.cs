namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    
    /// <summary>
    /// A class indicates the Connect request type failed response body.
    /// </summary>
    public class ConnectFailedResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Parse the Connect request type Failed response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of ConnectFailedResponseBody class.</returns>
        public static ConnectFailedResponseBody Parse(byte[] rawData)
        {
            ConnectFailedResponseBody responseBody = new ConnectFailedResponseBody();
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