namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    
    /// <summary>
    /// A class indicates the NotificationWait request type success response body.
    /// </summary>
    public class NotificationWaitSuccessResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Gets or private set an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; private set; }

        /// <summary>
        /// Gets or private set an unsigned integer that indicates whether an event is pending.
        /// </summary>
        public uint EventPending { get; private set; }

        /// <summary>
        /// Parse the NotificationWait request type success response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of NotificationWaitSuccessResponseBody class.</returns>
        public static NotificationWaitSuccessResponseBody Parse(byte[] rawData)
        {
            NotificationWaitSuccessResponseBody responseBody = new NotificationWaitSuccessResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.EventPending = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);

            return responseBody;
        }
    }
}