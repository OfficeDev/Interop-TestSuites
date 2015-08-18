namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Text;

    /// <summary>
    /// A successful response body for the connect request for Mailbox Server Endpoint.
    /// </summary>
    public class ConnectSuccessResponseBody : MailboxResponseBodyBase
    {
        /// <summary>
        /// Gets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; private set; }

        /// <summary>
        /// Gets an unsigned integer that specifies the number of milliseconds for the maximum polling interval.
        /// </summary>
        public uint PollsMax { get; private set; }

        /// <summary>
        /// Gets an unsigned integer that specifies the number of times to retry request types.
        /// </summary>
        public uint RetryCount { get; private set; }

        /// <summary>
        /// Gets an unsigned integer that specifies the number of milliseconds for the client to wait before retrying a failed request type.
        /// </summary>
        public uint RetryDelay { get; private set; }

        /// <summary>
        /// Gets a null-terminated ASCII string that specifies the DN prefix to be used for building message recipients.
        /// </summary>
        public string DNPrefix { get; private set; }

        /// <summary>
        /// Gets a null-terminated Unicode string that specifies the display name of the user who is specified in the UserDn field of the Connect request type request body.
        /// </summary>
        public string DisplayName { get; private set; }

        /// <summary>
        /// Parse the Connect request type success response body.
        /// </summary>
        /// <param name="rawData">The raw data which is returned by server.</param>
        /// <returns>An instance of ConnectSuccessResponseBody class.</returns>
        public static ConnectSuccessResponseBody Parse(byte[] rawData)
        {
            ConnectSuccessResponseBody responseBody = new ConnectSuccessResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.PollsMax = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.RetryCount = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.RetryDelay = BitConverter.ToUInt32(rawData, index);
            index += 4;

            // The length in bytes of the unicode string to parse
            int strBytesLen = 0;

            // Find the string with '\0' end
            for (int i = index; i < rawData.Length; i++)
            {
                strBytesLen++;
                if (rawData[i] == 0)
                {
                    break;
                }
            }

            byte[] prefixBuffer = new byte[strBytesLen];
            Array.Copy(rawData, index, prefixBuffer, 0, strBytesLen);
            index += strBytesLen;
            responseBody.DNPrefix = Encoding.ASCII.GetString(prefixBuffer);

            // The length in bytes of the unicode string to parse
            strBytesLen = 0;

            // Find the string with '\0''\0' end
            for (int i = index; i < rawData.Length; i += 2)
            {
                strBytesLen += 2;
                if ((rawData[i] == 0) && (rawData[i + 1] == 0))
                {
                    break;
                }
            }

            byte[] displayNameBuffer = new byte[strBytesLen];
            Array.Copy(rawData, index, displayNameBuffer, 0, strBytesLen);
            index += strBytesLen;
            responseBody.DisplayName = Encoding.Unicode.GetString(displayNameBuffer);
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}