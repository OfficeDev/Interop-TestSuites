namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;

    /// <summary>
    /// A class indicates the response body of unbind request
    /// </summary>
    public class UnbindResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of unbind request</returns>
        public static UnbindResponseBody Parse(byte[] rawData)
        {
            UnbindResponseBody responseBody = new UnbindResponseBody();
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