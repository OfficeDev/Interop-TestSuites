namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;

    /// <summary>
    /// A class indicates the response body of bind request 
    /// </summary>
    public class BindResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a GUID that is associated with a specific address book server.
        /// </summary>
        public Guid ServerGuid { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of bind request</returns>
        public static BindResponseBody Parse(byte[] rawData)
        {
            BindResponseBody responseBody = new BindResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;

            byte[] serverGuidBytes = new byte[16];
            Array.Copy(rawData, index, serverGuidBytes, 0, 16);
            responseBody.ServerGuid = new Guid(serverGuidBytes);
            index += 16;
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}