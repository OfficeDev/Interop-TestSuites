namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the GetAddressBookUrl request type failure response body.
    /// </summary>
    public class GetAddressBookUrlFailureResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the GetAddressBookUrl request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of GetAddressBookUrl request.</returns>
        public static GetAddressBookUrlFailureResponseBody Parse(byte[] rawData)
        {
            GetAddressBookUrlFailureResponseBody responseBody = new GetAddressBookUrlFailureResponseBody();
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
