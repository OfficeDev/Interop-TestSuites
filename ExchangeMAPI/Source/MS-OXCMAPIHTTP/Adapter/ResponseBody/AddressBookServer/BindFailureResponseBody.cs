
namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    ///  A class indicates the Bind request type failure response body.
    /// </summary>
    public class BindFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the Bind request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The response body of bind request.</returns>
        public static BindFailureResponseBody Parse(byte[] rawData)
        {
            BindFailureResponseBody responseBody = new BindFailureResponseBody();
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
