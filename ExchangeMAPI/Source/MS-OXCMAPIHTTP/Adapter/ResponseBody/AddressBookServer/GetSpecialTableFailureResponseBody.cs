
namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the GetSpecialTable request type failure response body.
    /// </summary>
    public class GetSpecialTableFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the GetSpecialTable request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of GetSpecialTable request.</returns>
        public static GetSpecialTableFailureResponseBody Parse(byte[] rawData)
        {
            GetSpecialTableFailureResponseBody responseBody = new GetSpecialTableFailureResponseBody();
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
