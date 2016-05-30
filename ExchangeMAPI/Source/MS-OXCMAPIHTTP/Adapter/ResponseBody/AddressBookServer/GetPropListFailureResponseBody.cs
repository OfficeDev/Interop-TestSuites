namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the GetPropList request type failure response body.
    /// </summary>
    public class GetPropListFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the GetPropList request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of GetPropList request.</returns>
        public static GetPropListFailureResponseBody Parse(byte[] rawData)
        {
            GetPropListFailureResponseBody responseBody = new GetPropListFailureResponseBody();
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
