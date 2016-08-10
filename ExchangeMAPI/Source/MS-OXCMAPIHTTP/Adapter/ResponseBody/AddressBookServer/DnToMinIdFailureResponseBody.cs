namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the DnToMinId request type failure response body.
    /// </summary>
    public class DnToMinIdFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the DoToMinId request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of DnToMinId request</returns>
        public static DnToMinIdFailureResponseBody Parse(byte[] rawData)
        {
            DnToMinIdFailureResponseBody responseBody = new DnToMinIdFailureResponseBody();
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
