namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the GetMailboxUrl request type failure response body.
    /// </summary>
    public class GetMailboxUrlFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the GetMailboxUrl request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of GetMailboxUrl request.</returns>
        public static GetMailboxUrlFailureResponseBody Parse(byte[] rawData)
        {
            GetMailboxUrlFailureResponseBody responseBody = new GetMailboxUrlFailureResponseBody();
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
