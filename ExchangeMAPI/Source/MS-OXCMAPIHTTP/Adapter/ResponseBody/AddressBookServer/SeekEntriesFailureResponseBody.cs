namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    ///  A class indicates the SeekEntries request type failure response body.
    /// </summary>
    public class SeekEntriesFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the SeekEntries request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of SeekEntries request.</returns>
        public static SeekEntriesFailureResponseBody Parse(byte[] rawData)
        {
            SeekEntriesFailureResponseBody responseBody = new SeekEntriesFailureResponseBody();
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
