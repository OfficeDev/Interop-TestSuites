namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the ModProps request type failure response body.
    /// </summary>
    public class ModPropsFailureResponseBody: AddressBookResponseBodyBase
    {
        /// <summary>
        /// Parse the ModProps request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The failure response body of ModProps request.</returns>
        public static ModPropsFailureResponseBody Parse(byte[] rawData)
        {
            ModPropsFailureResponseBody responseBody = new ModPropsFailureResponseBody();
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
