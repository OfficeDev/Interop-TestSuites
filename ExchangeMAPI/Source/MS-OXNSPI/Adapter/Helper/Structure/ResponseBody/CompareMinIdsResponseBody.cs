namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;

    /// <summary>
    /// A class indicates the response body of CompareMinIds request 
    /// </summary>
    public class CompareMinIdsResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a signed integer that specifies the result of the comparison.
        /// </summary>
        public int Result { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of compare minimal IDs request</returns>
        public static CompareMinIdsResponseBody Parse(byte[] rawData)
        {
            CompareMinIdsResponseBody responseBody = new CompareMinIdsResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.Result = BitConverter.ToInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}