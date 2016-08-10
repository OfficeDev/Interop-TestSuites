namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;

    /// <summary>
    /// A class indicates the response body of GetPropList request 
    /// </summary>
    public class GetPropListResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the PropertyTags field is present.
        /// </summary>
        public bool HasPropertyTags { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that contains the property tags of properties that have values on the requested object.
        /// </summary>
        public LargePropertyTagArray? PropertyTags { get; set; }

        /// <summary>
        /// Parse the GetPropList request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The GetPropList request type response body.</returns>
        public static GetPropListResponseBody Parse(byte[] rawData)
        {
            GetPropListResponseBody responseBody = new GetPropListResponseBody();
            int index = 0;

            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasPropertyTags = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasPropertyTags)
            {
                responseBody.PropertyTags = LargePropertyTagArray.Parse(rawData, ref index);
            }
            else
            {
                responseBody.PropertyTags = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}