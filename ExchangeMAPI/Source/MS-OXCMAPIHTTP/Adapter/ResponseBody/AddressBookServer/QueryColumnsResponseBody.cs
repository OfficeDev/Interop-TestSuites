namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the response body of QueryColumns request 
    /// </summary>
    public class QueryColumnsResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Columns field is present.
        /// </summary>
        public bool HasColumns { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the properties that exist on the address book.
        /// </summary>
        public LargePropertyTagArray? Columns { get; set; }

        /// <summary>
        /// Parse the QueryColumns request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The QueryColumns request type response body.</returns>
        public static QueryColumnsResponseBody Parse(byte[] rawData)
        {
            QueryColumnsResponseBody responseBody = new QueryColumnsResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasColumns = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasColumns)
            {
                responseBody.Columns = LargePropertyTagArray.Parse(rawData, ref index);
            }
            else
            {
                responseBody.Columns = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}