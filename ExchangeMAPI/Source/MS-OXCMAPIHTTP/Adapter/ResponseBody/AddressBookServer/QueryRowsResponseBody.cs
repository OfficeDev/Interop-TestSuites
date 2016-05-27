namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the response body of QueryRows request 
    /// </summary>
    public class QueryRowsResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the State field is present.
        /// </summary>
        public bool HasState { get; set; }

        /// <summary>
        /// Gets or sets a STAT structure that specifies the state of a specific address book container.
        /// </summary>
        public STAT? State { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Columns, RowCount and RowData fields are present.
        /// </summary>
        public bool HasColumnsAndRows { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the columns used for the rows returned.
        /// </summary>
        public LargePropertyTagArray? Columns { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures contained in the RowData field.
        /// </summary>
        public uint? RowCount { get; set; }

        /// <summary>
        /// Gets or sets an array of AddressBookPropertyRow structures, each of which specifies the row data for the entries queried.
        /// </summary>
        public AddressBookPropertyRow[] RowData { get; set; }

        /// <summary>
        /// Parse the QueryRows request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The QueryRows request type response body.</returns>
        public static QueryRowsResponseBody Parse(byte[] rawData)
        {
            QueryRowsResponseBody responseBody = new QueryRowsResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasState = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasState)
            {
                responseBody.State = STAT.Parse(rawData, ref index);
            }

            responseBody.HasColumnsAndRows = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasColumnsAndRows)
            {
                responseBody.Columns = LargePropertyTagArray.Parse(rawData, ref index);
                responseBody.RowCount = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
                responseBody.RowData = new AddressBookPropertyRow[(uint)responseBody.RowCount];
                for (int i = 0; i < responseBody.RowCount; i++)
                {
                    responseBody.RowData[i] = AddressBookPropertyRow.Parse(rawData, (LargePropertyTagArray)responseBody.Columns, ref index);
                }
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}