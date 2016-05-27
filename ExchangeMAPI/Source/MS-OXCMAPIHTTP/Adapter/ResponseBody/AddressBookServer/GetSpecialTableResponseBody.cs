namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the response body of GetSpecialTable request 
    /// </summary>
    public class GetSpecialTableResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the code page that the server used to express string properties.
        /// </summary>
        public uint CodePage { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the Version field is present.
        /// </summary>
        public bool HasVersion { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the version number of the address book hierarchy table that the server has.
        /// </summary>
        public uint? Version { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the PropertyValues field is present.
        /// </summary>
        public bool HasRows { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures in the Rows field.
        /// </summary>
        public uint? RowCount { get; set; }

        /// <summary>
        /// Gets or sets a AddressBookPropertyValueList structure that contains the values of properties requested.
        /// </summary>
        public AddressBookPropertyValueList[] Rows { get; set; }

        /// <summary>
        /// Parse the GetSpecialTable request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The GetSpecialTable request type response body.</returns>
        public static GetSpecialTableResponseBody Parse(byte[] rawData)
        {
            GetSpecialTableResponseBody responseBody = new GetSpecialTableResponseBody();
            int index = 0;

            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.CodePage = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasVersion = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasVersion)
            {
                responseBody.Version = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
            }
            else
            {
                responseBody.Version = null;
            }

            responseBody.HasRows = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasRows)
            {
                responseBody.RowCount = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
                responseBody.Rows = new AddressBookPropertyValueList[(uint)responseBody.RowCount];
                for (int i = 0; i < responseBody.RowCount; i++)
                {
                    responseBody.Rows[i] = AddressBookPropertyValueList.Parse(rawData, ref index);
                }
            }
            else
            {
                responseBody.RowCount = null;
                responseBody.Rows = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}