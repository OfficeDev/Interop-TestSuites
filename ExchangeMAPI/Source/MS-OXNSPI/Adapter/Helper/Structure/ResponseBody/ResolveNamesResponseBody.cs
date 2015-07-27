//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;

    /// <summary>
    /// A class indicates the response body of ResolveNames request 
    /// </summary>
    public class ResolveNamesResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the code page of the operation.
        /// </summary>
        public uint CodePage { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the MinimalIdCount and MinimalIds fields are present. 
        /// </summary>
        public bool HasMinimalIds { get; set; }

        /// <summary>
        /// Gets or sets an integer that specifies the number of structures present in the MinimalIds field.
        /// </summary>
        public uint? MinimalIdCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures, each of which is the ID of an object found.
        /// </summary>
        public uint[] MinimalIds { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the PropertyTags, RowCount and RowData fields are present.
        /// </summary>
        public bool HasRowsAndPropertyTags { get; set; }

        /// <summary>
        /// Gets or sets a LargePropTagArray structure that specifies the properties returned for the rows in the RowData field.
        /// </summary>
        public LargePropTagArray? PropertyTags { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures contained in the RowData field.
        /// </summary>
        public uint? RowCount { get; set; }

        /// <summary>
        /// Gets or sets an array of AddressBookPropertyRow structures, each of which specifies the row data for the entries queried.
        /// </summary>
        public AddressBookPropertyRow[] RowData { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of the request</returns>
        public static ResolveNamesResponseBody Parse(byte[] rawData)
        {
            ResolveNamesResponseBody responseBody = new ResolveNamesResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.CodePage = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);

            responseBody.HasMinimalIds = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasMinimalIds)
            {
                responseBody.MinimalIdCount = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
                responseBody.MinimalIds = new uint[(uint)responseBody.MinimalIdCount];
                for (int i = 0; i < responseBody.MinimalIdCount; i++)
                {
                    responseBody.MinimalIds[i] = BitConverter.ToUInt32(rawData, index);
                    index += sizeof(uint);
                }
            }
            else
            {
                responseBody.MinimalIdCount = null;
                responseBody.MinimalIds = null;
            }

            responseBody.HasRowsAndPropertyTags = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasRowsAndPropertyTags)
            {
                responseBody.PropertyTags = LargePropTagArray.Parse(rawData, ref index);
                responseBody.RowCount = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
                responseBody.RowData = new AddressBookPropertyRow[(uint)responseBody.RowCount];
                for (int i = 0; i < responseBody.RowCount; i++)
                {
                    responseBody.RowData[i] = AddressBookPropertyRow.Parse(rawData, (LargePropTagArray)responseBody.PropertyTags, ref index);
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
