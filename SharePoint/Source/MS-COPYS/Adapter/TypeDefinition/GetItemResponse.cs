//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    using System;

    /// <summary>
    /// A class used to store the response data of "GetItem" operation.
    /// </summary>
    public class GetItemResponse
    {   
        /// <summary>
        /// Represents the file binaries.
        /// </summary>
        private byte[] fileContentBinariesData;

        /// <summary>
        /// Represents the result status of "GetItem" operation.
        /// </summary>
        private uint getItemResultValue;

        /// <summary>
        /// Represents the fields information of a file.
        /// </summary>
        private FieldInformation[] fieldsCollection;

        /// <summary>
        /// Initializes a new instance of the GetItemResponse class.
        /// </summary>
        /// <param name="getItemResult">A parameter represents the result status of "CopyIntoItems" operation.</param>
        /// <param name="fields">A parameter represents the fields information of a file.</param>
        /// <param name="streamRawValue">A parameter represents the stream raw value.</param>
        public GetItemResponse(uint getItemResult, FieldInformation[] fields, byte[] streamRawValue)
        {
            this.getItemResultValue = getItemResult;
            this.fieldsCollection = fields;
            this.fileContentBinariesData = streamRawValue;
        }

        /// <summary>
        /// Gets the result status of the "GetItems" operation.
        /// </summary>
        public uint GetItemResult 
        {
            get
            {
                return this.getItemResultValue;
            }
        }

        /// <summary>
        /// Gets the field information of a file.
        /// </summary>
        public FieldInformation[] Fields 
        {
            get
            {
                return this.fieldsCollection;
            }
        }

        /// <summary>
        /// Gets the stream value, it is encoded by base64.
        /// </summary>
        public string Stream
        {
            get
            {
                if (null == this.fileContentBinariesData || 0 == this.fileContentBinariesData.Length)
                {
                    return string.Empty;
                }

                return Convert.ToBase64String(this.fileContentBinariesData);
            }
        }

        /// <summary>
        /// Gets the raw stream values, it represents the file content binaries.
        /// </summary>
        public byte[] StreamRawValues
        {
            get
            {
                return this.fileContentBinariesData;
            }
        }
    }
}