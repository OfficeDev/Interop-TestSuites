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
    using System.Text;

    /// <summary>
    /// A class indicates the response body of GetMailboxUrl request 
    /// </summary>
    public class GetAddressBookUrlResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a null terminated Unicode string that specifies URL of the address book server.
        /// </summary>
        public string ServerUrl { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of the request</returns>
        public static GetAddressBookUrlResponseBody Parse(byte[] rawData)
        {
            GetAddressBookUrlResponseBody responseBody = new GetAddressBookUrlResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += 4;

            // The length in bytes of the Unicode string to parse
            int strBytesLen = 0;

            // Find the string with '\0''\0' end
            for (int i = index; i < rawData.Length; i += 2)
            {
                strBytesLen += 2;
                if ((rawData[i] == 0) && (rawData[i + 1] == 0))
                {
                    break;
                }
            }

            byte[] serverUrlBuffer = new byte[strBytesLen];
            Array.Copy(rawData, index, serverUrlBuffer, 0, strBytesLen);
            index += strBytesLen;
            responseBody.ServerUrl = Encoding.Unicode.GetString(serverUrlBuffer);
            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}
