namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the response body of GetProps request 
    /// </summary>
    public class GetTemplateInfoResponseBody : AddressBookResponseBodyBase
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
        /// Gets or sets a value indicating whether the Row field is present.
        /// </summary>
        public bool HasRow { get; set; }

        /// <summary>
        /// Gets or sets a AddressBookPropertyValueList structure that specifies the information that the client request.
        /// </summary>
        public AddressBookPropertyValueList? Row { get; set; }

        /// <summary>
        /// Parse the GetTemplateInfo request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The GetTemplateInfo request type response body.</returns>
        public static GetTemplateInfoResponseBody Parse(byte[] rawData)
        {
            GetTemplateInfoResponseBody responseBody = new GetTemplateInfoResponseBody();
            int index = 0;

            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.CodePage = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasRow = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasRow)
            {
                responseBody.Row = AddressBookPropertyValueList.Parse(rawData, ref index);
            }
            else
            {
                responseBody.Row = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}