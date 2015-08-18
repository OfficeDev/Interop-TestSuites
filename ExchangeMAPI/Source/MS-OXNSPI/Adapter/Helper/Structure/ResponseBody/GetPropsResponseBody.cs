namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;

    /// <summary>
    /// A class indicates the response body of GetProps request 
    /// </summary>
    public class GetPropsResponseBody : AddressBookResponseBodyBase
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
        /// Gets or sets a value indicating whether the PropertyValues field is present.
        /// </summary>
        public bool HasPropertyValues { get; set; }

        /// <summary>
        /// Gets or sets a AddressBookPropValueList structure that contains the values of properties requested.
        /// </summary>
        public AddressBookPropValueList? PropertyValues { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of the request</returns>
        public static GetPropsResponseBody Parse(byte[] rawData)
        {
            GetPropsResponseBody responseBody = new GetPropsResponseBody();
            int index = 0;

            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.CodePage = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasPropertyValues = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasPropertyValues)
            {
                responseBody.PropertyValues = AddressBookPropValueList.Parse(rawData, ref index);
            }
            else
            {
                responseBody.PropertyValues = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}