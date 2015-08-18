namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;

    /// <summary>
    /// A class indicates the response body of DNToMinId request 
    /// </summary>
    public class DnToMinIdResponseBody : AddressBookResponseBodyBase
    {
        /// <summary>
        /// Gets or sets an unsigned integer that specifies the return status of the operation.
        /// </summary>
        public uint ErrorCode { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the MinimalIdCount and MinimalIds fields are present.
        /// </summary>
        public bool HasMinimalIds { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures in the MinimalIds field.
        /// </summary>
        public uint? MinimalIdCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures, each of which specifies a Minimal Entry ID that matches a requested distinguished name.
        /// </summary>
        public uint[] MinimalIds { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of the request</returns>
        public static DnToMinIdResponseBody Parse(byte[] rawData)
        {
            DnToMinIdResponseBody responseBody = new DnToMinIdResponseBody();
            int index = 0;
            responseBody.StatusCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.ErrorCode = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            responseBody.HasMinimalIds = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasMinimalIds)
            {
                responseBody.MinimalIdCount = BitConverter.ToUInt32(rawData, index);
                responseBody.MinimalIds = new uint[(uint)responseBody.MinimalIdCount];
                index += sizeof(uint);
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

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}