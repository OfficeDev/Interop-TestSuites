namespace Microsoft.Protocols.TestSuites.MS_OXNSPI
{
    using System;

    /// <summary>
    /// A class indicates the response body of Update STAT request
    /// </summary>
    public class UpdateStatResponseBody : AddressBookResponseBodyBase
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
        /// Gets or sets a value indicating whether the Delta field is present.
        /// </summary>
        public bool HasDelta { get; set; }

        /// <summary>
        /// Gets or sets a signed integer that specifies the movement within the address book container that was specified in the State field of the request.
        /// </summary>
        public int? Delta { get; set; }

        /// <summary>
        /// Parse the response data into response body.
        /// </summary>
        /// <param name="rawData">The raw data of response</param>
        /// <returns>The response body of the request</returns>
        public static UpdateStatResponseBody Parse(byte[] rawData)
        {
            UpdateStatResponseBody responseBody = new UpdateStatResponseBody();
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
            else
            {
                responseBody.State = null;
            }

            responseBody.HasDelta = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);
            if (responseBody.HasDelta)
            {
                responseBody.Delta = BitConverter.ToInt32(rawData, index);
                index += sizeof(int);
            }
            else
            {
                responseBody.Delta = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}