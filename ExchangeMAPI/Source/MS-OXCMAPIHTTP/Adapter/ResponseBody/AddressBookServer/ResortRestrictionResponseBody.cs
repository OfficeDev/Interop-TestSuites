namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A class indicates the response body of ResortRestriction request 
    /// </summary>
    public class ResortRestrictionResponseBody : AddressBookResponseBodyBase
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
        /// Gets or sets a value indicating whether the MinimalIdCount and MinimalIds fields are present.
        /// </summary>
        public bool HasMinimalIds { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the number of structures in the MinimalIds field.
        /// </summary>
        public uint? MinimalIdCount { get; set; }

        /// <summary>
        /// Gets or sets an array of MinimalEntryID structures that compose a restricted address book container.
        /// </summary>
        public uint[] MinimalIds { get; set; }

        /// <summary>
        /// Parse the ResortRestriction request type response body.
        /// </summary>
        /// <param name="rawData">The raw data of response.</param>
        /// <returns>The the ResortRestriction request type response body.</returns>
        public static ResortRestrictionResponseBody Parse(byte[] rawData)
        {
            ResortRestrictionResponseBody responseBody = new ResortRestrictionResponseBody();
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

            responseBody.HasMinimalIds = BitConverter.ToBoolean(rawData, index);
            index += sizeof(bool);

            List<uint> minimalIdsList = new List<uint>();
            if (responseBody.HasMinimalIds)
            {
                responseBody.MinimalIdCount = BitConverter.ToUInt32(rawData, index);
                index += sizeof(uint);
                for (int i = 0; i < responseBody.MinimalIdCount; i++)
                {
                    uint minId = BitConverter.ToUInt32(rawData, index);
                    minimalIdsList.Add(minId);
                    index += sizeof(uint);
                }

                responseBody.MinimalIds = minimalIdsList.ToArray();
            }
            else
            {
                responseBody.MinimalIdCount = null;
                minimalIdsList = null;
            }

            responseBody.AuxiliaryBufferSize = BitConverter.ToUInt32(rawData, index);
            index += 4;
            responseBody.AuxiliaryBuffer = new byte[responseBody.AuxiliaryBufferSize];
            Array.Copy(rawData, index, responseBody.AuxiliaryBuffer, 0, responseBody.AuxiliaryBufferSize);
            return responseBody;
        }
    }
}