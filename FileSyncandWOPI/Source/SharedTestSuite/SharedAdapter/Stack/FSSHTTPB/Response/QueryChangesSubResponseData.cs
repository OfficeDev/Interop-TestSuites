//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    /// <summary>
    /// Gives back the set of changes the server has for the data elements
    /// </summary>
    public class QueryChangesSubResponseData : SubResponseData
    {
        /// <summary>
        /// Initializes a new instance of the QueryChangesSubResponseData class. 
        /// </summary>
        public QueryChangesSubResponseData()
        {
        }

        /// <summary>
        /// Gets or sets Query Changes Response (4 bytes): A 32-bit stream object header that specifies a query changes response.
        /// </summary>
        public StreamObjectHeaderStart32bit QueryChangesResponseStart { get; set; }

        /// <summary>
        /// Gets or sets Storage Index Extended GUID (variable): An extended GUID that specifies storage index.
        /// </summary>
        public ExGuid StorageIndexExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether the result is a partial result.
        /// </summary>
        public bool PartialResult { get; set; }

        /// <summary>
        /// Gets or sets Reserved (7 bits): A 7-bit reserved and MUST be set to 0 and MUST be ignored.
        /// </summary>
        public byte ReservedQueryChanges { get; set; }

        /// <summary>
        /// Gets or sets Knowledge (variable): A knowledge that specifies the current state of the file on the server.
        /// </summary>
        public Knowledge Knowledge { get; set; }

        /// <summary>
        /// Deserialize sub response data from byte array.
        /// </summary>
        /// <param name="byteArray">The byte array which contains sub response data.</param>
        /// <param name="currentIndex">The index special where to start.</param>
        protected override void DeserializeSubResponseDataFromByteArray(byte[] byteArray, ref int currentIndex)
        {
            int index = currentIndex;
            int headerLength = 0;
            StreamObjectHeaderStart header;
            if ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) == 0)
            {
                throw new ResponseParseErrorException(index, "Failed to parse the QueryChangesData stream object header", null);
            }

            if (header.Type != StreamObjectTypeHeaderStart.QueryChangesResponse)
            {
                throw new ResponseParseErrorException(index, "Failed to extract the QueryChangesData stream object header type, unexpected value " + header.Type, null);
            }

            index += headerLength;
            this.QueryChangesResponseStart = header as StreamObjectHeaderStart32bit;

            int currentTmpIndex = index;
            this.StorageIndexExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.PartialResult = (byteArray[index] & 0x1) == 0x1 ? true : false;
            this.ReservedQueryChanges = (byte)(byteArray[index] >> 1);
            index += 1;

            if (index - currentTmpIndex != header.Length)
            {
                throw new ResponseParseErrorException(-1, "QueryChangesData object over-parse error", null);
            }

            this.Knowledge = StreamObject.GetCurrent<Knowledge>(byteArray, ref index);
            currentIndex = index;
        }
    }
}