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
    using System;

    /// <summary>
    /// This class is used to represent the allocate extended guid range response data.
    /// </summary>
    public class AllocateExtendedGuidRangeSubResponseData : SubResponseData
    {
        /// <summary>
        /// Initializes a new instance of the AllocateExtendedGuidRangeSubResponseData class. 
        /// </summary>
        public AllocateExtendedGuidRangeSubResponseData()
        {
        }

        /// <summary>
        /// Gets or sets Allocate ExtendedGuid Range Response (4 bytes): A stream object header that specifies an allocate extendedGUID range response.
        /// </summary>
        public StreamObjectHeaderStart32bit AllocateExtendedGUIDRangeResponse { get; set; }

        /// <summary>
        /// Gets or sets GUID Component (16 bytes): A GUID that specifies the GUID portion of the reserved extended GUIDs.
        /// </summary>
        public Guid GUIDComponent { get; set; }

        /// <summary>
        /// Gets or sets Integer Range Min (variable): A compact unsigned 64-bit integer that specifies the first integer element in the range of extended GUIDs.
        /// </summary>
        public Compact64bitInt IntegerRangeMin { get; set; }

        /// <summary>
        /// Gets or sets Integer Range Max (variable): A compact unsigned 64-bit integer that specifies the last + 1 integer element in the range of extended GUIDs.
        /// </summary>
        public Compact64bitInt IntegerRangeMax { get; set; }

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
                throw new ResponseParseErrorException(index, "Failed to parse the AllocateExtendedGuidRangeData stream object header", null);
            }

            if (header.Type != StreamObjectTypeHeaderStart.AllocateExtendedGUIDRangeResponse)
            {
                throw new ResponseParseErrorException(index, "Failed to extract the AllocateExtendedGuidRangeData stream object header type, unexpected value " + header.Type, null);
            }

            this.AllocateExtendedGUIDRangeResponse = new StreamObjectHeaderStart32bit(header.Type, header.Length);
            index += headerLength;
            int currentTmpIndex = index;
            byte[] guidarray = new byte[16];
            Array.Copy(byteArray, index, guidarray, 0, 16);
            this.GUIDComponent = new Guid(guidarray);
            index += 16;
            this.IntegerRangeMin = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.IntegerRangeMax = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentTmpIndex != header.Length)
            {
                throw new ResponseParseErrorException(-1, "AllocateExtendedGuidRangeData object over-parse error", null);
            }

            currentIndex = index;
        }
    }
}