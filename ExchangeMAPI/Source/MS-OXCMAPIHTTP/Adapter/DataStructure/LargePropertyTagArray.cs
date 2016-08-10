namespace Microsoft.Protocols.TestSuites.MS_OXCMAPIHTTP
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools.Messages.Marshaling;

    /// <summary>
    /// The LargePropTagArray structure contains a list of property tags.
    /// </summary>
    public struct LargePropertyTagArray
    {
        /// <summary>
        /// The number of PropertyName_r structures in this aggregation. The value MUST NOT exceed 100,000.
        /// </summary>
        public uint PropertyTagCount;

        /// <summary>
        /// The list of PropertyName_r structures in this aggregation.
        /// </summary>
        [Size("propertyTagCount")]
        public PropertyTag[] PropertyTags;

        /// <summary>
        /// Parse the Large property tag array from the response data.
        /// </summary>
        /// <param name="rawData">The response data.</param>
        /// <param name="index">The start index of response data.</param>
        /// <returns>The result of parse the response data</returns>
        public static LargePropertyTagArray Parse(byte[] rawData, ref int index)
        {
            LargePropertyTagArray largePropTagArray = new LargePropertyTagArray();

            largePropTagArray.PropertyTagCount = BitConverter.ToUInt32(rawData, index);
            index += sizeof(uint);
            largePropTagArray.PropertyTags = new PropertyTag[largePropTagArray.PropertyTagCount];

            int count = 0;
            while (largePropTagArray.PropertyTagCount > count)
            {
                largePropTagArray.PropertyTags[count].PropertyType = BitConverter.ToUInt16(rawData, index);
                index += 2;
                largePropTagArray.PropertyTags[count].PropertyId = BitConverter.ToUInt16(rawData, index);
                index += 2;
                count++;
            }

            return largePropTagArray;
        }
    }
}