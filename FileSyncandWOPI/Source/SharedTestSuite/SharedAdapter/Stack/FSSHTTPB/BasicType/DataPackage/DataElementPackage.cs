namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// A data element package contains the serialized file data elements made up of storage index, storage manifest, cell manifest, revision manifest, and object group or object data, or both.
    /// </summary>
    public class DataElementPackage : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DataElementPackage class.
        /// </summary>
        public DataElementPackage()
            : base(StreamObjectTypeHeaderStart.DataElementPackage)
        {
            this.DataElements = new List<DataElement>();
        }

        /// <summary>
        /// Gets or sets an optional array of data elements or data elements from hashes that specifies the serialized file data elements. If the client doesn’t have any data elements, this MUST NOT be present.
        /// </summary>
        public List<DataElement> DataElements { get; set; }

        /// <summary>
        /// Gets or sets a reserved field that MUST be set to zero, and MUST be ignored.
        /// </summary>
        public byte Reserved { get; set; }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>A constant value 1</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            // Add the reserved byte
            byteList.Add(0);

            foreach (DataElement dataElement in this.DataElements)
            {
                byteList.AddRange(dataElement.SerializeToByteList());
            }

            return 1;
        }    
     
        /// <summary>
        /// Used to de-serialize the elements.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">Length of the element</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 1)
            {
                throw new StreamObjectParseErrorException(currentIndex, "DataElementPackage", "Stream object over-parse error", null);
            }

            int index = currentIndex;

            this.Reserved = byteArray[index++];

            this.DataElements = new List<DataElement>();
            DataElement dataElement;
            while (StreamObject.TryGetCurrent<DataElement>(byteArray, ref index, out dataElement))
            {
                this.DataElements.Add(dataElement);
            }

            currentIndex = index;
        }
    }
}