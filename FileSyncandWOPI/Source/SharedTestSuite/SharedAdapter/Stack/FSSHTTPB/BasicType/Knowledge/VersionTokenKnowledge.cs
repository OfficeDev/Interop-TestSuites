namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class contains version token knowledge identifying the version of the file.
    /// </summary>
    public class VersionTokenKnowledge : SpecializedKnowledgeData
    {
        /// <summary>
        /// Initializes a new instance of the VersionTokenKnowledge class.
        /// </summary>
        public VersionTokenKnowledge()
            : base(StreamObjectTypeHeaderStart.VersionTokenKnowledge)
        {
            this.TokenData = new BinaryItem();
        }

        /// <summary>
        /// Gets or sets a binary item that specifies version token.
        /// </summary>
        public BinaryItem TokenData { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the version token knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the version token knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "VersionTokenKnowledge object over-parse error", null);
            }

            int index = currentIndex;
            this.TokenData = BasicObject.Parse<BinaryItem>(byteArray, ref index);
            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the content version token to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of version token knowledge.</param>
        /// <returns>Return the length in byte of the items in version token knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.TokenData.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }
}
