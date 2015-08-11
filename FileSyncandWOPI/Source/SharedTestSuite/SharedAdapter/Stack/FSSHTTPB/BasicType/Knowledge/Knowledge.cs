namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies the client knows about a state of a file.
    /// </summary>
    public class Knowledge : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the Knowledge class.
        /// </summary>
        public Knowledge()
            : base(StreamObjectTypeHeaderStart.Knowledge)
        {
            this.SpecializedKnowledges = new List<SpecializedKnowledge>();
        }

        /// <summary>
        /// Gets or sets a list of specialized knowledge structures.
        /// </summary>
        public List<SpecializedKnowledge> SpecializedKnowledges { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "Knowledge object over-parse error", null);
            }

            int index = currentIndex;

            SpecializedKnowledge specializedKnowledge;
            this.SpecializedKnowledges = new List<SpecializedKnowledge>();
            while (StreamObject.TryGetCurrent<SpecializedKnowledge>(byteArray, ref index, out specializedKnowledge))
            {
                this.SpecializedKnowledges.Add(specializedKnowledge);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of the knowledge.</param>
        /// <returns>Return the length in byte of the items in the knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.SpecializedKnowledges != null)
            {
                foreach (SpecializedKnowledge specializedKnowledge in this.SpecializedKnowledges)
                {
                    byteList.AddRange(specializedKnowledge.SerializeToByteList());
                }
            }

            return 0;
        }
    }
}