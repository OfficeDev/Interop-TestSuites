namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// This class is used to tell the client which fragments of a large data element have been uploaded to the server.
    /// </summary>
    public class FragmentKnowledge : SpecializedKnowledgeData
    {
        /// <summary>
        /// Initializes a new instance of the FragmentKnowledge class.
        /// </summary>
        public FragmentKnowledge()
            : base(StreamObjectTypeHeaderStart.FragmentKnowledge)
        {
            this.FragmentKnowledgeEntriesList = new List<FragmentKnowledgeEntry>();
        }

        /// <summary>
        /// Gets or sets a list of fragment knowledge entry structures specifying the fragments which have been uploaded.
        /// </summary>
        public List<FragmentKnowledgeEntry> FragmentKnowledgeEntriesList { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the fragment knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the fragment knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "FragmentKnowledge object over-parse error", null);
            }

            int index = currentIndex;

            FragmentKnowledgeEntry fragmentKnowledgeEntry = null;
            while (StreamObject.TryGetCurrent<FragmentKnowledgeEntry>(byteArray, ref index, out fragmentKnowledgeEntry))
            {
                this.FragmentKnowledgeEntriesList.Add(fragmentKnowledgeEntry);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the fragment knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of the fragment knowledge.</param>
        /// <returns>Return the length in byte of the items in the fragment knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.FragmentKnowledgeEntriesList != null)
            {
                foreach (FragmentKnowledgeEntry fragmentKnowledgeEntry in this.FragmentKnowledgeEntriesList)
                {
                    byteList.AddRange(fragmentKnowledgeEntry.SerializeToByteList());
                }
            }

            return 0;
        }
    }

    /// <summary>
    /// This class is used to specify Fragment Knowledge Entry.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class FragmentKnowledgeEntry : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the FragmentKnowledgeEntry class.
        /// </summary>
        public FragmentKnowledgeEntry()
            : base(StreamObjectTypeHeaderStart.FragmentKnowledgeEntry)
        {
            this.ExtendedGUID = new ExGuid();
            this.DataElementSize = new Compact64bitInt();
            this.DataElementChunkReference = new FileChunk();
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the data element this fragment knowledge entry contains knowledge about.
        /// </summary>
        public ExGuid ExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a compact uint64 specifying the size in bytes of the data element specified by the preceding EXGUID.
        /// </summary>
        public Compact64bitInt DataElementSize { get; set; }

        /// <summary>
        /// Gets or sets a file chunk reference specifying which part of the data element with the preceding GUID this fragment knowledge entry contains knowledge about.
        /// </summary>
        public FileChunk DataElementChunkReference { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the fragment knowledge entry from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the fragment knowledge entry.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.DataElementSize = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.DataElementChunkReference = BasicObject.Parse<FileChunk>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "FragmentKnowledgeEntry", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the fragment knowledge entry to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of fragment knowledge entry.</param>
        /// <returns>Return the length in byte of the items in fragment knowledge entry.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;

            byteList.AddRange(this.ExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.DataElementSize.SerializeToByteList());
            byteList.AddRange(this.DataElementChunkReference.SerializeToByteList());

            return byteList.Count - itemsIndex;
        }
    }
}