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
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies the content tag for each BLOB heap.
    /// </summary>
    public class ContentTagKnowledge : SpecializedKnowledgeData
    {
        /// <summary>
        /// Initializes a new instance of the ContentTagKnowledge class.
        /// </summary>
        public ContentTagKnowledge()
            : base(StreamObjectTypeHeaderStart.ContentTagKnowledge)
        {
            this.ContentTagEntryArray = new List<ContentTagKnowledgeEntry>();
        }

        /// <summary>
        /// Gets or sets a content Tag Entry Array that specifies the BLOB heap entries.
        /// </summary>
        public List<ContentTagKnowledgeEntry> ContentTagEntryArray { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the content tag knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the content tag knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "ContentTagKnowledge object over-parse error", null);
            }

            int index = currentIndex;

            this.ContentTagEntryArray = new List<ContentTagKnowledgeEntry>();
            ContentTagKnowledgeEntry outValue;
            while (StreamObject.TryGetCurrent<ContentTagKnowledgeEntry>(byteArray, ref index, out outValue))
            {
                this.ContentTagEntryArray.Add(outValue);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the content tag knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of content tag knowledge.</param>
        /// <returns>Return the length in byte of the items in content tag knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.ContentTagEntryArray != null)
            {
                foreach (ContentTagKnowledgeEntry contentTagKnowledgeEntry in this.ContentTagEntryArray)
                {
                    byteList.AddRange(contentTagKnowledgeEntry.SerializeToByteList());
                }
            }

            return 0;
        }
    }

    /// <summary>
    /// This class specifies a BLOB heap GUID and its associated content tag.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ContentTagKnowledgeEntry : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the ContentTagKnowledgeEntry class.
        /// </summary>
        public ContentTagKnowledgeEntry()
            : base(StreamObjectTypeHeaderStart.ContentTagKnowledgeEntry)
        {
            this.BLOBHeapExtendedGUID = new ExGuid();
            this.ClockData = new BinaryItem();
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the BLOB heap which this content tag is for.
        /// </summary>
        public ExGuid BLOBHeapExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a binary item that specifies changes when the contents of the BLOB heap change on the server.
        /// </summary>
        public BinaryItem ClockData { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the content tag knowledge entry from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the content tag knowledge entry.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.BLOBHeapExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ClockData = BasicObject.Parse<BinaryItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "ContentTagKnowledgeEntry", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the content tag knowledge entry to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of content tag knowledge entry.</param>
        /// <returns>Return the length in byte of the items in content tag knowledge entry.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.BLOBHeapExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.ClockData.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }
}