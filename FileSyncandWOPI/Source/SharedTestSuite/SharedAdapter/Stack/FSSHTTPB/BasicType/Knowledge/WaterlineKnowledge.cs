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
    /// This class specifies the waterline knowledge.
    /// </summary>
    public class WaterlineKnowledge : SpecializedKnowledgeData
    {
        /// <summary>
        /// Initializes a new instance of the WaterlineKnowledge class.
        /// </summary>
        public WaterlineKnowledge()
            : base(StreamObjectTypeHeaderStart.WaterlineKnowledge)
        {
            this.WaterlineKnowledgeData = new List<WaterlineKnowledgeEntry>();
        }

        /// <summary>
        /// Gets or sets a list of waterline entries that specifies what the server has already delivered to the client or what the client has already received from the server.
        /// </summary>
        public List<WaterlineKnowledgeEntry> WaterlineKnowledgeData { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the waterline knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the waterline knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "WaterlineKnowledge object over-parse error", null);
            }

            int index = currentIndex;
            WaterlineKnowledgeEntry waterlineKnowledgeEntry;
            this.WaterlineKnowledgeData = new List<WaterlineKnowledgeEntry>();
            while (StreamObject.TryGetCurrent<WaterlineKnowledgeEntry>(byteArray, ref index, out waterlineKnowledgeEntry))
            {
                this.WaterlineKnowledgeData.Add(waterlineKnowledgeEntry);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the waterline knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of the waterline knowledge.</param>
        /// <returns>Return the length in byte of the items in the waterline knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.WaterlineKnowledgeData != null)
            {
                foreach (WaterlineKnowledgeEntry data in this.WaterlineKnowledgeData)
                {
                    byteList.AddRange(data.SerializeToByteList());
                }
            }

            return 0;
        }
    }

    /// <summary>
    /// This class specifies the waterline Knowledge Entry.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class WaterlineKnowledgeEntry : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the WaterlineKnowledgeEntry class.
        /// </summary>
        public WaterlineKnowledgeEntry()
            : base(StreamObjectTypeHeaderStart.WaterlineKnowledgeEntry)
        {
            this.CellStorageExtendedGUID = new ExGuid();
            this.Waterline = new Compact64bitInt();
            this.Reserved = new Compact64bitInt();
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the cell storage this entry specifies the waterline for.
        /// </summary>
        public ExGuid CellStorageExtendedGUID { get; set; }
        
        /// <summary>
        /// Gets or sets a compressed unit64 that specifies a sequential serial number.
        /// </summary>
        public Compact64bitInt Waterline { get; set; }

        /// <summary>
        /// Gets or sets a compressed unit64 that specifies a reserved field that MUST have value of 0 and MUST be ignored.
        /// </summary>
        public Compact64bitInt Reserved { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the waterline Knowledge Entry from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the waterline Knowledge Entry.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.CellStorageExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.Waterline = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.Reserved = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "WaterlineKnowledgeEntry", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the waterline Knowledge Entry to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of the waterline Knowledge Entry.</param>
        /// <returns>Return the length in byte of the items in the waterline Knowledge Entry.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;

            byteList.AddRange(this.CellStorageExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.Waterline.SerializeToByteList());
            byteList.AddRange(this.Reserved.SerializeToByteList());

            return byteList.Count - itemsIndex;
        }
    }
}