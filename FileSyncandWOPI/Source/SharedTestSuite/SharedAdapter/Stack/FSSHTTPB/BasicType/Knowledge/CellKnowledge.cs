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
    using System.Collections.Generic;

    /// <summary>
    /// This class specifies the data element knowledge of the client. 
    /// </summary>
    public class CellKnowledge : SpecializedKnowledgeData
    {
        /// <summary>
        /// Initializes a new instance of the CellKnowledge class.
        /// </summary>
        public CellKnowledge()
            : base(StreamObjectTypeHeaderStart.CellKnowledge)
        {
            this.CellKnowledgeEntryList = new List<CellKnowledgeEntry>();
            this.CellKnowledgeRangeList = new List<CellKnowledgeRange>();
        }

        /// <summary>
        /// Gets or sets a list of cell knowledge ranges.
        /// </summary>
        public List<CellKnowledgeRange> CellKnowledgeRangeList { get; set; }

        /// <summary>
        /// Gets or sets a list of cell knowledge entries.
        /// </summary>
        public List<CellKnowledgeEntry> CellKnowledgeEntryList { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the cell knowledge from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the cell knowledge.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            if (lengthOfItems != 0)
            {
                throw new KnowledgeParseErrorException(currentIndex, "CellKnowledge object over-parse error", null);
            }

            int index = currentIndex;
            StreamObjectHeaderStart header;
            int length = 0;

            this.CellKnowledgeEntryList = new List<CellKnowledgeEntry>();
            this.CellKnowledgeRangeList = new List<CellKnowledgeRange>();
            while ((length = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                index += length;
                if (header.Type == StreamObjectTypeHeaderStart.CellKnowledgeEntry)
                {
                    this.CellKnowledgeEntryList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as CellKnowledgeEntry);
                }
                else if (header.Type == StreamObjectTypeHeaderStart.CellKnowledgeRange)
                {
                    this.CellKnowledgeRangeList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as CellKnowledgeRange);
                }
                else
                {
                    throw new KnowledgeParseErrorException(currentIndex, "Failed to parse CellKnowledge, expect the inner object type is either CellKnowledgeEntry or CellKnowledgeRange but actual type value is " + header.Type, null);
                }
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the cell knowledge to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of cell knowledge.</param>
        /// <returns>Return the length in byte of the items in cell knowledge.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            if (this.CellKnowledgeRangeList != null)
            {
                foreach (CellKnowledgeRange cellKnowledgeRange in this.CellKnowledgeRangeList)
                {
                    byteList.AddRange(cellKnowledgeRange.SerializeToByteList());
                }
            }

            if (this.CellKnowledgeEntryList != null)
            {
                foreach (CellKnowledgeEntry cellKnowledgeEntry in this.CellKnowledgeEntryList)
                {
                    byteList.AddRange(cellKnowledgeEntry.SerializeToByteList());
                }
            }

            return 0;
        }
    }

    /// <summary>
    /// This class specifies cell knowledge range of data elements.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class CellKnowledgeRange : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the CellKnowledgeRange class.
        /// </summary>
        public CellKnowledgeRange()
            : base(StreamObjectTypeHeaderStart.CellKnowledgeRange)
        {
            this.CellKnowledgeRangeGUID = Guid.NewGuid();
            this.From = new Compact64bitInt(0x0);
            this.To = new Compact64bitInt(0x5D);
        }
        
        /// <summary>
        /// Gets or sets a GUID (16 bytes) that specifies the data element.
        /// </summary>
        public Guid CellKnowledgeRangeGUID { get; set; }

        /// <summary>
        /// Gets or sets a compact unit64 that specifies the starting sequence number.
        /// </summary>
        public Compact64bitInt From { get; set; }

        /// <summary>
        /// Gets or sets a compact unit64 that specifies the ending sequence number.
        /// </summary>
        public Compact64bitInt To { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the cell knowledge range from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the cell knowledge range.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            byte[] temp = new byte[16];
            Array.Copy(byteArray, index, temp, 0, 16);
            this.CellKnowledgeRangeGUID = new Guid(temp);
            index += 16;

            this.From = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.To = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "CellKnowledgeRange", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the cell knowledge range to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of cell knowledge range.</param>
        /// <returns>Return the length in byte of the items in cell knowledge range.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.CellKnowledgeRangeGUID.ToByteArray());
            byteList.AddRange(this.From.SerializeToByteList());
            byteList.AddRange(this.To.SerializeToByteList());

            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// This class specifies cell knowledge entry.
    /// </summary>
    public class CellKnowledgeEntry : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the CellKnowledgeEntry class.
        /// </summary>
        public CellKnowledgeEntry()
            : base(StreamObjectTypeHeaderStart.CellKnowledgeEntry)
        {
            this.SerialNumber = new SerialNumber(System.Guid.NewGuid(), SequenceNumberGenerator.GetCurrentSerialNumber());
        }

        /// <summary>
        /// Gets or sets a serial number that specifies the cell.
        /// </summary>
        public SerialNumber SerialNumber { get; set; }

        /// <summary>
        /// This method is used to deserialize the items of the cell knowledge entry from the byte array.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="currentIndex">Specify the start index from the byte array.</param>
        /// <param name="lengthOfItems">Specify the current length of items in the cell knowledge entry.</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.SerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "CellKnowledgeEntry", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// This method is used to serialize the items of the cell knowledge entry to the byte list.
        /// </summary>
        /// <param name="byteList">Specify the byte list which stores the information of cell knowledge entry.</param>
        /// <returns>Return the length in byte of the items in cell knowledge entry.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.SerialNumber.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }
}