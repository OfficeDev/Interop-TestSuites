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
    /// This class specifies the data part in the data element fragment structure.
    /// </summary>
    public class FragmentDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the FragmentDataElementData class.
        /// </summary>
        public FragmentDataElementData()
        {
            this.DataElementFragment = new DataElementFragment();
        }

        /// <summary>
        /// Gets or sets a 32-bit stream object header (section 2.2.1.5.2) that specifies a data element fragment.
        /// </summary>
        public DataElementFragment DataElementFragment { get; set; }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The element length</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.DataElementFragment = StreamObject.GetCurrent<DataElementFragment>(byteArray, ref index);
            return index - startIndex;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte List</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.DataElementFragment.SerializeToByteList();
        }
    }

    /// <summary>
    /// Specifies a data element fragment.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class DataElementFragment : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the DataElementFragment class.
        /// </summary>
        public DataElementFragment()
            : base(StreamObjectTypeHeaderStart.DataElementFragment)
        {
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the data element fragment.
        /// </summary>
        public ExGuid FragmentExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the size in bytes of the fragmented data element.
        /// </summary>
        public Compact64bitInt FragmentDataElementSize { get; set; }

        /// <summary>
        /// Gets or sets a file chunk reference that specifies the data element fragment.
        /// </summary>
        public FileChunk FragmentFileChunkReference { get; set; }

        /// <summary>
        /// Gets or sets a byte stream that specifies the binary data opaque to this protocol.
        /// </summary>
        public BinaryItem FragmentData { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;

            this.FragmentExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.FragmentDataElementSize = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);
            this.FragmentFileChunkReference = BasicObject.Parse<FileChunk>(byteArray, ref index);
            this.FragmentData = BasicObject.Parse<BinaryItem>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "DataElementFragment", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of elements</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int startItemsLength = byteList.Count;
            byteList.AddRange(this.FragmentExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.FragmentDataElementSize.SerializeToByteList());
            byteList.AddRange(this.FragmentFileChunkReference.SerializeToByteList());
            byteList.AddRange(this.FragmentData.SerializeToByteList());

            return byteList.Count - startItemsLength;
        }
    }
}