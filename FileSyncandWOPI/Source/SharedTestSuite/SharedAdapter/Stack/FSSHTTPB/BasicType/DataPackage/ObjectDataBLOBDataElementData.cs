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
    /// Object Data BLOB data element
    /// </summary>
    public class ObjectDataBLOBDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the ObjectDataBLOBDataElementData class.
        /// </summary>
        public ObjectDataBLOBDataElementData()
            : base()
        {
            this.ObjectDataBLOB = new ObjectDataBLOB();
        }

        /// <summary>
        /// Gets or sets a 16-bit or 32-bit stream object header that specifies an object data BLOB.
        /// </summary>
        public ObjectDataBLOB ObjectDataBLOB { get; set; }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The length of the element</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.ObjectDataBLOB = StreamObject.GetCurrent<ObjectDataBLOB>(byteArray, ref index);
            return index - startIndex;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.ObjectDataBLOB.SerializeToByteList();
        }
    }

    /// <summary>
    /// A 16-bit or 32-bit stream object header that specifies an object data BLOB. 
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ObjectDataBLOB : StreamObject
    {
        /// <summary>
        /// A byte stream that specifies the binary data opaque to this protocol.
        /// </summary>
        private List<byte> data = new List<byte>();

        /// <summary>
        /// Initializes a new instance of the ObjectDataBLOB class.
        /// </summary>
        public ObjectDataBLOB()
            : base(StreamObjectTypeHeaderStart.ObjectDataBLOB)
        {
        }

        /// <summary>
        /// Gets or sets a byte stream that specifies the binary data opaque to this protocol.
        /// </summary>
        public List<byte> Data
        {
            get { return this.data; }
            set { this.data = value; }
        }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            this.data = new List<byte>();

            int index = currentIndex;
            for (; index - currentIndex < lengthOfItems; index++)
            {
                this.data.Add(byteArray[index]);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The number of elements actually contained in the list.</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            List<byte> tmpList = this.data;
            byteList.AddRange(tmpList);
            return tmpList.Count;
        }
    }
}