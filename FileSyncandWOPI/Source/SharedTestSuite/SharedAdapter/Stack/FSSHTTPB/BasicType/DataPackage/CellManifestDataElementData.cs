namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// Cell manifest data element
    /// </summary>
    public class CellManifestDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the CellManifestDataElementData class.
        /// </summary>
        public CellManifestDataElementData()
        {
            this.CellManifestCurrentRevision = new CellManifestCurrentRevision();
        }

        /// <summary>
        /// Gets or sets a Cell Manifest Current Revision.
        /// </summary>
        public CellManifestCurrentRevision CellManifestCurrentRevision { get; set; }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The element length</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.CellManifestCurrentRevision = StreamObject.GetCurrent<CellManifestCurrentRevision>(byteArray, ref index);
            return index - startIndex;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>The Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            return this.CellManifestCurrentRevision.SerializeToByteList();
        }
    }

    /// <summary>
    /// Cell Manifest Current Revision
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class CellManifestCurrentRevision : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the CellManifestCurrentRevision class.
        /// </summary>
        public CellManifestCurrentRevision()
            : base(StreamObjectTypeHeaderStart.CellManifestCurrentRevision)
        {
        }

        /// <summary>
        /// Gets or sets a 16-bit stream object header that specifies a cell manifest current revision.
        /// </summary>
        public ExGuid CellManifestCurrentRevisionExtendedGUID { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.CellManifestCurrentRevisionExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "CellManifestCurrentRevision", "Stream object over-parse error", null);
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
            List<byte> tmpList = this.CellManifestCurrentRevisionExtendedGUID.SerializeToByteList();
            byteList.AddRange(tmpList);
            return tmpList.Count;
        }
    }
}