namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System.Collections.Generic;

    /// <summary>
    /// Revision Manifest data element 
    /// </summary>
    public class RevisionManifestDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the RevisionManifestDataElementData class.
        /// </summary>
        public RevisionManifestDataElementData()
        {
            this.RevisionManifest = new RevisionManifest();
            this.RevisionManifestRootDeclareList = new List<RevisionManifestRootDeclare>();
            this.RevisionManifestObjectGroupReferencesList = new List<RevisionManifestObjectGroupReferences>();
        }

        /// <summary>
        /// Gets or sets a 16-bit stream object header that specifies a revision manifest.
        /// </summary>
        public RevisionManifest RevisionManifest { get; set; }

        /// <summary>
        /// Gets or sets  a revision manifest root declare, each followed by root and object extended GUIDs.
        /// </summary>
        public List<RevisionManifestRootDeclare> RevisionManifestRootDeclareList { get; set; }

        /// <summary>
        /// Gets or sets  a list of revision manifest object group references, each followed by object group extended GUIDs.
        /// </summary>
        public List<RevisionManifestObjectGroupReferences> RevisionManifestObjectGroupReferencesList { get; set; }

        /// <summary>
        /// Used to return the length of this element.
        /// </summary>
        /// <param name="byteArray">A Byte list</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The length of the element</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            this.RevisionManifest = StreamObject.GetCurrent<RevisionManifest>(byteArray, ref index);

            this.RevisionManifestRootDeclareList = new List<RevisionManifestRootDeclare>();
            this.RevisionManifestObjectGroupReferencesList = new List<RevisionManifestObjectGroupReferences>();
            StreamObjectHeaderStart header;
            int headerLength = 0;
            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                if (header.Type == StreamObjectTypeHeaderStart.RevisionManifestRootDeclare)
                {
                    index += headerLength;
                    this.RevisionManifestRootDeclareList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as RevisionManifestRootDeclare);
                }
                else if (header.Type == StreamObjectTypeHeaderStart.RevisionManifestObjectGroupReferences)
                {
                    index += headerLength;
                    this.RevisionManifestObjectGroupReferencesList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as RevisionManifestObjectGroupReferences);
                }
                else
                {
                    throw new DataElementParseErrorException(index, "Failed to parse RevisionManifestDataElement, expect the inner object type RevisionManifestRootDeclare or RevisionManifestObjectGroupReferences, but actual type value is " + header.Type, null);
                }
            }

            return index - startIndex;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();
            byteList.AddRange(this.RevisionManifest.SerializeToByteList());

            if (this.RevisionManifestRootDeclareList != null)
            {
                foreach (RevisionManifestRootDeclare revisionManifestRootDeclare in this.RevisionManifestRootDeclareList)
                {
                    byteList.AddRange(revisionManifestRootDeclare.SerializeToByteList());
                }
            }

            if (this.RevisionManifestObjectGroupReferencesList != null)
            {
                foreach (RevisionManifestObjectGroupReferences revisionManifestObjectGroupReferences in this.RevisionManifestObjectGroupReferencesList)
                {
                    byteList.AddRange(revisionManifestObjectGroupReferences.SerializeToByteList());
                }
            }

            return byteList;
        }
    }

    /// <summary>
    /// Specifies a revision manifest.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class RevisionManifest : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the RevisionManifest class.
        /// </summary>
        public RevisionManifest()
            : base(StreamObjectTypeHeaderStart.RevisionManifest)
        {
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the revision identifier represented by this data element.
        /// </summary>
        public ExGuid RevisionID { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the revision identifier of a base revision that could contain additional information for this revision.
        /// </summary>
        public ExGuid BaseRevisionID { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.RevisionID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.BaseRevisionID = BasicObject.Parse<ExGuid>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "RevisionManifest", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The length of list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.RevisionID.SerializeToByteList());
            byteList.AddRange(this.BaseRevisionID.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Specifies a revision manifest root declare, each followed by root and object extended GUIDs.
    /// </summary>
    public class RevisionManifestRootDeclare : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the RevisionManifestRootDeclare class.
        /// </summary>
        public RevisionManifestRootDeclare()
            : base(StreamObjectTypeHeaderStart.RevisionManifestRootDeclare)
        {
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the root revision for each revision manifest root declare.
        /// </summary>
        public ExGuid RootExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the object for each revision manifest root declare.
        /// </summary>
        public ExGuid ObjectExtendedGUID { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte list</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.RootExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ObjectExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "RevisionManifestRootDeclare", "Stream object over-parse error", null);
            }

            currentIndex = index;
        }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <param name="byteList">A Byte list</param>
        /// <returns>The length of list</returns>
        protected override int SerializeItemsToByteList(List<byte> byteList)
        {
            int itemsIndex = byteList.Count;
            byteList.AddRange(this.RootExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.ObjectExtendedGUID.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Specifies a revision manifest object group references, each followed by object group extended GUIDs.
    /// </summary>
    public class RevisionManifestObjectGroupReferences : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the RevisionManifestObjectGroupReferences class.
        /// </summary>
        public RevisionManifestObjectGroupReferences()
            : base(StreamObjectTypeHeaderStart.RevisionManifestObjectGroupReferences)
        {
        }

        /// <summary>
        /// Initializes a new instance of the RevisionManifestObjectGroupReferences class.
        /// </summary>
        /// <param name="objectGroupExtendedGUID">Extended GUID</param>
        public RevisionManifestObjectGroupReferences(ExGuid objectGroupExtendedGUID)
            : this()
        {
            this.ObjectGroupExtendedGUID = objectGroupExtendedGUID;
        }

        /// <summary>
        /// Gets or sets an extended GUID that specifies the object group for each Revision Manifest Object Group References.
        /// </summary>
        public ExGuid ObjectGroupExtendedGUID { get; set; }

        /// <summary>
        /// Used to de-serialize the element.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ObjectGroupExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "RevisionManifestObjectGroupReferences", "Stream object over-parse error", null);
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
            List<byte> tmpList = this.ObjectGroupExtendedGUID.SerializeToByteList();
            byteList.AddRange(tmpList);
            return tmpList.Count;
        }
    }
}