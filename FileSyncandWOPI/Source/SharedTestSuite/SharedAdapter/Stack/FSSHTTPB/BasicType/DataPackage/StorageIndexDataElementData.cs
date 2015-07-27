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
    /// Storage Index data element
    /// </summary>
    public class StorageIndexDataElementData : DataElementData
    {
        /// <summary>
        /// Initializes a new instance of the StorageIndexDataElementData class.
        /// </summary>
        public StorageIndexDataElementData()
        {
            this.StorageIndexManifestMapping = new StorageIndexManifestMapping();
            this.StorageIndexCellMappingList = new List<StorageIndexCellMapping>();
            this.StorageIndexRevisionMappingList = new List<StorageIndexRevisionMapping>();
        }

        /// <summary>
        /// Gets or sets the storage index manifest mappings (with manifest mapping extended GUID and serial number).
        /// </summary>
        public StorageIndexManifestMapping StorageIndexManifestMapping { get; set; }

        /// <summary>
        /// Gets or sets  storage index manifest mappings.
        /// </summary>
        public List<StorageIndexCellMapping> StorageIndexCellMappingList { get; set; }

        /// <summary>
        /// Gets or sets the list of storage index revision mappings (with revision and revision mapping extended GUIDs, and revision mapping serial number).
        /// </summary>
        public List<StorageIndexRevisionMapping> StorageIndexRevisionMappingList { get; set; }

        /// <summary>
        /// Used to convert the element into a byte List.
        /// </summary>
        /// <returns>A Byte list</returns>
        public override List<byte> SerializeToByteList()
        {
            List<byte> byteList = new List<byte>();

            if (this.StorageIndexManifestMapping != null)
            {
                byteList.AddRange(this.StorageIndexManifestMapping.SerializeToByteList());
            }
            
            if (this.StorageIndexCellMappingList != null)
            {
                foreach (StorageIndexCellMapping cellMapping in this.StorageIndexCellMappingList)
                {
                    byteList.AddRange(cellMapping.SerializeToByteList());
                }
            }

            // Storage Index Revision Mapping 
            if (this.StorageIndexRevisionMappingList != null)
            {
                foreach (StorageIndexRevisionMapping revisionMapping in this.StorageIndexRevisionMappingList)
                {
                    byteList.AddRange(revisionMapping.SerializeToByteList());
                }
            }

            return byteList;
        }

        /// <summary>
        /// Used to de-serialize the data element.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="startIndex">Start position</param>
        /// <returns>The length of the element</returns>
        public override int DeserializeDataElementDataFromByteArray(byte[] byteArray, int startIndex)
        {
            int index = startIndex;
            int headerLength = 0;
            StreamObjectHeaderStart header;
            bool isStorageIndexManifestMappingExist = false;
            while ((headerLength = StreamObjectHeaderStart.TryParse(byteArray, index, out header)) != 0)
            {
                index += headerLength;
                if (header.Type == StreamObjectTypeHeaderStart.StorageIndexManifestMapping)
                {
                    if (isStorageIndexManifestMappingExist)
                    {
                        throw new DataElementParseErrorException(index - headerLength, "Failed to parse StorageIndexDataElement, only can contain zero or one StorageIndexManifestMapping", null);
                    }

                    this.StorageIndexManifestMapping = StreamObject.ParseStreamObject(header, byteArray, ref index) as StorageIndexManifestMapping;
                    isStorageIndexManifestMappingExist = true;
                }
                else if (header.Type == StreamObjectTypeHeaderStart.StorageIndexCellMapping)
                {
                    this.StorageIndexCellMappingList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as StorageIndexCellMapping);
                }
                else if (header.Type == StreamObjectTypeHeaderStart.StorageIndexRevisionMapping)
                {
                    this.StorageIndexRevisionMappingList.Add(StreamObject.ParseStreamObject(header, byteArray, ref index) as StorageIndexRevisionMapping);
                }
                else
                {
                    throw new DataElementParseErrorException(index - headerLength, "Failed to parse StorageIndexDataElement, expect the inner object type StorageIndexCellMapping or StorageIndexRevisionMapping, but actual type value is " + header.Type, null);
                }
            }

            return index - startIndex;
        }
    }

    /// <summary>
    /// Specifies the storage index manifest mappings (with manifest mapping extended GUID and serial number).
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class StorageIndexManifestMapping : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the StorageIndexManifestMapping class.
        /// </summary>
        public StorageIndexManifestMapping()
            : base(StreamObjectTypeHeaderStart.StorageIndexManifestMapping)
        {
        }

        /// <summary>
        /// Gets or sets the extended GUID of the manifest mapping.
        /// </summary>
        public ExGuid ManifestMappingExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets the serial number of the manifest mapping.
        /// </summary>
        public SerialNumber ManifestMappingSerialNumber { get; set; }

        /// <summary>
        /// Used to Deserialize the items.
        /// </summary>
        /// <param name="byteArray">Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.ManifestMappingExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.ManifestMappingSerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "StorageIndexManifestMapping", "Stream object over-parse error", null);
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
            byteList.AddRange(this.ManifestMappingExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.ManifestMappingSerialNumber.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Specifies the storage index cell mappings (with cell identifier, cell mapping extended GUID, and cell mapping serial number).
    /// </summary>
    public class StorageIndexCellMapping : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the StorageIndexCellMapping class.
        /// </summary>
        public StorageIndexCellMapping()
            : base(StreamObjectTypeHeaderStart.StorageIndexCellMapping)
        {
        }

        /// <summary>
        /// Gets or sets the cell identifier.
        /// </summary>
        public CellID CellID { get; set; }

        /// <summary>
        /// Gets or sets the extended GUID of the cell mapping.
        /// </summary>
        public ExGuid CellMappingExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets the serial number of the cell mapping.
        /// </summary>
        public SerialNumber CellMappingSerialNumber { get; set; }

        /// <summary>
        /// Used to de-serialize the items.
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.CellID = BasicObject.Parse<CellID>(byteArray, ref index);
            this.CellMappingExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.CellMappingSerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "StorageIndexCellMapping", "Stream object over-parse error", null);
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
            byteList.AddRange(this.CellID.SerializeToByteList());
            byteList.AddRange(this.CellMappingExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.CellMappingSerialNumber.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }

    /// <summary>
    /// Specifies the storage index revision mappings (with revision and revision mapping extended GUIDs, and revision mapping serial number).
    /// </summary>
    public class StorageIndexRevisionMapping : StreamObject
    {
        /// <summary>
        /// Initializes a new instance of the StorageIndexRevisionMapping class.
        /// </summary>
        public StorageIndexRevisionMapping()
            : base(StreamObjectTypeHeaderStart.StorageIndexRevisionMapping)
        {
        }

        /// <summary>
        /// Gets or sets the extended GUID of the revision.
        /// </summary>
        public ExGuid RevisionExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets the extended GUID of the revision mapping.
        /// </summary>
        public ExGuid RevisionMappingExtendedGUID { get; set; }

        /// <summary>
        /// Gets or sets the serial number of the revision mapping.
        /// </summary>
        public SerialNumber RevisionMappingSerialNumber { get; set; }

        /// <summary>
        /// Used to de-serialize the items
        /// </summary>
        /// <param name="byteArray">A Byte array</param>
        /// <param name="currentIndex">Start position</param>
        /// <param name="lengthOfItems">The length of the items</param>
        protected override void DeserializeItemsFromByteArray(byte[] byteArray, ref int currentIndex, int lengthOfItems)
        {
            int index = currentIndex;
            this.RevisionExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.RevisionMappingExtendedGUID = BasicObject.Parse<ExGuid>(byteArray, ref index);
            this.RevisionMappingSerialNumber = BasicObject.Parse<SerialNumber>(byteArray, ref index);

            if (index - currentIndex != lengthOfItems)
            {
                throw new StreamObjectParseErrorException(currentIndex, "StorageIndexRevisionMapping", "Stream object over-parse error", null);
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
            byteList.AddRange(this.RevisionExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.RevisionMappingExtendedGUID.SerializeToByteList());
            byteList.AddRange(this.RevisionMappingSerialNumber.SerializeToByteList());
            return byteList.Count - itemsIndex;
        }
    }
}